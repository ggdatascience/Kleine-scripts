library(tidyverse)
library(ggdnog)
library(labelled)
library(haven)
library(zoo)
library(openxlsx)
library(cbsodataR)
library(this.path)

setwd(dirname(this.path()))

recode_values = function (data, dictionary) {
  old = names(dictionary)
  new = unname(dictionary)
  output = rep(NA, length(data))
  for (i in 1:length(dictionary)) {
    output[data == old[i]] = new[i]
  }
  return(output)
}

# wijkindeling stelen van het CBS - wijkindeling staat in de kerncijfers wijken en buurten 2023
wijkindeling = cbs_get_meta("85618NED")
gemeenten = wijkindeling$WijkenEnBuurten[str_starts(wijkindeling$WijkenEnBuurten$Key, "GM"),] %>%
  rename(Gemeentenaam=Title)
wijkindeling = wijkindeling$WijkenEnBuurten[str_starts(wijkindeling$WijkenEnBuurten$Key, "WK"),] %>%
  rename(Wijknaam=Title, Wijkcode=DetailRegionCode) %>%
  left_join(gemeenten %>% select(Gemeentenaam, Municipality), by="Municipality") %>%
  select(Wijkcode, Wijknaam, Municipality, Gemeentenaam) %>%
  mutate(Wijkcode=str_sub(Wijkcode, start=3))

# data van de gewichtsmetingen
data = read_spss("L:\\AfdgGGDKenniscentrum\\02  Kompas Volksgezondheid NOG\\KVNOG\\Gegevens op KVNOG\\. Leefstijl\\Gewicht - Jeugd\\2025\\Kompas gewicht 2021-2025.sav")
#data = read_spss("L:\\AfdgGGDKenniscentrum\\02  Kompas Volksgezondheid NOG\\KVNOG\\Gegevens op KVNOG\\. Leefstijl\\Gewicht - Jeugd\\2024\\Kompas gewicht 2010-2024.sav")
data$eindjaar = as.numeric(str_sub(data$Schooljaar, end=4)) + 1
labels = lapply(colnames(data), function (v) { return(val_labels(data[[v]])) })
labels = setNames(labels, colnames(data))
labels[["Geslacht"]] = c("Man"=1, "Vrouw"=2, "Niet"=3)
data = unlabelled(data)
data$Gewicht5[data$Gewicht5 == names(labels[["Gewicht5"]])[5]] = names(labels[["Gewicht5"]])[4]
data$Gewicht5 = fct_drop(data$Gewicht5)
labels[["Gewicht5"]] = labels[["Gewicht5"]][1:4]

jaar.min = min(data$eindjaar)
jaar.max = max(data$eindjaar)
# AH = 3, N-V = 4, MIJ-OV = 1
labels[["Regiokind"]] = c("Midden-IJssel/Oost-Veluwe"=1, "Achterhoek"=3, "Noord-Veluwe"=4)
subregios = unique(data$Regiokind)

freq.table(data$Gewicht5, data$Schooljaar)

ggplot(data, aes(x=Schooljaar, group=Gewicht5, fill=Gewicht5)) +
  geom_bar(position="dodge")

# we willen hebben:
# - percentage per gewichtklasse
# - percentage overgewicht of obesitas
# - percentage obesitas
# alles gesplitst naar leeftijd en geslacht, per gemeente, subregio, NOG
# vervolgens alles met een moving average over 4 jaar
# aangezien niet alle groepen data hebben in alle jaren moeten we dit helaas wel synthetisch opvullen (mag gewoon met missings)

# gewichtsklasse per wijk
printf("Gewichtsklasse per wijk")
gewichtsklasse.wijk = data %>%
  group_by(Gemeentekind, Wijkkind, Geslacht, LFTkompas, eindjaar, Gewicht5) %>%
  summarize(n=n()) %>%
  ungroup() %>%
  complete(Gemeentekind, Wijkkind, Geslacht, LFTkompas, eindjaar, Gewicht5, fill=list(n=0)) %>%
  group_by(Gemeentekind, Wijkkind, Geslacht, LFTkompas, eindjaar) %>%
  mutate(perc=n/sum(n)*100, n_total=sum(n)) %>%
  ungroup()
for (gem in gemeenten.nog) {
  wijken = unique(gewichtsklasse.wijk$Wijkkind[gewichtsklasse.wijk$Gemeentekind == gem])
  for (wijk in wijken) {
    for (jaar in jaar.min:jaar.max) {
      if (nrow(gewichtsklasse.wijk[gewichtsklasse.wijk$Gemeentekind == gem & gewichtsklasse.wijk$Wijkkind == wijk & gewichtsklasse.wijk$eindjaar == jaar,]) < 1) {
        printf("  Missend jaar: %s (%d)", gem, jaar)
        gewichtsklasse.wijk = bind_rows(gewichtsklasse.wijk, data.frame(Gemeentekind=gem, Wijkkind=wijk, eindjaar=jaar, Geslacht="Man"))
        gewichtsklasse.wijk$Gewicht5[nrow(gewichtsklasse.wijk)] = "obesitas"
        gewichtsklasse.wijk$LFTkompas[nrow(gewichtsklasse.wijk)] = names(labels[["LFTkompas"]])[1]
      }
    }
  }
}
gewichtsklasse.wijk = gewichtsklasse.wijk %>%
  complete(Gemeentekind, Wijkkind, Geslacht, LFTkompas, eindjaar, Gewicht5) %>%
  group_by(Gemeentekind, Wijkkind, Geslacht, LFTkompas, Gewicht5) %>%
  mutate(n.ma=rollapply(n, 4, mean, fill=NA, partial=T, align="right", na.rm=T),
         perc.ma=rollapply(perc, 4, mean, fill=NA, partial=T, align="right", na.rm=T),
         n_total.ma=rollapply(n_total, 4, mean, fill=NA, partial=T, align="right", na.rm=T),
         n.weighted=perc.ma*n_total/100)

# gewichtsklasse per gemeente, gesplitst op alles
printf("Gewichtsklasse per gemeente")
gewichtsklasse.gemeente = data %>%
  group_by(Gemeentekind, Geslacht, LFTkompas, eindjaar, Gewicht5) %>%
  summarize(n=n()) %>%
  ungroup() %>%
  complete(Gemeentekind, Geslacht, LFTkompas, eindjaar, Gewicht5, fill=list(n=0)) %>%
  group_by(Gemeentekind, Geslacht, LFTkompas, eindjaar) %>%
  mutate(perc=n/sum(n)*100, n_total=sum(n)) %>%
  ungroup()
for (gem in gemeenten.nog) {
  for (jaar in jaar.min:jaar.max) {
    if (nrow(gewichtsklasse.gemeente[gewichtsklasse.gemeente$Gemeentekind == gem & gewichtsklasse.gemeente$eindjaar == jaar,]) < 1) {
      printf("  Missend jaar: %s (%d)", gem, jaar)
      gewichtsklasse.gemeente = bind_rows(gewichtsklasse.gemeente, data.frame(Gemeentekind=gem, eindjaar=jaar, Geslacht="Man"))
      gewichtsklasse.gemeente$Gewicht5[nrow(gewichtsklasse.gemeente)] = "obesitas"
      gewichtsklasse.gemeente$LFTkompas[nrow(gewichtsklasse.gemeente)] = names(labels[["LFTkompas"]])[1]
    }
  }
}
gewichtsklasse.gemeente = gewichtsklasse.gemeente %>%
  complete(Gemeentekind, Geslacht, LFTkompas, eindjaar, Gewicht5) %>%
  group_by(Gemeentekind, Geslacht, LFTkompas, Gewicht5) %>%
  mutate(n.ma=rollapply(n, 4, mean, fill=NA, partial=T, align="right", na.rm=T),
         perc.ma=rollapply(perc, 4, mean, fill=NA, partial=T, align="right", na.rm=T),
         n_total.ma=rollapply(n_total, 4, mean, fill=NA, partial=T, align="right", na.rm=T),
         n.weighted=perc.ma*n_total/100)

# gewichtsklasse per subregio, gesplitst op alles
printf("Gewichtsklasse per subregio")
gewichtsklasse.subregio = data %>%
  group_by(Regiokind, Geslacht, LFTkompas, eindjaar, Gewicht5) %>%
  summarize(n=n()) %>%
  ungroup() %>%
  complete(Regiokind, Geslacht, LFTkompas, eindjaar, Gewicht5, fill=list(n=0)) %>%
  group_by(Regiokind, Geslacht, LFTkompas, eindjaar) %>%
  mutate(perc=n/sum(n)*100, n_total=sum(n)) %>%
  ungroup()
for (subregio in subregios) {
  for (jaar in jaar.min:jaar.max) {
    if (nrow(gewichtsklasse.subregio[gewichtsklasse.subregio$Regiokind == subregio & gewichtsklasse.subregio$eindjaar == jaar,]) < 1) {
      printf("  Missend jaar: %s (%d)", subregio, jaar)
      gewichtsklasse.subregio = bind_rows(gewichtsklasse.subregio, data.frame(Regiokind = subregio, eindjaar=jaar, Geslacht="Man"))
      gewichtsklasse.subregio$Gewicht5[nrow(gewichtsklasse.subregio)] = "obesitas"
      gewichtsklasse.subregio$LFTkompas[nrow(gewichtsklasse.subregio)] = names(labels[["LFTkompas"]])[1]
    }
  }
}
gewichtsklasse.subregio = gewichtsklasse.subregio %>%
  complete(Regiokind, Geslacht, LFTkompas, eindjaar, Gewicht5) %>%
  group_by(Regiokind, Geslacht, LFTkompas, Gewicht5) %>%
  mutate(n.ma=rollapply(n, 4, mean, fill=NA, partial=T, align="right", na.rm=T),
         perc.ma=rollapply(perc, 4, mean, fill=NA, partial=T, align="right", na.rm=T),
         n_total.ma=rollapply(n_total, 4, mean, fill=NA, partial=T, align="right", na.rm=T),
         n.weighted=perc.ma*n_total/100)

# gewichtsklasse totaal, gesplitst op alles
printf("Gewichtsklasse voor NOG")
gewichtsklasse.nog = data %>%
  group_by(Geslacht, LFTkompas, eindjaar, Gewicht5) %>%
  summarize(n=n()) %>%
  ungroup() %>%
  complete(Geslacht, LFTkompas, eindjaar, Gewicht5, fill=list(n=0)) %>%
  group_by(Geslacht, LFTkompas, eindjaar) %>%
  mutate(perc=n/sum(n)*100, n_total=sum(n)) %>%
  ungroup()
for (jaar in jaar.min:jaar.max) {
  if (nrow(gewichtsklasse.nog[gewichtsklasse.nog$eindjaar == jaar,]) < 1) {
    printf("  Missend jaar: %d", jaar)
    gewichtsklasse.nog = bind_rows(gewichtsklasse.nog, data.frame(eindjaar=jaar, Geslacht="Man"))
    gewichtsklasse.nog$Gewicht5[nrow(gewichtsklasse.nog)] = "obesitas"
    gewichtsklasse.nog$LFTkompas[nrow(gewichtsklasse.nog)] = names(labels[["LFTkompas"]])[1]
  }
}
gewichtsklasse.nog = gewichtsklasse.nog %>%
  complete(Geslacht, LFTkompas, eindjaar, Gewicht5) %>%
  group_by(Geslacht, LFTkompas, Gewicht5) %>%
  mutate(n.ma=rollapply(n, 4, mean, fill=NA, partial=T, align="right", na.rm=T),
         perc.ma=rollapply(perc, 4, mean, fill=NA, partial=T, align="right", na.rm=T),
         n_total.ma=rollapply(n_total, 4, mean, fill=NA, partial=T, align="right", na.rm=T),
         n.weighted=perc.ma*n_total/100)


####
#### export
####

meta = cbs_get_meta("85755NED")
gemeentecodes = meta$RegioS[meta$RegioS$Title %in% gemeenten.nog, c("Key", "Title")]
gemeentecodes$code = str_extract(gemeentecodes$Key, "\\d+") %>% as.numeric()
gemeentecodes = gemeentecodes %>% select(-Key) %>% rename(gemeentenaam=Title)
subregiocodes = data.frame(regionaam=names(labels[["Regiokind"]]), code=unname(labels[["Regiokind"]]))

# opmaak van de export:
# tabblad Data: Jaar, geolevel, geoitem, crossings, gewicht_1, gewicht_2, enz.
wb = createWorkbook()
addWorksheet(wb, "Data")
addWorksheet(wb, "Data_def")
writeData(wb, "Data_def", cbind("col"=c("Jaar", "geolevel", "geoitem", "JG_Geslacht", "JG_Lftklasse", paste0("JG_Gewicht_", seq(1, 4))),
                                "type"=c("period", "geolevel", "geoitem", "dim", "dim", rep("var", 4))))

# combineren met wijkdata CBS
data = data %>%
  left_join(wijkindeling, by=c("Gemeentekind"="Gemeentenaam", "Wijkkind"="Wijknaam"))

output = gewichtsklasse.wijk %>%
  filter(!is.na(n_total), n_total != 0) %>%
  ungroup() %>%
  left_join(wijkindeling %>% rename(geoitem=Wijkcode) %>% select(-Municipality), by=c("Gemeentekind"="Gemeentenaam", "Wijkkind"="Wijknaam")) %>%
  mutate(geolevel="wijk23", .before=1) %>%
  rename(Jaar=eindjaar) %>%
  select(Jaar, geolevel, geoitem, Geslacht, LFTkompas, Gewicht5, n.weighted)

# output = output %>%
#   bind_rows(gewichtsklasse.subregio %>%
#               ungroup() %>%
#               left_join(subregiocodes %>% rename(geoitem=code), by=c("Regiokind"="regionaam")) %>%
#               mutate(geolevel="subregio", .before=1) %>%
#               rename(Jaar=eindjaar) %>%
#               select(Jaar, geolevel, geoitem, Geslacht, LFTkompas, Gewicht5, n.ma))
# 
# output = output %>%
#   bind_rows(gewichtsklasse.nog %>%
#               ungroup() %>%
#               mutate(geolevel="ggd", geoitem=1, .before=1) %>%
#               rename(Jaar=eindjaar) %>%
#               select(Jaar, geolevel, geoitem, Geslacht, LFTkompas, Gewicht5, n.ma))

for (var in c("Geslacht", "LFTkompas", "Gewicht5")) {
  output[[var]] = recode_values(output[[var]], labels[[var]])
  
  varname = str_replace_all(var, c("Geslacht"="JG_Geslacht", "LFTkompas"="JG_Lftklasse"))
  
  # van de gewichtsklassen is geen dimensiesheet nodig, maar een labelblad
  if (var != "Gewicht5") {
    addWorksheet(wb, varname)
    writeData(wb, varname, data.frame(Itemcode=as.numeric(labels[[var]]), Naam=names(labels[[var]]), Volgnr=1:length(labels[[var]])))
  } else {
    addWorksheet(wb, "Label_var")
    writeData(wb, "Label_var", data.frame(Onderwerpcode=paste0("JG_Gewicht_", seq(1, 4)), Naam=paste0("Aantal personen met ", names(labels[[var]])), Eenheid="Personen"))
  }
}

output = output %>%
  filter(!is.na(geoitem)) %>%
  pivot_wider(names_from="Gewicht5", names_prefix="JG_Gewicht_", values_from="n.weighted") %>%
  rename(JG_Geslacht=Geslacht, JG_Lftklasse=LFTkompas) %>%
  mutate(across(JG_Gewicht_1:JG_Gewicht_4, ~ifelse(is.nan(.x), NA, .x)))

writeData(wb, "Data", output)

addWorksheet(wb, "Dimensies")
writeData(wb, "Dimensies", data.frame(Dimensiecode=c("JG_Geslacht", "JG_Lftklasse"), Naam=c("Geslacht", "Leeftijd in 3 groepen")))

addWorksheet(wb, "Indicators")
writeData(wb, "Indicators", tibble(`Indicator code`=c(paste0("JG_Gewicht_", seq(1, 4)), "JG_Gewicht_totaal", paste0("JG_Gewicht_", seq(1, 4), "_perc"), "JG_Overgewicht", "JG_Overgewicht_perc"),
                                   Name=c(paste0("Aantal personen met ", labels[["Gewicht5"]]), "Totaal aantal personen per jaar", paste0("Percentage personen met ", labels[["Gewicht5"]]), "Aantal personen met overgewicht", "Percentage personen met overgewicht"),
                                   Unit=c(rep("personen", 5), rep("percentage", 4), "personen", "percentage"),
                                   `Aggregation indicator`=NA,
                                   Formula=c(rep(NA, 4), str_c(paste0("JG_Gewicht_", seq(1, 4)), collapse="+"),
                                             sprintf("(%s/(%s))*100", paste0("JG_Gewicht_", seq(1, 4)),
                                                     str_c(paste0("JG_Gewicht_", seq(1, 4)), collapse="+")),
                                             str_c(paste0("JG_Gewicht_", seq(1, 2)), collapse="+"),
                                             sprintf("((%s)/(%s))*100", str_c(paste0("JG_Gewicht_", seq(1, 2)), collapse="+"),
                                                     str_c(paste0("JG_Gewicht_", seq(1, 4)), collapse="+"))),
                                   `Data type`=c(rep("Numeric", 5), rep("Percentage", 4), "Numeric", "Percentage"),
                                   Visible=c(rep(0, 5), rep(1, 4), 0, 1),
                                   `Threshold value`=c(rep(NA, 5), rep(50, 4), NA, 50),
                                   `Threshold Indicator`=c(rep(NA, 5), rep("JG_Gewicht_totaal", 4), NA, "JG_Overgewicht"),
                                   Cube=1,
                                   Source="Jeugdgezondheid"))

saveWorkbook(wb, "Swing_gewichtscijfers_voortschrijdend_wijk.xlsx", overwrite=T)

### gemeentecijfers
# opmaak van de export:
# tabblad Data: Jaar, geolevel, geoitem, crossings, gewicht_1, gewicht_2, enz.
wb = createWorkbook()
addWorksheet(wb, "Data")
addWorksheet(wb, "Data_def")
writeData(wb, "Data_def", cbind("col"=c("Jaar", "geolevel", "geoitem", "JG_Geslacht", "JG_Lftklasse", paste0("JG_Gewicht_", seq(1, 4))),
                                "type"=c("period", "geolevel", "geoitem", "dim", "dim", rep("var", 4))))

output = gewichtsklasse.gemeente %>%
  filter(!is.na(n_total), n_total != 0) %>%
  ungroup() %>%
  left_join(gemeentecodes %>% rename(geoitem=code), by=c("Gemeentekind"="gemeentenaam")) %>%
  mutate(geolevel="gemeente", .before=1) %>%
  rename(Jaar=eindjaar) %>%
  select(Jaar, geolevel, geoitem, Geslacht, LFTkompas, Gewicht5, n.weighted)

# output = output %>%
#   bind_rows(gewichtsklasse.subregio %>%
#               ungroup() %>%
#               left_join(subregiocodes %>% rename(geoitem=code), by=c("Regiokind"="regionaam")) %>%
#               mutate(geolevel="subregio", .before=1) %>%
#               rename(Jaar=eindjaar) %>%
#               select(Jaar, geolevel, geoitem, Geslacht, LFTkompas, Gewicht5, n.ma))
# 
# output = output %>%
#   bind_rows(gewichtsklasse.nog %>%
#               ungroup() %>%
#               mutate(geolevel="ggd", geoitem=1, .before=1) %>%
#               rename(Jaar=eindjaar) %>%
#               select(Jaar, geolevel, geoitem, Geslacht, LFTkompas, Gewicht5, n.ma))

for (var in c("Geslacht", "LFTkompas", "Gewicht5")) {
  output[[var]] = recode_values(output[[var]], labels[[var]])
  
  varname = str_replace_all(var, c("Geslacht"="JG_Geslacht", "LFTkompas"="JG_Lftklasse"))
  
  # van de gewichtsklassen is geen dimensiesheet nodig, maar een labelblad
  if (var != "Gewicht5") {
    addWorksheet(wb, varname)
    writeData(wb, varname, data.frame(Itemcode=as.numeric(labels[[var]]), Naam=names(labels[[var]]), Volgnr=1:length(labels[[var]])))
  } else {
    addWorksheet(wb, "Label_var")
    writeData(wb, "Label_var", data.frame(Onderwerpcode=paste0("JG_Gewicht_", seq(1, 4)), Naam=paste0("Aantal personen met ", names(labels[[var]])), Eenheid="Personen"))
  }
}

output = output %>%
  pivot_wider(names_from="Gewicht5", names_prefix="JG_Gewicht_", values_from="n.weighted") %>%
  rename(JG_Geslacht=Geslacht, JG_Lftklasse=LFTkompas) %>%
  mutate(across(JG_Gewicht_1:JG_Gewicht_4, ~ifelse(is.nan(.x), NA, .x)))

writeData(wb, "Data", output)

addWorksheet(wb, "Dimensies")
writeData(wb, "Dimensies", data.frame(Dimensiecode=c("JG_Geslacht", "JG_Lftklasse"), Naam=c("Geslacht", "Leeftijd in 3 groepen")))

addWorksheet(wb, "Indicators")
writeData(wb, "Indicators", tibble(`Indicator code`=c(paste0("JG_Gewicht_", seq(1, 4)), "JG_Gewicht_totaal", paste0("JG_Gewicht_", seq(1, 4), "_perc"), "JG_Overgewicht", "JG_Overgewicht_perc"),
                                   Name=c(paste0("Aantal personen met ", labels[["Gewicht5"]]), "Totaal aantal personen per jaar", paste0("Percentage personen met ", labels[["Gewicht5"]]), "Aantal personen met overgewicht", "Percentage personen met overgewicht"),
                                   Unit=c(rep("personen", 5), rep("percentage", 4), "personen", "percentage"),
                                   `Aggregation indicator`=NA,
                                   Formula=c(rep(NA, 4), str_c(paste0("JG_Gewicht_", seq(1, 4)), collapse="+"),
                                             sprintf("(%s/(%s))*100", paste0("JG_Gewicht_", seq(1, 4)),
                                                     str_c(paste0("JG_Gewicht_", seq(1, 4)), collapse="+")),
                                             str_c(paste0("JG_Gewicht_", seq(1, 2)), collapse="+"),
                                             sprintf("((%s)/(%s))*100", str_c(paste0("JG_Gewicht_", seq(1, 2)), collapse="+"),
                                                     str_c(paste0("JG_Gewicht_", seq(1, 4)), collapse="+"))),
                                   `Data type`=c(rep("Numeric", 5), rep("Percentage", 4), "Numeric", "Percentage"),
                                   Visible=c(rep(0, 5), rep(1, 4), 0, 1),
                                   `Threshold value`=c(rep(NA, 5), rep(50, 4), NA, 50),
                                   `Threshold Indicator`=c(rep(NA, 5), rep("JG_Gewicht_totaal", 4), NA, "JG_Overgewicht"),
                                   Cube=1,
                                   Source="Jeugdgezondheid"))

saveWorkbook(wb, "Swing_gewichtscijfers_voortschrijdend_gemeente.xlsx", overwrite=T)

# absolute aantallen (dus niet voortschrijdend)
wb = createWorkbook()
addWorksheet(wb, "Data")
addWorksheet(wb, "Data_def")
writeData(wb, "Data_def", cbind("col"=c("Jaar", "geolevel", "geoitem", "JG_Geslacht", "JG_Lftklasse", paste0("JG_Gewicht_", seq(1, 4), "_abs")),
                                "type"=c("period", "geolevel", "geoitem", "dim", "dim", rep("var", 4))))

output = gewichtsklasse.gemeente %>%
  ungroup() %>%
  left_join(gemeentecodes %>% rename(geoitem=code), by=c("Gemeentekind"="gemeentenaam")) %>%
  mutate(geolevel="gemeente", .before=1) %>%
  rename(Jaar=eindjaar) %>%
  select(Jaar, geolevel, geoitem, Geslacht, LFTkompas, Gewicht5, n)

output = output %>%
  bind_rows(gewichtsklasse.subregio %>%
              ungroup() %>%
              left_join(subregiocodes %>% rename(geoitem=code), by=c("Regiokind"="regionaam")) %>%
              mutate(geolevel="subregio", .before=1) %>%
              rename(Jaar=eindjaar) %>%
              select(Jaar, geolevel, geoitem, Geslacht, LFTkompas, Gewicht5, n))

output = output %>%
  bind_rows(gewichtsklasse.nog %>%
              ungroup() %>%
              mutate(geolevel="ggd", geoitem=1, .before=1) %>%
              rename(Jaar=eindjaar) %>%
              select(Jaar, geolevel, geoitem, Geslacht, LFTkompas, Gewicht5, n))

for (var in c("Geslacht", "LFTkompas", "Gewicht5")) {
  output[[var]] = recode_values(output[[var]], labels[[var]])
  
  varname = str_replace_all(var, c("Geslacht"="JG_Geslacht", "LFTkompas"="JG_Lftklasse"))
  
  # van de gewichtsklassen is geen dimensiesheet nodig, maar een labelblad
  if (var != "Gewicht5") {
    addWorksheet(wb, varname)
    writeData(wb, varname, data.frame(Itemcode=as.numeric(labels[[var]]), Naam=names(labels[[var]]), Volgnr=1:length(labels[[var]])))
  } else {
    addWorksheet(wb, "Label_var")
    writeData(wb, "Label_var", data.frame(Onderwerpcode=paste0("JG_Gewicht_", seq(1, 4), "_abs"), Naam=paste0("Aantal personen met ", names(labels[[var]])), Eenheid="Personen"))
  }
}

output = output %>%
  pivot_wider(names_from="Gewicht5", names_glue=sprintf("JG_Gewicht_{%s}_abs", var), values_from="n") %>%
  rename(JG_Geslacht=Geslacht, JG_Lftklasse=LFTkompas) %>%
  mutate(across(JG_Gewicht_1_abs:JG_Gewicht_4_abs, ~ifelse(is.nan(.x), NA, .x)))

writeData(wb, "Data", output)

addWorksheet(wb, "Dimensies")
writeData(wb, "Dimensies", data.frame(Dimensiecode=c("JG_Geslacht", "JG_Lftklasse"), Naam=c("Geslacht", "Leeftijd in 3 groepen")))

addWorksheet(wb, "Indicators")
writeData(wb, "Indicators", tibble(`Indicator code`=c(paste0("JG_Gewicht_", seq(1, 4), "_abs"), paste0("JG_Gewicht_", seq(1, 4), "_abs_perc")),
                                   Name=c(paste0("Aantal personen met ", labels[["Gewicht5"]]), paste0("Percentage personen met ", labels[["Gewicht5"]])),
                                   Unit=c(rep("personen", 4), rep("percentage", 4)),
                                   `Aggregation indicator`=NA,
                                   Formula=c(rep(NA, 4), sprintf("(%s/(%s))*100", paste0("JG_Gewicht_", seq(1, 4), "_abs"), str_c(paste0("JG_Gewicht_", seq(1, 4), "_abs"), collapse="+"))),
                                   `Data type`=c(rep("Numeric", 4), rep("Percentage", 4)),
                                   Visible=c(rep(0, 4), rep(1, 4)),
                                   `Threshold value`=c(rep(NA, 4), rep(50, 4)),
                                   `Threshold Indicator`=c(rep(NA, 4), paste0("JG_Gewicht_", seq(1, 4), "_abs")),
                                   Cube=1,
                                   Source="Jeugdgezondheid"))

saveWorkbook(wb, "Swing_gewichtscijfers_absoluut.xlsx", overwrite=T)


# menselijk leesbare versie
wb = createWorkbook()
addWorksheet(wb, "Alle data")
writeData(wb, "Alle data", gewichtsklasse.gemeente %>% select(-n.ma), withFilter=T)

for (gem in gemeenten.nog) {
  output = gewichtsklasse.gemeente %>%
    ungroup() %>%
    filter(Gemeentekind == gem) %>%
    select(Geslacht, LFTkompas, eindjaar, Gewicht5, perc.ma) %>%
    group_by(Geslacht, LFTkompas, eindjaar) %>%
    pivot_wider(names_from="Gewicht5", values_from="perc.ma")
  
  addWorksheet(wb, gem)
  writeData(wb, gem, output, withFilter=T)
}

saveWorkbook(wb, "gewichtscijfers_tabel.xlsx", overwrite=T)









###
### losse berekeningen, mochten we het ooit willen controleren
###


# gewichtsklasse per geslacht per gemeente
printf("Gewichtsklasse per geslacht per gemeente")
gewichtsklasse.geslacht.gemeente = data %>%
  group_by(Gemeentekind, Geslacht, eindjaar, Gewicht5) %>%
  summarize(n=n()) %>%
  group_by(Gemeentekind, Geslacht, eindjaar) %>%
  mutate(perc=n/sum(n)*100) %>%
  ungroup()
for (gem in gemeenten.nog) {
  for (jaar in jaar.min:jaar.max) {
    if (nrow(gewichtsklasse.geslacht.gemeente[gewichtsklasse.geslacht.gemeente$Gemeentekind == gem & gewichtsklasse.geslacht.gemeente$eindjaar == jaar,]) < 1) {
      printf("  Missend jaar: %s (%d)", gem, jaar)
      gewichtsklasse.geslacht.gemeente = bind_rows(gewichtsklasse.geslacht.gemeente, data.frame(Gemeentekind=gem, eindjaar=jaar))
      gewichtsklasse.geslacht.gemeente$Gewicht5[nrow(gewichtsklasse.geslacht.gemeente)] = "obesitas"
    }
  }
}
gewichtsklasse.geslacht.gemeente = gewichtsklasse.geslacht.gemeente %>%
  complete(Gemeentekind, Geslacht, eindjaar, Gewicht5) %>%
  group_by(Gemeentekind, Geslacht, Gewicht5) %>%
  mutate(n.ma=rollapply(n, 4, mean, fill=NA, partial=T, align="right", na.rm=T),
         perc.ma=rollapply(perc, 4, mean, fill=NA, partial=T, align="right", na.rm=T))

# gewichtsklasse per geslacht per subregio
printf("Gewichtsklasse per geslacht per subregio")
gewichtsklasse.geslacht.subregio = data %>%
  group_by(Regiokind, Geslacht, eindjaar, Gewicht5) %>%
  summarize(n=n()) %>%
  group_by(Regiokind, Geslacht, eindjaar) %>%
  mutate(perc=n/sum(n)*100) %>%
  ungroup()
for (subregio in subregios) {
  for (jaar in jaar.min:jaar.max) {
    if (nrow(gewichtsklasse.geslacht.subregio[gewichtsklasse.geslacht.subregio$Regiokind == subregio & gewichtsklasse.geslacht.subregio$eindjaar == jaar,]) < 1) {
      printf("  Missend jaar: %s (%d)", subregio, jaar)
      gewichtsklasse.geslacht.subregio = bind_rows(gewichtsklasse.geslacht.subregio, data.frame(Regiokind = subregio, eindjaar=jaar))
      gewichtsklasse.geslacht.subregio$Gewicht5[nrow(gewichtsklasse.geslacht.subregio)] = "obesitas"
    }
  }
}
gewichtsklasse.geslacht.subregio = gewichtsklasse.geslacht.subregio %>%
  complete(Regiokind, Geslacht, eindjaar, Gewicht5) %>%
  group_by(Regiokind, Geslacht, Gewicht5) %>%
  mutate(n.ma=rollapply(n, 4, mean, fill=NA, partial=T, align="right", na.rm=T),
         perc.ma=rollapply(perc, 4, mean, fill=NA, partial=T, align="right", na.rm=T))

# gewichtsklasse per geslacht voor NOG
printf("Gewichtsklasse per geslacht voor NOG")
gewichtsklasse.geslacht.nog = data %>%
  group_by(Geslacht, eindjaar, Gewicht5) %>%
  summarize(n=n()) %>%
  group_by(Geslacht, eindjaar) %>%
  mutate(perc=n/sum(n)*100) %>%
  ungroup()
for (jaar in jaar.min:jaar.max) {
  if (nrow(gewichtsklasse.geslacht.nog[gewichtsklasse.geslacht.nog$eindjaar == jaar,]) < 1) {
    printf("  Missend jaar: %d (NOG)", jaar)
    gewichtsklasse.geslacht.nog = bind_rows(gewichtsklasse.geslacht.nog, data.frame(eindjaar=jaar))
    gewichtsklasse.geslacht.nog$Gewicht5[nrow(gewichtsklasse.geslacht.nog)] = "obesitas"
  }
}
gewichtsklasse.geslacht.nog = gewichtsklasse.geslacht.nog %>%
  complete(Geslacht, eindjaar, Gewicht5) %>%
  group_by(Geslacht, Gewicht5) %>%
  mutate(n.ma=rollapply(n, 4, mean, fill=NA, partial=T, align="right", na.rm=T),
         perc.ma=rollapply(perc, 4, mean, fill=NA, partial=T, align="right", na.rm=T))

# gewichtsklasse per leeftijd per gemeente
printf("Gewichtsklasse per leeftijd per gemeente")
gewichtsklasse.leeftijd.gemeente = data %>%
  group_by(Gemeentekind, LFTkompas, eindjaar, Gewicht5) %>%
  summarize(n=n()) %>%
  group_by(Gemeentekind, LFTkompas, eindjaar) %>%
  mutate(perc=n/sum(n)*100) %>%
  ungroup()
for (gem in gemeenten.nog) {
  for (jaar in jaar.min:jaar.max) {
    if (nrow(gewichtsklasse.leeftijd.gemeente[gewichtsklasse.leeftijd.gemeente$Gemeentekind == gem & gewichtsklasse.leeftijd.gemeente$eindjaar == jaar,]) < 1) {
      printf("  Missend jaar: %s (%d)", gem, jaar)
      gewichtsklasse.leeftijd.gemeente = bind_rows(gewichtsklasse.leeftijd.gemeente, data.frame(Gemeentekind=gem, eindjaar=jaar))
      gewichtsklasse.leeftijd.gemeente$Gewicht5[nrow(gewichtsklasse.leeftijd.gemeente)] = "obesitas"
    }
  }
}
gewichtsklasse.leeftijd.gemeente = gewichtsklasse.leeftijd.gemeente %>%
  complete(Gemeentekind, LFTkompas, eindjaar, Gewicht5) %>%
  group_by(Gemeentekind, LFTkompas, Gewicht5) %>%
  mutate(n.ma=rollapply(n, 4, mean, fill=NA, partial=T, align="right", na.rm=T),
         perc.ma=rollapply(perc, 4, mean, fill=NA, partial=T, align="right", na.rm=T))

# gewichtsklasse per leeftijd per subregio
printf("Gewichtsklasse per leeftijd per subregio")
gewichtsklasse.leeftijd.subregio = data %>%
  group_by(Regiokind, LFTkompas, eindjaar, Gewicht5) %>%
  summarize(n=n()) %>%
  group_by(Regiokind, LFTkompas, eindjaar) %>%
  mutate(perc=n/sum(n)*100) %>%
  ungroup()
for (subregio in subregios) {
  for (jaar in jaar.min:jaar.max) {
    if (nrow(gewichtsklasse.leeftijd.subregio[gewichtsklasse.leeftijd.subregio$Regiokind == subregio & gewichtsklasse.leeftijd.subregio$eindjaar == jaar,]) < 1) {
      printf("  Missend jaar: %s (%d)", subregio, jaar)
      gewichtsklasse.leeftijd.subregio = bind_rows(gewichtsklasse.leeftijd.subregio, data.frame(Regiokind = subregio, eindjaar=jaar))
      gewichtsklasse.leeftijd.subregio$Gewicht5[nrow(gewichtsklasse.leeftijd.subregio)] = "obesitas"
    }
  }
}
gewichtsklasse.leeftijd.subregio = gewichtsklasse.leeftijd.subregio %>%
  complete(Regiokind, LFTkompas, eindjaar, Gewicht5) %>%
  group_by(Regiokind, LFTkompas, Gewicht5) %>%
  mutate(n.ma=rollapply(n, 4, mean, fill=NA, partial=T, align="right", na.rm=T),
         perc.ma=rollapply(perc, 4, mean, fill=NA, partial=T, align="right", na.rm=T))

# gewichtsklasse per leeftijd voor NOG
printf("Gewichtsklasse per leeftijd voor NOG")
gewichtsklasse.leeftijd.nog = data %>%
  group_by(LFTkompas, eindjaar, Gewicht5) %>%
  summarize(n=n()) %>%
  group_by(LFTkompas, eindjaar) %>%
  mutate(perc=n/sum(n)*100) %>%
  ungroup()
for (jaar in jaar.min:jaar.max) {
  if (nrow(gewichtsklasse.leeftijd.nog[gewichtsklasse.leeftijd.nog$eindjaar == jaar,]) < 1) {
    printf("  Missend jaar: %d (NOG)", jaar)
    gewichtsklasse.leeftijd.nog = bind_rows(gewichtsklasse.leeftijd.nog, data.frame(eindjaar=jaar))
    gewichtsklasse.leeftijd.nog$Gewicht5[nrow(gewichtsklasse.leeftijd.nog)] = "obesitas"
  }
}
gewichtsklasse.leeftijd.nog = gewichtsklasse.leeftijd.nog %>%
  complete(LFTkompas, eindjaar, Gewicht5) %>%
  group_by(LFTkompas, Gewicht5) %>%
  mutate(n.ma=rollapply(n, 4, mean, fill=NA, partial=T, align="right", na.rm=T),
         perc.ma=rollapply(perc, 4, mean, fill=NA, partial=T, align="right", na.rm=T))

# gewichtsklasse totaal per gemeente
printf("Gewichtsklasse totaal per gemeente")
gewichtsklasse.totaal.gemeente = data %>%
  group_by(Gemeentekind, eindjaar, Gewicht5) %>%
  summarize(n=n()) %>%
  group_by(Gemeentekind, eindjaar) %>%
  mutate(perc=n/sum(n)*100) %>%
  ungroup()
for (gem in gemeenten.nog) {
  for (jaar in jaar.min:jaar.max) {
    if (nrow(gewichtsklasse.totaal.gemeente[gewichtsklasse.totaal.gemeente$Gemeentekind == gem & gewichtsklasse.totaal.gemeente$eindjaar == jaar,]) < 1) {
      printf("  Missend jaar: %s (%d)", gem, jaar)
      gewichtsklasse.totaal.gemeente = bind_rows(gewichtsklasse.totaal.gemeente, data.frame(Gemeentekind=gem, eindjaar=jaar))
      gewichtsklasse.totaal.gemeente$Gewicht5[nrow(gewichtsklasse.totaal.gemeente)] = "obesitas"
    }
  }
}
gewichtsklasse.totaal.gemeente = gewichtsklasse.totaal.gemeente %>%
  complete(Gemeentekind, eindjaar, Gewicht5) %>%
  group_by(Gemeentekind, Gewicht5) %>%
  mutate(n.ma=rollapply(n, 4, mean, fill=NA, partial=T, align="right", na.rm=T),
         perc.ma=rollapply(perc, 4, mean, fill=NA, partial=T, align="right", na.rm=T))

# gewichtsklasse totaal per subregio
printf("Gewichtsklasse totaal per subregio")
gewichtsklasse.totaal.subregio = data %>%
  group_by(Regiokind, eindjaar, Gewicht5) %>%
  summarize(n=n()) %>%
  group_by(Regiokind, eindjaar) %>%
  mutate(perc=n/sum(n)*100) %>%
  ungroup()
for (subregio in subregios) {
  for (jaar in jaar.min:jaar.max) {
    if (nrow(gewichtsklasse.totaal.subregio[gewichtsklasse.totaal.subregio$Regiokind == subregio & gewichtsklasse.totaal.subregio$eindjaar == jaar,]) < 1) {
      printf("  Missend jaar: %s (%d)", subregio, jaar)
      gewichtsklasse.totaal.subregio = bind_rows(gewichtsklasse.totaal.subregio, data.frame(Regiokind = subregio, eindjaar=jaar))
      gewichtsklasse.totaal.subregio$Gewicht5[nrow(gewichtsklasse.totaal.subregio)] = "obesitas"
    }
  }
}
gewichtsklasse.totaal.subregio = gewichtsklasse.totaal.subregio %>%
  complete(Regiokind, eindjaar, Gewicht5) %>%
  group_by(Regiokind, Gewicht5) %>%
  mutate(n.ma=rollapply(n, 4, mean, fill=NA, partial=T, align="right", na.rm=T),
         perc.ma=rollapply(perc, 4, mean, fill=NA, partial=T, align="right", na.rm=T))

# gewichtsklasse totaal voor NOG
printf("Gewichtsklasse totaal voor NOG")
gewichtsklasse.totaal.nog = data %>%
  group_by(eindjaar, Gewicht5) %>%
  summarize(n=n()) %>%
  group_by(eindjaar) %>%
  mutate(perc=n/sum(n)*100) %>%
  ungroup()
for (jaar in jaar.min:jaar.max) {
  if (nrow(gewichtsklasse.totaal.nog[gewichtsklasse.totaal.nog$eindjaar == jaar,]) < 1) {
    printf("  Missend jaar: %d (NOG)", jaar)
    gewichtsklasse.totaal.nog = bind_rows(gewichtsklasse.totaal.nog, data.frame(eindjaar=jaar))
    gewichtsklasse.totaal.nog$Gewicht5[nrow(gewichtsklasse.totaal.nog)] = "obesitas"
  }
}
gewichtsklasse.totaal.nog = gewichtsklasse.totaal.nog %>%
  complete(eindjaar, Gewicht5) %>%
  group_by(Gewicht5) %>%
  mutate(n.ma=rollapply(n, 4, mean, fill=NA, partial=T, align="right", na.rm=T),
         perc.ma=rollapply(perc, 4, mean, fill=NA, partial=T, align="right", na.rm=T))

####
#### overgewicht of obesitas
####

# proof of concept: som van gemiddelde percentages zou identiek moeten zijn

# obesitas per geslacht per gemeente
printf("Obesitas per geslacht per gemeente")
obesitas.geslacht.gemeente = data %>%
  group_by(Gemeentekind, Geslacht, eindjaar, tezwaar) %>%
  summarize(n=n()) %>%
  group_by(Gemeentekind, Geslacht, eindjaar) %>%
  mutate(perc=n/sum(n)*100) %>%
  ungroup()
for (gem in gemeenten.nog) {
  for (jaar in jaar.min:jaar.max) {
    if (nrow(obesitas.geslacht.gemeente[obesitas.geslacht.gemeente$Gemeentekind == gem & obesitas.geslacht.gemeente$eindjaar == jaar,]) < 1) {
      printf("  Missend jaar: %s (%d)", gem, jaar)
      obesitas.geslacht.gemeente = bind_rows(obesitas.geslacht.gemeente, data.frame(Gemeentekind=gem, eindjaar=jaar))
      obesitas.geslacht.gemeente$tezwaar[nrow(obesitas.geslacht.gemeente)] = "Nee"
    }
  }
}
obesitas.geslacht.gemeente = obesitas.geslacht.gemeente %>%
  complete(Gemeentekind, Geslacht, eindjaar, tezwaar) %>%
  group_by(Gemeentekind, Geslacht, tezwaar) %>%
  mutate(n.ma=rollapply(n, 4, mean, fill=NA, partial=T, align="right", na.rm=T),
         perc.ma=rollapply(perc, 4, mean, fill=NA, partial=T, align="right", na.rm=T))

obesitas.geslacht.gemeente %>% filter(Gemeentekind == "Aalten")
gewichtsklasse.geslacht.gemeente %>% filter(Gemeentekind == "Aalten", Gewicht5 %in% c("obesitas", "overgewicht"))
# klopt -> exporteren vanuit basistabellen


