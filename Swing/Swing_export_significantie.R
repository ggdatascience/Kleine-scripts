# Script om verschillen tussen indicatoren van gemeentes en de gehele regio te berekenen en te exporteren naar Swing.
#
# Huidige opzet:
# 1 = de gemeente scoort ongunstig tov NOG én met een p<0.01.
# 2 = de gemeente scoort ongunstig tov NOG én met een p tussen 0.01 en 0.05
# 3 = de p-waarde is groter dan 0.1
# 4 = de gemeente scoort gunstig tov NOG én met een p tussen 0.01 en 0.05
# 5 = de gemeente scoort gunstig tov NOG én met een p<0.01. 

library(tidyverse)
library(haven)
library(labelled)
library(survey)
library(openxlsx)
library(this.path)
library(ggdnog)

setwd(dirname(this.path()))

# gewenste variabelen om te berekenen
# let op: enkel dichotome variabelen toegestaan! (0/1 voor nee/ja, respectievelijk)
vars = read.xlsx("Overzicht keuze indicatoren VO.xlsx")
#vars = data.frame(Code=c("AGHHB203", "GGEEB201", "LFRKA205", "LFALA218"), Richting=c("Hoger = ongunstig", "Hoger = gunstig", "Hoger = ongunstig", "Hoger = gunstig"))
# naamgeving van de variabelen; {var} bevat de variabelenaam, {.value} de berekende waarde (bijv. p, pgroep, hooglaag)
var_naamgeving = "{var}_{.value}"
# naamgeving van de output: %s wordt vervangen door de variabelenaam
xlsx_naamgeving = "kubus_%s_p.xlsx"
# bron moet aangegeven worden in Swing - dit wordt in de uitvoersheets geplakt
bron = "Gezondheidsmonitor Volwassenen en Ouderen"
# jaartal, nodig voor Swing
jaar = 2024
# Let op: beveiliging met een wachtwoord is niet mogelijk in R. Als het originele bestand dit wel is, sla het dan in SPSS opnieuw op zonder wachtwoord.
data = read_sav(file.choose(), user_na=T) %>% user_na_to_na()
# moet er geselecteerd worden op iets? 
data = data %>%
  filter(VO_AGOJB201 == 2024)
# de gewichten en strata verschillen per monitor
# vul deze in zoals hieronder, dus "~[variabelenaam]"
gewichten = ~VO_ewCBSGGD
strata = ~VO_PrimaireEenheid

###
### VANAF HIER NIKS AANPASSEN
###

vars$Richting_code = str_match(vars$Richting, ".*= (\\w+)")[,2]
vars$hoger_beter = vars$Richting_code == "gunstig"
if (any(is.na(vars$hoger_beter))) {
  stop("Ongeldige waarde in kolom Richting. Deze cellen moeten voldoen aan de indeling 'Hoger = ongunstig' of 'Hoger = gunstig'.")
}

# variabelenamen corrigeren indien verkeerde hoofdletters
varnames_corr = sapply(vars$Code, \(x) {
  corrected = str_detect(colnames(data), regex(x, ignore_case=T))
  if (any(corrected)) {
    return(colnames(data)[corrected])
  }
  return("")
})
vars$Code = unlist(varnames_corr)
vars = vars[vars$Code != "",]

colnames(data)[grep("gemeente", colnames(data), ignore.case=T)] = "Gemeentecode"

# overzicht met gemeenten met data; die hebben we later nodig
gemeenten = data.frame(code=unname(val_labels(data$Gemeentecode)), naam=names(val_labels(data$Gemeentecode))) %>%
  filter(code %in% unique(data$Gemeentecode))

# nu gaan we door de lijst met variabelen lopen en berekenen we per gemeente het percentage en de afwijking daarvan
# dit moet met een dummy, aangezien survey niet kan werken met selecties binnen []
results = data.frame()
for (v in 1:nrow(vars)) {
  var = vars$Code[v]
  for (i in 1:nrow(gemeenten)) {
    data$dummy_gemeente = data$Gemeentecode == gemeenten$code[i]
    
    design = svydesign(ids=~1, strata=strata, weights=gewichten, data=data)
    
    n = svytable(as.formula(paste0("~",var,"+dummy_gemeente")), design)
    perc = proportions(n, margin=2)
    test = svychisq(as.formula(paste0("~",var,"+dummy_gemeente")), design)
    
    if (nrow(perc) != 2) {
      printf("Let op: afwijkend aantal antwoordmogelijkheden voor %s: %s", var, str_c(rownames(perc), collapse=", "))
    }
    
    results = bind_rows(results,
                        data.frame(var=var, gemeente=gemeenten$code[i],
                                   perc_gemeente=perc[2, "TRUE"], perc_nog=perc[2, "FALSE"],
                                   p=test$p.value, hoger_beter=vars$hoger_beter[v]))
  }
}

# omzetten naar Swingformaat en exporteren
for (v in 1:nrow(vars)) {
  # Swing-format:
  # Jaar, geolevel, geoitem, [var]_hooglaag, [var]_p, [var]_pgroep
  # pivot_wider doet normaal gesproken [value]_[var], maar dit kun je omdraaien met names_glue
  swing_data = results %>% 
    filter(var == vars$Code[v]) %>%
    rename(geoitem=gemeente) %>%
    mutate(Jaar=jaar, geolevel="gemeente", gunstigongunstig=ifelse(hoger_beter, perc_gemeente > perc_nog, perc_nog > perc_gemeente),
           pgroep=case_when(p > 0.05 ~ 3,
                            !gunstigongunstig & p < 0.01 ~ 1,
                            !gunstigongunstig & p >= 0.01 ~ 2,
                            gunstigongunstig & p >= 0.01 ~ 4,
                            gunstigongunstig & p < 0.01 ~ 5),
           gunstigongunstig=as.numeric(gunstigongunstig)) %>%
    select(-c(perc_gemeente, perc_nog, hoger_beter, gunstigongunstig, p)) %>%
    pivot_wider(names_from="var", values_from=c("pgroep"), names_glue=var_naamgeving) %>%
    relocate(geoitem, .after="geolevel")
  
  wb = createWorkbook()
  addWorksheet(wb, "Data")
  writeData(wb, "Data", swing_data)
  
  addWorksheet(wb, "Data_def")
  writeData(wb, "Data_def", cbind("col"=colnames(swing_data),
                                  "type"=c("period", "geolevel", "geoitem", "var")))
                                  # "type"=c("period", "geolevel", "geoitem", "var", "var", "var")))
  
  addWorksheet(wb, "Label_var")
  writeData(wb, "Label_var", data.frame(Onderwerpcode=colnames(swing_data)[4],
                                        Naam=c("Significantie groep"), Eenheid="Geen"))
                                        # Naam=c("Gunstig of ongunstig tov NOG", "p-waarde", "Significantie groep"), Eenheid="Geen"))
  
  addWorksheet(wb, "Indicators")
  writeData(wb, "Indicators", tibble(`Indicator code`=colnames(swing_data)[4],
                                     Name=c("Significantie groep"),
                                     #Name=c("Gunstig of ongunstig tov NOG", "p-waarde", "Significantie groep"),
                                     Unit=rep("Geen", 1),
                                     `Aggregation indicator`=NA,
                                     Formula=rep(NA, 1),
                                     `Data type`=rep("Numeric", 1),
                                     Visible=rep(1, 1),
                                     `Threshold value`=rep(NA, 1),
                                     `Threshold Indicator`=rep(NA, 1),
                                     Cube=0,
                                     Source=bron))
  
  saveWorkbook(wb, sprintf(xlsx_naamgeving, vars$Code[v]), overwrite=T)
}

