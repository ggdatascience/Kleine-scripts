library(tidyverse)
library(ggdnog)
library(cbsodataR)
library(openxlsx)
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

meta = cbs_get_meta("50141NED", catalog="RIVM")
vacc.graad = cbs_get_data("50141NED", catalog="RIVM", RegioS=meta$RegioS$Key[meta$RegioS$Title %in% gemeenten.nog])
vacc.graad = vacc.graad %>%
  cbs_add_label_columns() %>%
  cbs_add_date_column() %>%
  mutate(Jaar=as.numeric(str_extract(Perioden, "\\d+")),
         Regio=str_trim(RegioS)) %>%
  rename(Populatie=Populatie_1) %>%
  mutate(Gevaccineerden=coalesce(GevaccineerdenZonderLeeftijdsgrens_4, GevaccineerdenMetLeeftijdsgrens_2), 
         Vaccinatiegraad=coalesce(VaccinatiegraadZonderLeeftijdsgrens_5, VaccinatiegraadMetLeeftijdsgrens_3),
         gemeentecode=str_sub(RegioS, start=3) %>% as.numeric())
jaar.vacc = max(vacc.graad$Jaar, na.rm=T) # nieuwste cohort

vacc.graad.nog = cbs_get_data("50141NED", catalog="RIVM", RegioS="GG1413")
vacc.graad.nog = vacc.graad.nog %>%
  cbs_add_label_columns() %>%
  mutate(Jaar=as.numeric(str_extract(Perioden, "\\d+")),
         Regio=str_trim(RegioS)) %>%
  rename(Populatie=Populatie_1) %>%
  mutate(Gevaccineerden=coalesce(GevaccineerdenZonderLeeftijdsgrens_4, GevaccineerdenMetLeeftijdsgrens_2), 
         Vaccinatiegraad=coalesce(VaccinatiegraadZonderLeeftijdsgrens_5, VaccinatiegraadMetLeeftijdsgrens_3)) %>%
  filter(!is.na(Vaccinatiegraad)) %>%
  arrange(Jaar)

vacc.graad.nl = cbs_get_data("50141NED", catalog="RIVM", RegioS="NL01  ")
vacc.graad.nl = vacc.graad.nl %>%
  cbs_add_label_columns() %>%
  mutate(Jaar=as.numeric(str_extract(Perioden, "\\d+")),
         Regio=str_trim(RegioS)) %>%
  rename(Populatie=Populatie_1) %>%
  mutate(Gevaccineerden=coalesce(GevaccineerdenZonderLeeftijdsgrens_4, GevaccineerdenMetLeeftijdsgrens_2), 
         Vaccinatiegraad=coalesce(VaccinatiegraadZonderLeeftijdsgrens_5, VaccinatiegraadMetLeeftijdsgrens_3)) %>%
  filter(!is.na(Vaccinatiegraad)) %>%
  arrange(Jaar)

vacc.types = c("DKTP basisimmuun (2 jaar)", "Hib volledig (2 jaar)", "BMR basisimmuun (2 jaar)", "MenC/ACWY basisimmuun (2 jaar)",
               "Pneumokokken volledig (2 jaar)", "Hepatitis B volledig (2 jaar)", "D(K)TP volledig (10 jaar)", "BMR volledig (10 jaar)",
               "HPV volledig (14 jaar)", "MenACWY volledig (cohorten 2001-2005)", "MenACWY volledig (15 jaar)", "Geen enkele vaccinatie (2 jaar)",
               "HPV volledig meisjes (11 jaar)", "HPV volledig jongens (11 jaar)", "D(K)TP voldoende beschermd (5 jaar)")
vacc.types = setNames(1:length(vacc.types), vacc.types)


# opmaak van de export:
# tabblad Data: Jaar, geolevel, geoitem, vaccinatie, n, populatie
wb = createWorkbook()
addWorksheet(wb, "Data")

output = vacc.graad %>%
  select(Jaar, gemeentecode, Vaccinaties_label, Gevaccineerden, Populatie) %>%
  rename(geoitem="gemeentecode") %>%
  filter(Vaccinaties_label %in% names(vacc.types),
         !is.na(Gevaccineerden)) %>%
  mutate(geolevel="gemeente", .before=2) %>%
  bind_rows(vacc.graad.nog %>%
              select(Jaar, Vaccinaties_label, Gevaccineerden, Populatie) %>%
              filter(Vaccinaties_label %in% names(vacc.types),
                     !is.na(Gevaccineerden)) %>%
              mutate(geolevel="ggd", geoitem=1, .before=2)) %>%
  bind_rows(vacc.graad.nl %>%
              select(Jaar, Vaccinaties_label, Gevaccineerden, Populatie) %>%
              filter(Vaccinaties_label %in% names(vacc.types),
                     !is.na(Gevaccineerden)) %>%
              mutate(geolevel="nederland", geoitem=1, .before=2)) %>%
  rename(RVP_Gevaccineerden="Gevaccineerden", RVP_Populatie="Populatie", RVP_Vaccinatie="Vaccinaties_label")

output$RVP_Vaccinatie = recode_values(output$RVP_Vaccinatie, vacc.types)

writeData(wb, "Data", output)

addWorksheet(wb, "Data_def")
writeData(wb, "Data_def", cbind("col"=c("Jaar", "geolevel", "geoitem", "RVP_Vaccinatie", "RVP_Gevaccineerden", "RVP_Populatie"),
                                "type"=c("period", "geolevel", "geoitem", "dim", "var", "var")))

addWorksheet(wb, "RVP_Vaccinatie")
writeData(wb, "RVP_Vaccinatie", data.frame(Itemcode=unname(vacc.types), Naam=names(vacc.types), Volgnr=unname(vacc.types)))

addWorksheet(wb, "Dimensies")
writeData(wb, "Dimensies", data.frame(Dimensiecode=c("RVP_Vaccinatie"), Naam=c("Vaccinatietype")))

addWorksheet(wb, "Label_var")
writeData(wb, "Label_var", data.frame(Onderwerpcode=c("RVP_Gevaccineerden", "RVP_Populatie"), Naam=c("Aantal gevaccineerde individuen", "Aantal individuen in de populatie"), Eenheid="Personen"))

addWorksheet(wb, "Indicators")
writeData(wb, "Indicators", tibble(`Indicator code`=c("RVP_Gevaccineerden", "RVP_Populatie", "RVP_Vaccinatiegraad"),
                                   Name=c("Aantal gevaccineerde individuen", "Aantal individuen in de populatie", "Percentage gevaccineerden in de populatie"),
                                   Unit=c(rep("personen", 2), "percentage"),
                                   `Aggregation indicator`=NA,
                                   Formula=c(rep(NA, 2), "RVP_Gevaccineerden/RVP_Populatie*100"),
                                   `Data type`=c(rep("Numeric", 2), "Percentage"),
                                   Visible=c(rep(0, 2), 1),
                                   `Threshold value`=c(rep(NA, 2), 50),
                                   `Threshold Indicator`=c(rep(NA, 2), "RVP_Gevaccineerden"),
                                   Cube=1,
                                   Source="RIVM"))

saveWorkbook(wb, "Swing_vaccinatiegraad.xlsx", overwrite=T)
