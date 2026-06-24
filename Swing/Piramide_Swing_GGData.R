##
## Script voor het maken van een bevolkingspiramide 
## en het berekenen van de gemiddelde leeftijd voor in Swing
## 
## Het script geeft absolute en relatieve (%) aantallen 
## 
## Uitsplitsing op leeftijd (1 jaar, 5 jaren en 10 jaren)
## Uitsplitsing naar man/vrouw
##
## Instellen; jaar, missing value en GGD_code. Zie lijst met GGD-codes.
## 
## Gaat uit van de CBS cijfers op 1 januari en burgerlijke staat totaal
## Mocht het zo zijn dat er gemeenten zijn samengevoegd; zie samenvoegen in het script.
## 
## 
## GGD Zuid-Limburg
## 
## 

# Inladen benodigde libraries
library(openxlsx) 
library(cbsodataR)
library(tidyverse) 
# Geef hier de map waar de bestanden komen te staan.
setwd(paste0('C:/Users/', Sys.getenv('username'), '/Documents/'))

# Variabelen instellen; jaartal, codering missing values en GGD code
# Geef hier het gewenste jaartal, code voor missing values en GGD code
jaar = "2025"
missing_value <- -99996
GGD_code <- "GG6106"
# GGD codes:
# GG0111	GGD Groningen
# GG0706	GGD Drenthe
# GG1009	GGD IJsselland
# GG1106	GGD Twente
# GG1413	GGD Noord- en Oost-Gelderland
# GG1911	Veiligheids- en Gezondheidsregio Gelderland-Midden
# GG2014	GGD Gelderland-Zuid
# GG2209	GGD Flevoland
# GG2514	GGD Regio Utrecht
# GG2707	GGD Hollands-Noorden
# GG3109	GGD Kennemerland
# GG3406	GGD Amsterdam
# GG3606	GGD Gooi en Vechtstreek
# GG4506	GGD Hollands-Midden
# GG4607	GGD Rotterdam-Rijnmond
# GG4816	Dienst Gezondheid & Jeugd ZHZ
# GG5006	GGD Zeeland
# GG5206	GGD West-Brabant
# GG5406	GGD Hart voor Brabant
# GG5608	GGD Brabant-Zuidoost
# GG6011	GGD Limburg-Noord
# GG6106	GGD Zuid-Limburg
# GG7014	GGD Haaglanden
# GG7206	GGD Fryslân
# GG7306	GGD Zaanstreek/Waterland


# Selecteer gemeenten binnen GGD regio
Gebieden_NL <- cbsodataR::cbs_get_data("85755NED", catalog="CBS") %>% cbs_add_label_columns()
Gebieden_NL$Code_16 <- gsub(" ", "", Gebieden_NL$Code_16)
regio_ggd <- Gebieden_NL %>%
  filter(Code_16 == GGD_code) %>% # Code 14 naar code 16 aangepast
  select(c(RegioS)) %>%
  unlist()
regio_ggd <- gsub(" ", "", regio_ggd)

# Alternatief; gemeentecodes inkloppen in lijst
#regio_ggd <- c("GM0888","GM1954","GM0899","GM1903","GM1729","GM0917","GM0928","GM0882","GM0935","GM0938","GM0965","GM1883","GM0971","GM0981","GM0994","GM0986")

Gebieden_GGD <- Gebieden_NL %>%
  filter(Code_16 == GGD_code) %>%
  select(Code_16, Naam_17, RegioS, RegioS_label)

# Cijfers importeren en omzetten naar Swingnamen ----
# meta <- cbs_get_meta("03759ned", catalog = "CBS")
# Haal tabel op met bevolkingscijfers op 1 januari. Alleen burgerlijke staat totaal (geen uitsplitsingen)
bevolking_1jan_GGD <- cbsodataR::cbs_get_data("03759ned", catalog = "CBS", Perioden=paste(jaar,"JJ00",sep=""), BurgerlijkeStaat = "T001019", 
                                   RegioS = regio_ggd)


# Haal bij perioden JJ00 weg en zet kolommen goed voor Swing, filter voor jaar en haal ongebruikte kolommen weg
bevolking_1jan_GGD <- bevolking_1jan_GGD %>%
  mutate(Perioden = gsub("JJ00","",Perioden)) %>%
  filter(Perioden == jaar) %>%
  rename(dnc_gender = Geslacht) %>%
  rename(dnc_age = Leeftijd) %>%
  rename(geoitem = RegioS) %>%
  rename(periode = Perioden) %>%
  rename(kb_bev_1jan = BevolkingOp1Januari_1) %>%
  select(!"GemiddeldeBevolking_2") %>%
  select(!"BurgerlijkeStaat")

## Vangnet om cijfers op te tellen wanneer gemeenten zijn samengevoegd ----
# Regio afhankelijk!
# Beekdaelen vanaf 2019 samengevoegd. Was Nuth (0951), Onderbanken (0881) en Schinnen (0962).
# Optioneel: Eijsden-Margraten (1903) vanaf 2011 samengevoegd. Was Eijsden () en Margraten ().
# Optioneel: Sittard-Geleen vanaf 2001 samengevoegd. Was Sittard (), Geleen () en Born ().
if (jaar < 2019) {
  bevolking_1jan_1954 <- cbsodataR::cbs_get_data("03759ned", catalog = "CBS", period=jaar, BurgerlijkeStaat = "T001019", 
                                      RegioS = c("GM0951","GM0881","GM0962"))
  # optellen Nuth (0951), Onderbanken (0881) en Schinnen (0962).
  bevolking_1jan_1954 <- bevolking_1jan_1954 %>%
    mutate(Perioden = gsub("JJ00","",Perioden)) %>%
    filter(Perioden == jaar) %>%
    rename(dnc_gender = Geslacht) %>%
    rename(dnc_age = Leeftijd) %>%
    rename(geoitem = RegioS) %>%
    rename(periode = Perioden) %>%
    rename(kb_bev_1jan = BevolkingOp1Januari_1) %>%
    select(!"GemiddeldeBevolking_2") %>%
    select(!"BurgerlijkeStaat") %>%
    group_by(dnc_age, dnc_gender) %>%
    mutate(kb_bev_1jan = sum(kb_bev_1jan)) %>%
    filter(geoitem == "GM0951") %>%
    mutate(geoitem = case_when(geoitem == "GM0951" ~ "GM1954", TRUE ~ geoitem))
  # Voeg samen en vervang oude GM1954 cijfers
  bevolking_1jan_GGD <- rbind(bevolking_1jan_GGD,bevolking_1jan_1954)
  bevolking_1jan_GGD <- bevolking_1jan_GGD[complete.cases(bevolking_1jan_GGD),]

  
} else {
  # Niet optellen
}
## CHECK aantal geoitems
#length(table(bevolking_1jan_GGD$geoitem))



# Zet codering man (3000), vrouw (4000) en totaal (T001038) goed
bevolking_1jan_GGD$dnc_gender[bevolking_1jan_GGD$dnc_gender == "3000   "] = "m"
bevolking_1jan_GGD$dnc_gender[bevolking_1jan_GGD$dnc_gender == "4000   "] = "v"
bevolking_1jan_GGD$dnc_gender[bevolking_1jan_GGD$dnc_gender == "T001038"] = "t"
bevolking_1jan_GGD$dnc_age <- as.numeric(bevolking_1jan_GGD$dnc_age)

# Leefijd per 5 optellen
# Totaal = 10000, 0j=10010, 1j=10100, 2j=10200, etc  10j=11000, 99j=19900, 100j=19901
# 95+: 22000 105+: 22300

# Functie om leeftijdscodes om te zetten naar leeftijdsgroepen
leeftijd_omzetten <- function(age_code) {
  if (age_code >= 10010 & age_code <= 19900) {
    return(round((age_code - 10000) / 100))
  } else if (age_code >= 19901 & age_code < 22000) {
    return(round((age_code - 19900) + 99))
  } else if (age_code == 22000) {
    return(95000)
  } else if (age_code == 22300) {
    return(105000)
  } else {
    return(NA)
  }
}

# Pas de functie toe op de leeftijdscodes
bevolking_1jan_GGD <- bevolking_1jan_GGD %>%
  filter(!dnc_age==10000)  %>% # Totaal eruit halen
  mutate(Age = sapply(dnc_age, leeftijd_omzetten)) 

# Bereken gemiddelde leeftijd per gemeente
# gemlftt
# Beschrijving
# De gemiddelde leeftijd is een rekenkundig gemiddelde over 1-jaarsleeftijdklassen, van 0- tot 95-jarigen.
# 
# rekenvoorbeeld:
#   Er zijn 300 0-jarigen en 100 1-jarigen. In totaal zijn er daarmee 400 personen.
# 
# Gemiddeld zijn de 0-jarigen 0,5 jaar oud (want 0 jaar op 1 januari en 1 jaar op 31 december). 
# idem zijn de 1-jarigen gemiddeld 1,5 jaar oud, etc..
# 
# Deze gemiddelde leeftijd is vermenigvuldigd met het aantal personen dat deze gemiddelde leeftijd bezit:
#   Voor de 0-jarigen komt de cumulatieve leeftijd op 300 personen * 0,5 jaar = 150.
#   voor de 1-jarigen komt de cumulatieve leeftijd op 100 personen * 1,5 jaar = 150.
#   In totaal is de cumulatieve leeftijd daarmee 300.
# 
# De gemiddelde leeftijd is dan:
#   De cumulatieve leeftijd (300) gedeeld door het totaal aantal personen (400) = 0,75.

gemlft_GGD <- bevolking_1jan_GGD %>%
  filter(dnc_gender == "t") %>%
  filter(!dnc_age == 22300) %>%
  filter(!dnc_age %in% (19500:19905) ) %>%
  mutate(Age1 = (Age + 0.5)) %>%
  mutate(Age1 = case_when(Age==95000 ~ 95.5, TRUE ~ Age1)) %>%
  mutate(Age2 =  (Age1*kb_bev_1jan)) %>%
  group_by(geoitem) %>%
  summarise(gemlftt = sum((Age2) / sum(kb_bev_1jan))) %>%
  mutate(across(c('geoitem'), substr, 3, nchar(geoitem)))
# Voeg kolom periode en geolevel toe
gemlft_GGD <- cbind(geolevel = "gemeente", gemlft_GGD)
gemlft_GGD <- cbind(period = jaar, gemlft_GGD)
# Maak geoitems numeriek om de nul voor de gemeentecode weg te halen
gemlft_GGD$geoitem <- as.numeric(gemlft_GGD$geoitem)

# Maak leeftijdsgroepen per 5 jaar
bevolking_1jan_GGD <- bevolking_1jan_GGD %>%
  mutate(AgeGroup5 = cut(Age, breaks = seq(0, 110, by = 5), right = FALSE, 
                         labels = sprintf("bev%02d%02d", seq(0, 105, by = 5), seq(4, 109, by = 5))))

# Maak leeftijdsgroepen per 10 jaar
bevolking_1jan_GGD <- bevolking_1jan_GGD %>%
  mutate(AgeGroup10 = cut(Age, breaks = seq(0, 110, by = 10), right = FALSE, 
                          labels = sprintf("bev%02d%02d", seq(0, 100, by = 10), seq(9, 109, by = 10))))

# Maak categorie 95+ en haal 95-99, 100-104 en 105+ eruit en voeg geolevel gemeente toe
bevolking_1jan_GGD <- bevolking_1jan_GGD %>%
  mutate(AgeGroup5 = case_when(Age == 105000 ~ "bev105plus", TRUE ~ AgeGroup5)) %>%
  mutate(AgeGroup5 = case_when(Age == 95000 ~ "bev95plus", TRUE ~ AgeGroup5)) %>%
  mutate(AgeGroup10 = case_when(Age == 105000 ~ "bev105plus", TRUE ~ AgeGroup10)) %>%
  mutate(AgeGroup10 = case_when(Age == 95000 ~ "bev95plus", TRUE ~ AgeGroup10)) %>%
  filter(!AgeGroup5=="bev9599" & !AgeGroup5=="bev100104" & !AgeGroup5== "bev105plus") %>%
  mutate(Geolevel = "gemeente")

# Totaal geslacht eruit halen en de kolommen leeftijdsgroep 5+10. GM weghalen bij geoitem
bevolking_1jan_GGD <- bevolking_1jan_GGD %>%
  select(!Age & !AgeGroup5 & !AgeGroup10) %>%
  filter(!dnc_gender == "t") %>%
  mutate(across(c('geoitem'), substr, 3, nchar(geoitem)))

# Maak geoitems numeriek om de nul voor de gemeentecode weg te halen
bevolking_1jan_GGD$geoitem <- as.numeric(bevolking_1jan_GGD$geoitem)

# Indien missende waarden, vervangen door missing value code
colSums(is.na(bevolking_1jan_GGD))
bevolking_1jan_GGD[is.na(bevolking_1jan_GGD)] <- missing_value

## Maak het Excel bestand voor Swing import ----
# Crtl + Alt + T om deze sectie te draaien
workbook <- openxlsx::createWorkbook()
# Add data to workbook
openxlsx::addWorksheet(workbook, sheetName = "Data")
openxlsx::writeData(workbook,"Data",bevolking_1jan_GGD)
# Add data to workbook
openxlsx::addWorksheet(workbook, sheetName = "Data2")
openxlsx::writeData(workbook,"Data2",gemlft_GGD)

# Add Data_def col and type
openxlsx::addWorksheet(workbook, sheetName = "Data_def")
openxlsx::writeData(workbook, "Data_def",
                    cbind("col" = c("dnc_gender","dnc_age","geoitem", "periode","kb_bev_1jan","Geolevel"),
                          "type"= c("dim","dim","geoitem","period","var","geolevel")))

# Add dimensions_code - Dimension code	Name
openxlsx::addWorksheet(workbook, sheetName = "1Dimensies")
openxlsx::writeData(workbook, "1Dimensies",
                    cbind("Dimension code" = c("dc_gender","dc_age"),
                          "Name"= c("Geslacht","Leeftijd")))

# Add dimensions_levels - Dimlevel code	Name	Dimension code	AggregateType
openxlsx::addWorksheet(workbook, sheetName = "2Dimensieniveaus")
openxlsx::writeData(workbook, "2Dimensieniveaus",
                    cbind("Dimlevel code" = c("dnc_gender","dnc_age","dnc_age_10","dnc_age_5","dnc_age_doel"),
                          "Name"= c("Geslacht","Leeftijd","Leeftijd per 10 jaar","Leeftijd per 5 jaar","Leeftijd per doelgroep"),
                          "Dimension code" = c("dc_gender","dc_age","dc_age","dc_age","dc_age"),
                          "AggregateType" = c("Unknown","Unknown","Unknown","Unknown","Unknown")))

# Add Dim_gender - Itemcode	Naam	SequenceNr
openxlsx::addWorksheet(workbook, sheetName = "dnc_gender")
openxlsx::writeData(workbook, "dnc_gender",
                    cbind("Item code" = c("v","m"),
                          "Name"= c("Vrouw","Man"),
                          "SequenceNr" = c("1","2")))

# Add Dim_age_1 - Itemcode	Naam	SequenceNr
openxlsx::addWorksheet(workbook, sheetName = "dnc_age")
openxlsx::writeData(workbook, "dnc_age",
                    cbind("Item code" = c("10010","10100","10200","10300","10400","10500","10600","10700","10800","10900",
                                          "11000","11100","11200","11300","11400","11500","11600","11700","11800","11900",
                                          "12000","12100","12200","12300","12400","12500","12600","12700","12800","12900",
                                          "13000","13100","13200","13300","13400","13500","13600","13700","13800","13900",
                                          "14000","14100","14200","14300","14400","14500","14600","14700","14800","14900",
                                          "15000","15100","15200","15300","15400","15500","15600","15700","15800","15900",
                                          "16000","16100","16200","16300","16400","16500","16600","16700","16800","16900",
                                          "17000","17100","17200","17300","17400","17500","17600","17700","17800","17900",
                                          "18000","18100","18200","18300","18400","18500","18600","18700","18800","18900",
                                          "19000","19100","19200","19300","19400","22000"),
                          "Name"= c("0 jaar","1 jaar","2 jaar","3 jaar","4 jaar","5 jaar","6 jaar","7 jaar","8 jaar","9 jaar",
                                    "10 jaar","11 jaar","12 jaar","13 jaar","14 jaar","15 jaar","16 jaar","17 jaar","18 jaar","19 jaar",
                                    "20 jaar","21 jaar","22 jaar","23 jaar","24 jaar","25 jaar","26 jaar","27 jaar","28 jaar","29 jaar",
                                    "30 jaar","31 jaar","32 jaar","33 jaar","34 jaar","35 jaar","36 jaar","37 jaar","38 jaar","39 jaar",
                                    "40 jaar","41 jaar","42 jaar","43 jaar","44 jaar","45 jaar","46 jaar","47 jaar","48 jaar","49 jaar",
                                    "50 jaar","51 jaar","52 jaar","53 jaar","54 jaar","55 jaar","56 jaar","57 jaar","58 jaar","59 jaar",
                                    "60 jaar","61 jaar","62 jaar","63 jaar","64 jaar","65 jaar","66 jaar","67 jaar","68 jaar","69 jaar",
                                    "70 jaar","71 jaar","72 jaar","73 jaar","74 jaar","75 jaar","76 jaar","77 jaar","78 jaar","79 jaar",
                                    "80 jaar","81 jaar","82 jaar","83 jaar","84 jaar","85 jaar","86 jaar","87 jaar","88 jaar","89 jaar",
                                    "90 jaar","91 jaar","92 jaar","93 jaar","94 jaar","95 jaar en ouder"),
                          
                          "SequenceNr" = c("1","2","3","4","5","6","7","8","9","10",
                                           "11","12","13","14","15","16","17","18","19",
                                           "20","21","22","23","24","25","26","27","28","29",
                                           "30","31","32","33","34","35","36","37","38","39",
                                           "40","41","42","43","44","45","46","47","48","49",
                                           "50","51","52","53","54","55","56","57","58","59",
                                           "60","61","62","63","64","65","66","67","68","69",
                                           "70","71","72","73","74","75","76","77","78","79",
                                           "80","81","82","83","84","85","86","87","88","89",
                                           "90","91","92","93","94","95","96")))

# Add Dim_age_5 - Itemcode	Naam SequenceNr
openxlsx::addWorksheet(workbook, sheetName = "dnc_age_5")
openxlsx::writeData(workbook, "dnc_age_5",
                    cbind("Item code" = c("bev0004","bev0509","bev1014","bev1519","bev2024","bev2529","bev3034","bev3539","bev4044",
                                          "bev4549","bev5054","bev5559","bev6064","bev6569","bev7074","bev7579","bev8084","bev8589",
                                          "bev9094","bev95plus"),
                          "Name"= c("4 jaar en jonger", "5-9 jaar","10-14 jaar","15-19 jaar","20-24 jaar","25-29 jaar","30-34 jaar","35-39 jaar",
                                    "40-44 jaar","45-49 jaar","50-54 jaar","55-59 jaar","60-64 jaar","65-69 jaar","70-74 jaar","75-79 jaar",
                                    "80-84 jaar","85-89 jaar","90-94 jaar","95 jaar en ouder"),
                          
                          "SequenceNr" = c("1","2","3","4","5","6","7","8","9","10",
                                           "11","12","13","14","15","16","17","18","19","20")))


# Add Dim_age_10 - Itemcode	Naam SequenceNr
openxlsx::addWorksheet(workbook, sheetName = "dnc_age_10")
openxlsx::writeData(workbook, "dnc_age_10",
                    cbind("Item code" = c("bev0009","bev1019","bev2029","bev3039","bev4049","bev5059","bev6069","bev7079",
                                          "bev8089","bev90plus"),
                          "Name"= c("9 jaar en jonger","10-19 jaar","20-29 jaar","30-39 jaar","40-49 jaar","50-59 jaar",
                                    "60-69 jaar","70-79 jaar","80-89 jaar","90 jaar en ouder"),
                          
                          "SequenceNr" = c("1","2","3","4","5","6","7","8","9","10")))

# Add dnc_age_doel - Itemcode	Naam SequenceNr
openxlsx::addWorksheet(workbook, sheetName = "dnc_age_doel")
openxlsx::writeData(workbook, "dnc_age_doel",
                    cbind("Item code" = c("bev0003","bev0411","bev1217","bev1824","bev2539","bev4054",
                                          "bev5564","bev6574","bev7584","bev85plus"),
                          "Name"= c("3 jaar en jonger","4-11 jaar","12-17 jaar","18-24 jaar","25-39 jaar","40-54 jaar",
                                    "55-64 jaar","65-74 jaar","75-84 jaar","85 jaar en ouder"),
                          "SequenceNr" = c("1","2","3","4","5","6","7","8","9","10")))



# Add aggregatie op 5 tallen
openxlsx::addWorksheet(workbook, sheetName = "aggregatie_5")
openxlsx::writeData(workbook, "aggregatie_5",
                    cbind("dnc_age" = c("10010","10100","10200","10300","10400","10500","10600","10700","10800","10900",
                                        "11000","11100","11200","11300","11400","11500","11600","11700","11800","11900",
                                        "12000","12100","12200","12300","12400","12500","12600","12700","12800","12900",
                                        "13000","13100","13200","13300","13400","13500","13600","13700","13800","13900",
                                        "14000","14100","14200","14300","14400","14500","14600","14700","14800","14900",
                                        "15000","15100","15200","15300","15400","15500","15600","15700","15800","15900",
                                        "16000","16100","16200","16300","16400","16500","16600","16700","16800","16900",
                                        "17000","17100","17200","17300","17400","17500","17600","17700","17800","17900",
                                        "18000","18100","18200","18300","18400","18500","18600","18700","18800","18900",
                                        "19000","19100","19200","19300","19400","22000"),
                          "dnc_age_5"= c("bev0004","bev0004","bev0004","bev0004","bev0004","bev0509","bev0509","bev0509","bev0509","bev0509",
                                         "bev1014","bev1014","bev1014","bev1014","bev1014","bev1519","bev1519","bev1519","bev1519","bev1519",
                                         "bev2024","bev2024","bev2024","bev2024","bev2024","bev2529","bev2529","bev2529","bev2529","bev2529",
                                         "bev3034","bev3034","bev3034","bev3034","bev3034","bev3539","bev3539","bev3539","bev3539","bev3539",
                                         "bev4044","bev4044","bev4044","bev4044","bev4044","bev4549","bev4549","bev4549","bev4549","bev4549",
                                         "bev5054","bev5054","bev5054","bev5054","bev5054","bev5559","bev5559","bev5559","bev5559","bev5559",
                                         "bev6064","bev6064","bev6064","bev6064","bev6064","bev6569","bev6569","bev6569","bev6569","bev6569",
                                         "bev7074","bev7074","bev7074","bev7074","bev7074","bev7579","bev7579","bev7579","bev7579","bev7579",
                                         "bev8084","bev8084","bev8084","bev8084","bev8084","bev8589","bev8589","bev8589","bev8589","bev8589",
                                         "bev9094","bev9094","bev9094","bev9094","bev9094","bev95plus")))


# Add aggregatie op 10 tallen
openxlsx::addWorksheet(workbook, sheetName = "aggregatie_10")
openxlsx::writeData(workbook, "aggregatie_10",
                    cbind("dnc_age" = c("10010","10100","10200","10300","10400","10500","10600","10700","10800","10900",
                                        "11000","11100","11200","11300","11400","11500","11600","11700","11800","11900",
                                        "12000","12100","12200","12300","12400","12500","12600","12700","12800","12900",
                                        "13000","13100","13200","13300","13400","13500","13600","13700","13800","13900",
                                        "14000","14100","14200","14300","14400","14500","14600","14700","14800","14900",
                                        "15000","15100","15200","15300","15400","15500","15600","15700","15800","15900",
                                        "16000","16100","16200","16300","16400","16500","16600","16700","16800","16900",
                                        "17000","17100","17200","17300","17400","17500","17600","17700","17800","17900",
                                        "18000","18100","18200","18300","18400","18500","18600","18700","18800","18900",
                                        "19000","19100","19200","19300","19400","22000"),
                          "dnc_age_10"= c("bev0009","bev0009","bev0009","bev0009","bev0009","bev0009","bev0009","bev0009","bev0009","bev0009",
                                          "bev1019","bev1019","bev1019","bev1019","bev1019","bev1019","bev1019","bev1019","bev1019","bev1019",
                                          "bev2029","bev2029","bev2029","bev2029","bev2029","bev2029","bev2029","bev2029","bev2029","bev2029",
                                          "bev3039","bev3039","bev3039","bev3039","bev3039","bev3039","bev3039","bev3039","bev3039","bev3039",
                                          "bev4049","bev4049","bev4049","bev4049","bev4049","bev4049","bev4049","bev4049","bev4049","bev4049",
                                          "bev5059","bev5059","bev5059","bev5059","bev5059","bev5059","bev5059","bev5059","bev5059","bev5059",
                                          "bev6069","bev6069","bev6069","bev6069","bev6069","bev6069","bev6069","bev6069","bev6069","bev6069",
                                          "bev7079","bev7079","bev7079","bev7079","bev7079","bev7079","bev7079","bev7079","bev7079","bev7079",
                                          "bev8089","bev8089","bev8089","bev8089","bev8089","bev8089","bev8089","bev8089","bev8089","bev8089",
                                          "bev90plus","bev90plus","bev90plus","bev90plus","bev90plus","bev90plus")))

# Add aggregatie op doelgroepen (in 10)
openxlsx::addWorksheet(workbook, sheetName = "aggregatie_doel")
openxlsx::writeData(workbook, "aggregatie_doel",
                    cbind("dnc_age" = c("10010","10100","10200","10300","10400","10500","10600","10700","10800","10900",
                                        "11000","11100","11200","11300","11400","11500","11600","11700","11800","11900",
                                        "12000","12100","12200","12300","12400","12500","12600","12700","12800","12900",
                                        "13000","13100","13200","13300","13400","13500","13600","13700","13800","13900",
                                        "14000","14100","14200","14300","14400","14500","14600","14700","14800","14900",
                                        "15000","15100","15200","15300","15400","15500","15600","15700","15800","15900",
                                        "16000","16100","16200","16300","16400","16500","16600","16700","16800","16900",
                                        "17000","17100","17200","17300","17400","17500","17600","17700","17800","17900",
                                        "18000","18100","18200","18300","18400","18500","18600","18700","18800","18900",
                                        "19000","19100","19200","19300","19400","22000"),
                          "dnc_age_doel"= c("bev0003","bev0003","bev0003","bev0003","bev0411","bev0411","bev0411","bev0411","bev0411","bev0411",
                                          "bev0411","bev0411","bev1217","bev1217","bev1217","bev1217","bev1217","bev1217","bev1824","bev1824",
                                          "bev1824","bev1824","bev1824","bev1824","bev1824","bev2539","bev2539","bev2539","bev2539","bev2539",
                                          "bev2539","bev2539","bev2539","bev2539","bev2539","bev2539","bev2539","bev2539","bev2539","bev2539",
                                          "bev4054","bev4054","bev4054","bev4054","bev4054","bev4054","bev4054","bev4054","bev4054","bev4054",
                                          "bev4054","bev4054","bev4054","bev4054","bev4054","bev5564","bev5564","bev5564","bev5564","bev5564",
                                          "bev5564","bev5564","bev5564","bev5564","bev5564","bev6574","bev6574","bev6574","bev6574","bev6574",
                                          "bev6574","bev6574","bev6574","bev6574","bev6574","bev7584","bev7584","bev7584","bev7584","bev7584",
                                          "bev7584","bev7584","bev7584","bev7584","bev7584","bev85plus","bev85plus","bev85plus","bev85plus","bev85plus",
                                          "bev85plus","bev85plus","bev85plus","bev85plus","bev85plus","bev85plus")))


# Add Indicators; 
openxlsx::addWorksheet(workbook, sheetName = "Indicators")
openxlsx::writeData(workbook, "Indicators",
                    
                    cbind("Indicator code" = 
                            #All variable-levels
                            c("kb_bev_1jan","gemlftt"),
                          
                          "Name" = 
                            #All variable-levels
                            c("Bevolking op 1 januari","Leeftijd bevolking"), 
                          
                          #aantal rijen voor personen = levels variabele + 2(gewogen/ongewogen)
                          "Unit" = 
                            c("number","jaar"),
                          
                          #Zelfde logica als "Unit"
                          "Data type" = c("Numeric","Mean"),
                          
                          # Make it visible or not.
                          "Visible" = c(1,1),
                          
                          # Cube data
                          "Cube" = c(1,0),
                          
                          # RoundOff op  1
                          "RoundOff" = c("1","0.1"),
                          
                          # Formula for proportion calculation
                          #"Formula" = c("","Inhabitants"),
                          
                          # Source is overal hetzelfde
                          "Source" = c("CBS","CBS"),
                          
                          # Aggregation indicator
                          "Aggregation indicator" = c("","bevtot"),
                          
                          # Aggregate geoitems
                          "Aggregate geoitems" = c("","1"),
                          
                          # Aggregate periods
                          "Aggregate periods" = c("","1"),
                          
                          # Description 
                          "Description" = c("In de bevolkingsaantallen zijn uitsluitend personen begrepen die zijn opgenomen in het bevolkingsregister van een Nederlandse gemeente. In principe wordt iedereen die voor onbepaalde tijd in Nederland woont, opgenomen in het bevolkingsregister van de woongemeente.",
                                            "De gemiddelde leeftijd is een rekenkundig gemiddelde over 1-jaarsleeftijdklassen, van 0- tot 95-jarigen.")
                    )
)
openxlsx::saveWorkbook(workbook, file = glue::glue("{Sys.Date()}_CBS_piramide_{GGD_code}_{jaar}.xlsx"), overwrite = TRUE)
system2("open", glue::glue("{Sys.Date()}_CBS_piramide_{GGD_code}_{jaar}.xlsx"))


# Aanmaken percentages ----
bevolking_p_1jan_GGD <- bevolking_1jan_GGD %>%
  group_by(geoitem) %>%
  mutate(kb_p_bev_1jan = kb_bev_1jan/(sum(kb_bev_1jan))*100) %>%
  select(!kb_bev_1jan)

# Controle voor een geoitem, is totaal 100%
bevolking_p_1jan_GGD_test <- bevolking_p_1jan_GGD %>%
  filter(geoitem == 1954)
sum(bevolking_p_1jan_GGD_test$kb_p_bev_1jan)


## Maak het Excel bestand voor Swing import
# Crtl + Alt + T to run section
workbook <- openxlsx::createWorkbook()
# Add data to workbook
openxlsx::addWorksheet(workbook, sheetName = "Data")
openxlsx::writeData(workbook,"Data",bevolking_p_1jan_GGD)

# Add Data_def col and type
openxlsx::addWorksheet(workbook, sheetName = "Data_def")
openxlsx::writeData(workbook, "Data_def",
                    cbind("col" = c("dnc_gender","dnc_age","geoitem", "periode","kb_p_bev_1jan","Geolevel"),
                          "type"= c("dim","dim","geoitem","period","var","geolevel")))

# Add Indicators; 
openxlsx::addWorksheet(workbook, sheetName = "Indicators")
openxlsx::writeData(workbook, "Indicators",
                    
                    cbind("Indicator code" = 
                            #All variable-levels
                            c("kb_p_bev_1jan"),
                          
                          "Name" = 
                            #All variable-levels
                            c("Bevolking op 1 januari"), 
                          
                          #aantal rijen voor personen = levels variabele + 2(gewogen/ongewogen)
                          "Unit" = 
                            c("p"),
                          
                          #Zelfde logica als "Unit"
                          "Data type" = c("Percentage (sum)"),
                          
                          # Make it visible or not.
                          "Visible" = c(1),
 
                          # Aggregation indicator
                          "Aggregation indicator" = c("bevtot"),
                          
                          # Cube data
                          "Cube" = c(1),
                          
                          # RoundOff op  1
                          "RoundOff" = c("0.1"),
                          
                          # Formula for proportion calculation
                          #"Formula" = c("","Inhabitants"),
                          
                          # Source is overal hetzelfde
                          "Source" = c("CBS"),
                          
                          # Description 
                          "Description" = c("In de bevolkingsaantallen zijn uitsluitend personen begrepen die zijn opgenomen in het bevolkingsregister van een Nederlandse gemeente. In principe wordt iedereen die voor onbepaalde tijd in Nederland woont, opgenomen in het bevolkingsregister van de woongemeente.")
                    )
)
openxlsx::saveWorkbook(workbook, file = glue::glue("{Sys.Date()}_P_CBS_piramide_{GGD_code}_{jaar}.xlsx"), overwrite = TRUE)
#system2("open", glue::glue("{Sys.Date()}_P_CBS_piramide_{GGD_code}_{jaar}.xlsx"))
