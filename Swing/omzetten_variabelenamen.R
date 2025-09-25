library(tidyverse)
library(haven)
library(labelled)
library(this.path)

setwd(dirname(this.path()))

printf = function (...) cat(paste(sprintf(...),"\n"))

filenames = choose.files(caption="Selecteer bestanden...", filters=c("SPSS-bestand (*.sav)","*.sav"))

afkortingen = data.frame(zoekterm=c("E-MOVO", "Kindermonitor", "Jongvolwassenenmonitor", "Volwassenenmonitor"),
                         afkorting=c("J", "KM", "JV", "VO"))
for (filename in filenames) {
  map = dirname(filename)
  savname = str_sub(filename, start=str_length(map) + 2)
  
  afk = afkortingen$afkorting[str_which(map, afkortingen$zoekterm)]
  if (is_empty(afk)) {
    afk = readline(prompt="Pad niet leesbaar. Geef de gewenste afkorting op (J/KM/JV/VO): ")
  }
  
  printf("Verwerking bestand %s... (afkorting %s)", savname, afk)
  data = read_spss(filename, user_na=T)
  
  # gemeentenummers omzetten naar string met voorgaande nullen
  #gemeentekolommen = colnames(data)[str_which(colnames(data), regex("gemeente|gemcode", ignore_case=T))]
  #gemeentekolommen = gemeentekolommen[!str_detect(gemeentekolommen, regex("weeg|stratum|standaard|weging", ignore_case=T))]
  
  #for (gemeentekolom in gemeentekolommen) {
  #  if (!is.character(data[[gemeentekolom]])) {
  #    data[[paste0("Swing_", gemeentekolom)]] = sprintf("%04d", data[[gemeentekolom]])
  #  }
  #}
  
  # wijknummers omzetten naar string met voorgaande nullen
  if ("NOG_Wijk2020" %in% colnames(data)) {
    # moet zijn: wijk23
    data$NOG_Wijk2020 = sprintf("%06d", data$NOG_Wijk2020)
  }

  # V&O heeft wijk en gemeente apart; samenvoegen voor Swing
  if (afk == "VO") {
    if (is.character(data$Gemeentecode)) {
      data$Wijkcode_volledig = paste0(data$Gemeentecode, data$Wijkcode)
    } else {
      data$Wijkcode_volledig = sprintf("%04d%s", data$Gemeentecode, data$Wijkcode)
    }
    
    data$Wijkcode_volledig[str_length(data$Wijkcode_volledig) < 6] = NA
  }
  
  # woord 'opgeschoond' uit alle labels halen
  for (i in 1:ncol(data)) {
    if (is.null(var_label(data[[i]])))
      next
    
    huidig_label = var_label(data[[i]]) %>% str_replace(regex("( |).?opgeschoond.?", ignore_case=T), "")
    var_label(data[[i]]) = huidig_label
  }
  
  colnames(data) = paste0(afk, "_", colnames(data))
  write_sav(data, paste0(map, "/Swing-", savname))
  #write_sav(data, savname)
  printf("Bestand Swing-%s verwerkt en opgeslagen.", savname)
}
