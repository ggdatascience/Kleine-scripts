################################################################################
##
## Enalyzer exporteert data soms een beetje maf. Dit gebeurt enkel op basis van
## nummering, waardoor een omgedraaide vraag in een andere vragenlijst puur
## huilen wordt in SPSS. Dit script combineert meerdere datasets op basis van
## de vraagstellingen. 
##
################################################################################

# benodigde packages, installeren indien niet aanwezig
pkg_nodig = c("tidyverse", "haven", "labelled", "openxlsx")

for (pkg in pkg_nodig) {
  if (system.file(package = pkg) == "") {
    install.packages(pkg)
  }
}

library(tidyverse)
library(haven)
library(labelled)
library(openxlsx)


# basisdingetjes
printf = function (...) cat(paste(sprintf(...),"\n"))

rename_vars = F
ignore_vars = c()
searches = c()

config.file = choose.files(caption="Selecteer configuratiebestand", multi=F, filters=c("Excel-bestand (*.xlsx)","*.xlsx"))
if (str_length(config.file) <= 1) {
  warning("Geen configuratiebestand geselecteerd. Standaardinstellingen aangenomen: variabelen niet hernoemen.")
} else {
  config = read.xlsx(config.file, sheet="algemeen")
  if (!exists("config")) stop("Configuratiebestand kon niet gelezen worden. Wellicht is deze nog geopend in Excel?")
  rename_vars = config$hernoem_variabelen
  ignore_vars = config$negeer_variabelen
  
  config = read.xlsx(config.file, sheet="zoekwaarden")
  searches = config %>% rename("search"="zoek", "name"="naam", "notes"="notitie")
  searches$name = str_replace_all(searches$name, "\\s+", ".")
}

data.files = choose.files(caption="Selecteer databestanden", multi=T, filters=c("Excel-bestand (*.xlsx)","*.xlsx"))

# bind_rows doet raar met var_labels, dus die moeten we apart bewaren
labels.vars = c()
labels.labs = c()
rm("data.combined")
for (d in 1:length(data.files)) {
  dname = basename(data.files[d])
  printf("Verwerk data uit dataset %s...", dname)
  
  # de sheets heten meestal nl en Data, maar dat risico gaan we natuurlijk niet nemen
  data.overview = read.xlsx(data.files[d], sheet=1)
  if (!exists("data.overview")) stop("Databestand kon niet gelezen worden. Wellicht is deze nog geopend in Excel?")
  data = read.xlsx(data.files[d], sheet=2, detectDates=T)
  
  # eerst de vragen en antwoordmogelijkheden verwerken
  QandA = data.frame(matrix(nrow = nrow(data.overview), ncol=4))
  colnames(QandA) = c("code", "format", "question", "answers")
  # kolom 1 is de 'variabele', kolom 2 de vraag in menselijke tekst, kolom 3 het type,
  # kolommen 4:einde de antwoordmogelijheden
  questions = c()
  for (i in 1:nrow(data.overview)) {
    question = data.overview[i,2]
    answers = c()
    
    for (j in 4:ncol(data.overview)) {
      answer = data.overview[i,j]
      if (is.null(answer) || is.na(answer)) next
      
      # lelijke workaround, omdat een dynamische waarde niet kan
      answers = eval(parse(text=sprintf("c(answers, \"%s\" = %s)", answer, names(data.overview)[j])))
    }
    
    # toevoegen aan overzicht
    QandA[i,"code"] = data.overview[i,1]
    QandA[i,"format"] = data.overview[i,3]
    QandA[i,"question"] = question
    if (length(answers) > 0) QandA$answers[[i]] = list(answers)
  }
  
  data.clean = data %>%
    rename(Starttijd=Starttijd.enquête) %>%
    mutate(Bron=dname)
  
  toremove = c()
  for (i in 1:ncol(data.clean)) {
    name = colnames(data.clean)[i]
    if (is.null(name) || is.na(name)) next
    
    # sommige variabelen zijn onverklaarbaar volledig leeg, die verwijderen we
    if (all(is.na(data.clean[,i]))) {
      toremove = c(toremove, i)
      next
    }
    
    # alleen genummerde vragen verwerken
    if (!str_detect(name, "^[0-9]+")) next
    
    QandA_index = which(QandA$code == name)
    
    # labels toevoegen
    label = str_replace_all(QandA$question[QandA_index], "\\s+", " ")
    var_label(data.clean[,i]) = label
    if (!is.na(QandA$answers[QandA_index])) {
      val_labels(data.clean[,i]) = unlist(QandA$answers[[QandA_index]])
    }
    
    # volgende probleem: is het een losse vraag, of een vraag met meerdere antwoorden?
    # (deze zijn te herkennen aan patroon 7_2)
    prefix = "EN_"
    if (str_detect(name, "[0-9]+_[0-9]+")) {
      # vraag met subantwoorden: vraag toevoegen aan label
      match = str_match(name, "([0-9]+)_([0-9]+)")
      var_label(data.clean[,i]) = str_replace_all(paste(QandA$question[QandA$code == match[2]], "-", QandA$question[QandA_index]), "\\s+", " ")
      prefix = paste0(prefix, match[3], "_")
    } # geen else nodig
    
    # voor nu hernoemen we de kolommen naar de vraag, met een herkenningsteken voor later
    label.clean = str_replace_all(var_label(data.clean[,i]), "[^a-zA-Z]*", "")
    if (str_length(label.clean) > 50) {
      colnames(data.clean)[i] = paste0(prefix, str_sub(label.clean, end=25), "_", str_sub(label.clean, start=-25))
    } else {
      colnames(data.clean)[i] = paste0(prefix, label.clean)
    }
    
    # moeten we een kolom niet samenvoegen met een andere set? .[num] erachter
    if (!is.na(ignore_vars) && length(ignore_vars) > 0 && any(str_detect(label, ignore_vars))) {
      colnames(data.clean)[i] = paste0(colnames(data.clean)[i], ".", d)
    }
    
    # het kan voorkomen dat er meerdere variabelen zijn met exact dezelfde vraag, bijvoorbeeld als een open vraag
    # niet los gespecificeerd is in de vraag
    indexes = grep(colnames(data.clean)[i], colnames(data.clean), fixed=T)
    if (length(indexes) >= 1) {
      colnames(data.clean)[i] = paste0(colnames(data.clean)[i], length(indexes))
    }
    
    # bewaren voor later
    labels.vars = c(labels.vars, colnames(data.clean)[i])
    labels.labs = c(labels.labs, var_label(data.clean[,i]))
  }
  
  # verwijder volledig lege kolommen
  if (length(toremove) > 0) {
    data.clean = data.clean[,-toremove]
  }
  
  # samenvoegen met eerdere datasets
  if (exists("data.combined")) {
    data.combined = bind_rows(data.combined, data.clean)
  }
  else {
    data.combined = data.clean
  }
}

for (i in 1:length(labels.vars)) {
  var_label(data.combined[,labels.vars[i]]) = labels.labs[i]
}

# nu alle kolommen doorlopen en juiste variabelenaam geven
j = 1
for (i in 1:ncol(data.combined)) {
  name = colnames(data.combined)[i]
  
  if (!str_detect(name, "^EN_")) next
  
  if (rename_vars == "zoek" && length(searches$search) > 0) {
    keyword = str_detect(var_label(data.combined[,i]), regex(searches$search, ignore_case=T))
    if (any(keyword)) {
      if (sum(keyword, na.rm=T) > 1) {
        printf("Matches voor %s:", name)
        print(which(keyword))
      }
      # de vraag bevat 1 van de keywords; hernoem kolom
      colnames(data.combined)[i] = searches$name[keyword]
      searches$question[keyword] = var_label(data.combined[,i])
    } else {
      # geen vervangende kolomnaam gevonden? dan maar de vraag
      colnames(data.combined)[i] = str_replace(name, "^EN_([0-9_]+_|)", "") # EN_ eraf
    }
  } else if (rename_vars == "vraag") {
    colnames(data.combined)[i] = str_replace(name, "^EN_([0-9_]+_|)", "") # EN_ eraf
  } else {
    if (str_detect(name, "^EN_[0-9]+_")) {
      # subvraag, hernoem hiernaar
      colnames(data.combined)[i] = paste0("Vr", j, "_", str_match(name, "^EN_([0-9]+)")[2])
    } else {
      colnames(data.combined)[i] = paste0("Vr", j)
      j = j + 1
    }
  }
}

write_sav(data.combined, paste0(dirname(data.files[1]), "/resultaten.sav"))

printf("Dataset opgeslagen als %s", paste0(dirname(data.files[1]), "/resultaten.sav"))