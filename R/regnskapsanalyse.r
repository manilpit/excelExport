#' Analyser regnskapstall fra institusjon
#'
#' @param kilde_fil Filnavn for kildefilen
#' @param output_fil Filnavn for målfilen
#' @param søkenavn Navn på institusjonen som skal søkes etter i målarket
#' @param type Type institusjon: "hogskole", "fagskole" eller "studentsamskipnad"
#' @param søke_ark Navn på arket i målfilen der verdiene skal plasseres
#' @return Usynlig retur, skriver data til målfil
#' @export
analyser_regnskap <- function(kilde_fil, output_fil, søkenavn, 
  type = c("hogskole", "fagskole", "studentsamskipnad"),
  søke_ark = "Nøkkeltall") {

type <- match.arg(type)

# Last inn nødvendige pakker
requireNamespace("readxl", quietly = TRUE)
requireNamespace("writexl", quietly = TRUE)
requireNamespace("dplyr", quietly = TRUE)

# Velg riktige parametere basert på institusjonstype
if (type == "hogskole") {
verdi_kolonner <- c(4, 6, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29)

alle_nøkkelord <- c(
"Sum driftsinntekter", 
"Sum driftskostnader", 
"Driftsresultat", 
"Årsresultat",
"Sum finansielle anleggsmidler", 
"SUM EIENDELER", 
"Omløpsmidler",
"Sum opptjent egenkapital", 
"Sum egenkapital",
"Sum avsetning for forpliktelser", 
"Sum annen langsiktig gjeld", 
"Sum kortsiktig gjeld", 
"SUM EGENKAPITAL OG GJELD"
)
} else if (type == "fagskole") {
verdi_kolonner <- c(4, 6, 8, 10, 12, 14, 16)

alle_nøkkelord <- c(
"Driftsinntekter",
"Driftskostnader",
"Driftsresultat",          
"Årsresultat",
"Omløpsmidler",
"Egenkapital",
"Kortsiktig gjeld"
)
} else if (type == "studentsamskipnad") {
verdi_kolonner <- c(4, 6, 8, 10, 12, 14, 16, 18, 20, 22)

alle_nøkkelord <- c(
"Driftsinntekter",
"Driftskostnader",
"Driftsresultat",          
"Årsresultat",
"Omløpsmidler",
"Egenkapital",
"Avsetning for forpliktelser",
"Annen langsiktig gjeld",
"Kortsiktig gjeld",
"Sum gjeld"
)
}

# Gruppere nøkkelord etter hvilke ark de finnes i
resultat_nøkkelord <- alle_nøkkelord[grep("driftsinntekter|driftskostnader|driftsresultat|årsresultat", 
               tolower(alle_nøkkelord))]

eiendeler_nøkkelord <- alle_nøkkelord[grep("anleggsmidler|eiendeler|omløpsmidler", 
                tolower(alle_nøkkelord))]

gjeld_ek_nøkkelord <- alle_nøkkelord[grep("egenkapital|forpliktelser|gjeld", 
              tolower(alle_nøkkelord))]

# Komponenter for beregning av omløpsmidler om nødvendig
omlopsmidler_komponenter <- c(
"Sum varer", 
"Sum fordringer", 
"Sum investeringer", 
"Sum bankinnskudd, kontanter og lignende"
)

# 3. Les inn alle ark fra kildefilen
message(paste("Leser inn kildefil:", kilde_fil))
tilgjengelige_ark <- readxl::excel_sheets(kilde_fil)
message("Tilgjengelige ark:")
message(paste(tilgjengelige_ark, collapse = ", "))

# Definer arknavnene basert på tilgjengelige ark
resultat_ark <- "Resultatregnskap"
if (!(resultat_ark %in% tilgjengelige_ark)) {
# Prøv å finne et alternativt ark for resultatregnskap
resultat_ark_kandidater <- tilgjengelige_ark[grep("resultat", tilgjengelige_ark, ignore.case = TRUE)]
if (length(resultat_ark_kandidater) > 0) {
resultat_ark <- resultat_ark_kandidater[1]
} else {
message("ADVARSEL: Fant ikke resultatregnskap-ark!")
}
}

eiendeler_ark <- NULL
if ("Balanse - eiendeler" %in% tilgjengelige_ark) {
eiendeler_ark <- "Balanse - eiendeler"
} else {
# Prøv å finne et alternativt ark for eiendeler
eiendeler_ark_kandidater <- tilgjengelige_ark[grep("balanse.*eiendeler|eiendeler.*balanse", 
                        tilgjengelige_ark, ignore.case = TRUE)]
if (length(eiendeler_ark_kandidater) > 0) {
eiendeler_ark <- eiendeler_ark_kandidater[1]
}
}

gjeld_ek_ark_resultat <- grep("balanse.*gjeld|gjeld.*balanse|egenkapital", 
  tilgjengelige_ark, ignore.case = TRUE)

gjeld_ek_ark <- NULL
if (length(gjeld_ek_ark_resultat) > 0) {
gjeld_ek_ark <- tilgjengelige_ark[gjeld_ek_ark_resultat[1]]
}

message(paste("Bruker følgende ark:"))
message(paste("Resultatregnskap:", resultat_ark))
message(paste("Eiendeler:", ifelse(is.null(eiendeler_ark), "Ikke funnet", eiendeler_ark)))
message(paste("Gjeld og EK:", ifelse(is.null(gjeld_ek_ark), "Ikke funnet", gjeld_ek_ark)))

# 4. Les inn data fra ark i kildefilen - ALLE kolonner som tekst for å unngå dato-problemer
resultat_data <- tryCatch({
# Les alle kolonner som tekst
readxl::read_excel(kilde_fil, sheet = resultat_ark, col_types = "text")
}, error = function(e) {
message(paste("Feil ved lesing av", resultat_ark, ":", e$message))
NULL
})

eiendeler_data <- if (!is.null(eiendeler_ark)) {
tryCatch({
readxl::read_excel(kilde_fil, sheet = eiendeler_ark, col_types = "text")
}, error = function(e) {
message(paste("Feil ved lesing av", eiendeler_ark, ":", e$message))
NULL
})
} else NULL

gjeld_ek_data <- if (!is.null(gjeld_ek_ark)) {
tryCatch({
readxl::read_excel(kilde_fil, sheet = gjeld_ek_ark, col_types = "text")
}, error = function(e) {
message(paste("Feil ved lesing av", gjeld_ek_ark, ":", e$message))
NULL
})
} else NULL

# 5. Les inn målfilen hvis den eksisterer
mål_data <- NULL
  
if (file.exists(output_fil)) {
  message(paste("Leser inn målfil:", output_fil))
  
  # Liste over alle ark i målfilen
  tryCatch({
    mål_ark <- readxl::excel_sheets(output_fil)
    message("Tilgjengelige ark i målfilen:")
    message(paste(mål_ark, collapse = ", "))
    
    # Sjekk om søke_ark finnes
    if (søke_ark %in% mål_ark) {
      # Les alle kolonner som tekst for å unngå konverteringsproblemer
      mål_data <- readxl::read_excel(output_fil, sheet = søke_ark, col_types = "text")
      message(paste("Leste inn arket", søke_ark, "fra målfilen"))
    } else {
      message(paste("Advarsel: Arket", søke_ark, "finnes ikke i målfilen. Oppretter ny."))
      # Opprett en tom dataramme
      mål_data <- data.frame(matrix(ncol = max(verdi_kolonner) + 1, nrow = 1))
      colnames(mål_data) <- c("X1", paste0("X", 2:ncol(mål_data)))
    }
  }, error = function(e) {
    message(paste("Feil ved lesing av målfil:", e$message))
    # Opprett en tom dataramme
    mål_data <<- data.frame(matrix(ncol = max(verdi_kolonner) + 1, nrow = 1))
    colnames(mål_data) <<- c("X1", paste0("X", 2:ncol(mål_data)))
  })
} else {
  message("Målfil eksisterer ikke. Vil opprette ny fil.")
  # Opprett en tom dataramme
  mål_data <- data.frame(matrix(ncol = max(verdi_kolonner) + 1, nrow = 1))
  colnames(mål_data) <- c("X1", paste0("X", 2:ncol(mål_data)))
}

# Sikre at mål_data har riktig antall kolonner
if (ncol(mål_data) < max(verdi_kolonner) + 1) {
  message("Utvider mål_data med flere kolonner")
  # Legg til flere kolonner om nødvendig
  for (i in (ncol(mål_data) + 1):(max(verdi_kolonner) + 1)) {
    mål_data[[i]] <- NA
    if (!is.null(colnames(mål_data))) {
      colnames(mål_data)[i] <- paste0("X", i)
    }
  }
}

# 9. Samle alle nøkkelverdier
message("Henter verdier fra kildefil...")
verdier <- list()

for (nøkkelord in alle_nøkkelord) {
# Velg riktig dataramme basert på nøkkelord
if (nøkkelord %in% resultat_nøkkelord) {
resultat <- finn_verdi(resultat_data, nøkkelord)
} else if (nøkkelord == "Omløpsmidler") {
resultat <- finn_omlopsmidler(eiendeler_data, omlopsmidler_komponenter)
} else if (nøkkelord %in% eiendeler_nøkkelord) {
resultat <- finn_verdi(eiendeler_data, nøkkelord)
} else if (nøkkelord %in% gjeld_ek_nøkkelord) {
resultat <- finn_verdi(gjeld_ek_data, nøkkelord)
} else {
resultat <- list(funnet = FALSE)
}

if (resultat$funnet) {
verdier[[nøkkelord]] <- konverter_til_numerisk(resultat$verdi)
message(paste("Lagret", nøkkelord, "=", verdier[[nøkkelord]]))
} else {
message(paste("ADVARSEL: Kunne ikke finne verdi for", nøkkelord))
}
}

# Legg til søk etter Totalkapital for fagskole og studentsamskipnad
if (type %in% c("fagskole", "studentsamskipnad")) {
message("Søker etter 'Totalkapital'...")
totalkapital_resultat <- finn_verdi(gjeld_ek_data, "Totalkapital")

if (!is.null(eiendeler_data) && !totalkapital_resultat$funnet) {
totalkapital_resultat <- finn_verdi(eiendeler_data, "SUM EIENDELER")
}

if (!is.null(gjeld_ek_data) && !totalkapital_resultat$funnet) {
totalkapital_resultat <- finn_verdi(gjeld_ek_data, "SUM EGENKAPITAL OG GJELD")
}

if (totalkapital_resultat$funnet) {
verdier[["Totalkapital"]] <- konverter_til_numerisk(totalkapital_resultat$verdi)
message(paste("Fant Totalkapital =", verdier[["Totalkapital"]]))

# Legg Totalkapital til i alle_nøkkelord hvis den ikke allerede er der
if (!("Totalkapital" %in% alle_nøkkelord)) {
alle_nøkkelord <- c(alle_nøkkelord, "Totalkapital")
verdi_kolonner <- c(verdi_kolonner, max(verdi_kolonner) + 2)  # Legg til kolonne for totalkapital
message(paste("La til Totalkapital i kolonne", max(verdi_kolonner)))
}
} else if (type == "studentsamskipnad") {
# Prøv å beregne Totalkapital fra gjeldsposter for studentsamskipnad
if ("Egenkapital" %in% names(verdier) && "Sum gjeld" %in% names(verdier)) {
totalkapital <- verdier[["Egenkapital"]] + verdier[["Sum gjeld"]]
verdier[["Totalkapital"]] <- totalkapital
message(paste("Beregnet Totalkapital =", totalkapital, "som Egenkapital + Sum gjeld"))

if (!("Totalkapital" %in% alle_nøkkelord)) {
alle_nøkkelord <- c(alle_nøkkelord, "Totalkapital")
verdi_kolonner <- c(verdi_kolonner, max(verdi_kolonner) + 2)
message(paste("La til Totalkapital i kolonne", max(verdi_kolonner)))
}
} else if (!("Sum gjeld" %in% names(verdier)) && 
all(c("Avsetning for forpliktelser", "Annen langsiktig gjeld", "Kortsiktig gjeld") %in% names(verdier))) {
# Beregn Sum gjeld som sum av alle gjeldsposter
sum_gjeld <- verdier[["Avsetning for forpliktelser"]] + 
verdier[["Annen langsiktig gjeld"]] + 
verdier[["Kortsiktig gjeld"]]

verdier[["Sum gjeld"]] <- sum_gjeld
message(paste("Beregnet Sum gjeld =", sum_gjeld, "som sum av alle gjeldsposter"))

if ("Egenkapital" %in% names(verdier)) {
totalkapital <- verdier[["Egenkapital"]] + sum_gjeld
verdier[["Totalkapital"]] <- totalkapital
message(paste("Beregnet Totalkapital =", totalkapital, "som Egenkapital + Sum gjeld"))

if (!("Totalkapital" %in% alle_nøkkelord)) {
alle_nøkkelord <- c(alle_nøkkelord, "Totalkapital")
verdi_kolonner <- c(verdi_kolonner, max(verdi_kolonner) + 2)
message(paste("La til Totalkapital i kolonne", max(verdi_kolonner)))
}
}
}
}
}

# 10. Finn raden med søkenavn i målarket eller lag en ny rad
funnet_rad <- finn_institusjon_rad(mål_data, søkenavn)
  
# Sikre at mål_data har nok rader
if (nrow(mål_data) < funnet_rad) {
  message(paste("Utvider mål_data til", funnet_rad, "rader"))
  ekstra_rader <- funnet_rad - nrow(mål_data)
  nye_rader <- matrix(NA, nrow = ekstra_rader, ncol = ncol(mål_data))
  mål_data <- rbind(mål_data, nye_rader)
}

# 11. Skriv søkenavn til første kolonne
mål_data[funnet_rad, 1] <- søkenavn

# 12. Skriv verdier til de angitte kolonnene
verdier_skrevet <- 0

for (i in 1:length(alle_nøkkelord)) {
  nøkkelord <- alle_nøkkelord[i]
  
  if (nøkkelord %in% names(verdier) && i <= length(verdi_kolonner)) {
    kolonne <- verdi_kolonner[i]
    
    # Sikre at kolonnen eksisterer
    if (kolonne > ncol(mål_data)) {
      message(paste("Utvider mål_data til", kolonne, "kolonner"))
      for (j in (ncol(mål_data) + 1):kolonne) {
        mål_data[[j]] <- NA
        colnames(mål_data)[j] <- paste0("X", j)
      }
    }
    
    # VIKTIG: Konverter verdien til tekst før skriving til datarammen
    verdi_som_tekst <- as.character(verdier[[nøkkelord]])
    
    # Skriv verdien
    mål_data[funnet_rad, kolonne] <- verdi_som_tekst
    message(paste("Plasserte", nøkkelord, "=", verdi_som_tekst, "i kolonne", kolonne))
    verdier_skrevet <- verdier_skrevet + 1
  }
}

# 13. Les inn alle ark fra målfilen for å bevare dem
alle_ark <- list()
  
if (file.exists(output_fil)) {
  tryCatch({
    mål_ark_liste <- readxl::excel_sheets(output_fil)
    
    for (ark_navn in mål_ark_liste) {
      if (ark_navn != søke_ark) {
        # Les inn andre ark for å bevare dem
        alle_ark[[ark_navn]] <- readxl::read_excel(output_fil, sheet = ark_navn, col_types = "text")
      }
    }
  }, error = function(e) {
    message(paste("Feil ved lesing av ark fra målfil:", e$message))
  })
}

# Legg til vårt oppdaterte ark
alle_ark[[søke_ark]] <- mål_data

# 14. Skriv alle ark tilbake til målfilen
tryCatch({
  writexl::write_xlsx(alle_ark, path = output_fil)
  message(paste("Lagret", length(alle_ark), "ark til filen:", output_fil))
}, error = function(e) {
  message(paste("Feil ved skriving til målfil:", e$message))
  
  # Prøv å skrive bare det nye arket hvis det var feil med hele målfilen
  en_ark <- list()
  en_ark[[søke_ark]] <- mål_data
  writexl::write_xlsx(en_ark, path = output_fil)
  message(paste("Lagret kun", søke_ark, "til filen:", output_fil))
})
  
# 15. Skriv sammendrag
message("==== SAMMENDRAG ====")
message(paste("Totalt", length(alle_nøkkelord), "nøkkelord søkt etter"))
message(paste("Funnet", length(verdier), "verdier"))
message(paste("Skrevet", verdier_skrevet, "verdier til rad", funnet_rad, "i målfilen"))

if (length(verdier) < length(alle_nøkkelord)) {
mangler <- alle_nøkkelord[!alle_nøkkelord %in% names(verdier)]
message("Manglende nøkkelord:")
for (m in mangler) {
message(paste(" -", m))
}
}
message("===================")

# Returner usynlig resultat (ingen verdi)
invisible()
}