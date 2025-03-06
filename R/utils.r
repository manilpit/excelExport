#' Konverter tekstverdi til numerisk
#'
#' @param verdi Verdien som skal konverteres
#' @return Numerisk verdi
#' @export
konverter_til_numerisk <- function(verdi) {
  if (is.numeric(verdi)) return(verdi)
  if (is.character(verdi)) {
    # Fjern alle ikke-numeriske tegn unntatt minus og desimaltegn
    tall_tekst <- gsub("[^0-9\\.-]", "", verdi)
    # Bytt ut komma med punktum for desimaltall
    tall_tekst <- gsub(",", ".", tall_tekst)
    # Konverter til tall
    num_val <- suppressWarnings(as.numeric(tall_tekst))
    return(num_val)
  }
  return(NA)
}

#' Søk etter verdi for nøkkelord i dataramme
#'
#' @param data Datarammen som skal søkes i
#' @param nøkkelord Nøkkelordet som skal finnes
#' @return Liste med funnet-flagg og eventuell verdi
#' @export
finn_verdi <- function(data, nøkkelord) {
  if (is.null(data)) return(list(funnet = FALSE))
  
  # Konverter til matrise for enklere søk
  matrise <- as.matrix(data)
  
  # Søk etter nøkkelordet i alle celler
  for (rad in 1:nrow(matrise)) {
    for (kol in 1:ncol(matrise)) {
      cell_value <- matrise[rad, kol]
      
      # Sjekk om cellen matcher nøkkelordet
      if (!is.na(cell_value) && is.character(cell_value)) {
        # Normaliser strenger
        cell_normalized <- tolower(trimws(cell_value))
        keyword_normalized <- tolower(trimws(nøkkelord))
        
        if (cell_normalized == keyword_normalized || 
            grepl(keyword_normalized, cell_normalized, fixed = TRUE)) {
          
          # Nøkkelord funnet - se etter verdi i nærliggende kolonner
          for (offset in 1:5) {
            if (kol + offset <= ncol(matrise)) {
              verdi <- matrise[rad, kol + offset]
              
              # Sjekk om verdien ser ut som et tall (nå som alt er tekst)
              if (!is.na(verdi) && grepl("^\\s*-?[0-9.,]+\\s*$", verdi)) {
                message(paste("Fant", nøkkelord, "med verdi", verdi))
                return(list(
                  funnet = TRUE,
                  verdi = verdi
                ))
              }
            }
          }
        }
      }
    }
  }
  
  # Ikke funnet
  message(paste("Fant ikke", nøkkelord))
  return(list(funnet = FALSE))
}

#' Finn omløpsmidler med alternative metoder
#'
#' @param eiendeler_data Dataramme med eiendelsdata
#' @param omlopsmidler_komponenter Vektor med komponenter for omløpsmidler
#' @return Liste med funnet-flagg og eventuell verdi
#' @export
finn_omlopsmidler <- function(eiendeler_data, omlopsmidler_komponenter) {
  # Først prøv direkte søk
  message("Søker etter 'Omløpsmidler' direkte...")
  direkte <- finn_verdi(eiendeler_data, "Omløpsmidler")
  if (direkte$funnet) return(direkte)
  
  # Prøv også søk etter "Sum omløpsmidler"
  direkte <- finn_verdi(eiendeler_data, "Sum omløpsmidler")
  if (direkte$funnet) return(direkte)
  
  # Prøv å beregne som sum av komponenter
  message("Prøver å beregne omløpsmidler som sum av komponenter...")
  total_sum <- 0
  komponenter_funnet <- 0
  
  for (komponent in omlopsmidler_komponenter) {
    resultat <- finn_verdi(eiendeler_data, komponent)
    if (resultat$funnet) {
      verdi <- konverter_til_numerisk(resultat$verdi)
      if (!is.na(verdi)) {
        total_sum <- total_sum + verdi
        komponenter_funnet <- komponenter_funnet + 1
        message(paste("  Fant komponent:", komponent, "=", verdi))
      }
    }
  }
  
  if (komponenter_funnet > 0) {
    message(paste("Beregnet omløpsmidler:", total_sum, "basert på", komponenter_funnet, "komponenter"))
    return(list(funnet = TRUE, verdi = total_sum))
  }
  
  # Prøv å beregne som SUM EIENDELER - Sum anleggsmidler
  message("Prøver å beregne omløpsmidler som differanse...")
  
  # Søk etter SUM EIENDELER
  sum_eiendeler <- finn_verdi(eiendeler_data, "SUM EIENDELER")
  if (!sum_eiendeler$funnet) {
    sum_eiendeler <- finn_verdi(eiendeler_data, "Sum eiendeler")
  }
  
  # Søk etter Sum anleggsmidler
  sum_anleggsmidler <- finn_verdi(eiendeler_data, "Sum anleggsmidler")
  if (!sum_anleggsmidler$funnet) {
    sum_anleggsmidler <- finn_verdi(eiendeler_data, "Anleggsmidler")
  }
  
  if (sum_eiendeler$funnet && sum_anleggsmidler$funnet) {
    se_verdi <- konverter_til_numerisk(sum_eiendeler$verdi)
    sa_verdi <- konverter_til_numerisk(sum_anleggsmidler$verdi)
    
    if (!is.na(se_verdi) && !is.na(sa_verdi)) {
      omlopsmidler <- se_verdi - sa_verdi
      message(paste("Beregnet omløpsmidler:", omlopsmidler, "som SUM EIENDELER -", sa_verdi))
      return(list(funnet = TRUE, verdi = omlopsmidler))
    }
  }
  
  return(list(funnet = FALSE))
}

#' Finn eller opprett rad for institusjon i målarket
#'
#' @param mål_data Datarammen for målarket
#' @param søkenavn Navnet på institusjonen
#' @return Radnummer
#' @export
finn_institusjon_rad <- function(mål_data, søkenavn) {
  funnet_rad <- NULL
  
  # Sjekk om mål_data er NULL eller har 0 rader
  if (is.null(mål_data) || nrow(mål_data) == 0) {
    # Opprett en tom dataramme med en rad
    return(1)
  }
  
  # Sjekk om datarammen har minst en kolonne
  if (ncol(mål_data) >= 1) {
    # Konverter første kolonne til tekst for sammenligning
    første_kolonne <- as.character(mål_data[[1]])
    
    # Søk etter søkenavn
    for (i in 1:length(første_kolonne)) {
      if (!is.na(første_kolonne[i]) && 
          tolower(trimws(første_kolonne[i])) == tolower(trimws(søkenavn))) {
        funnet_rad <- i
        message(paste("Fant", søkenavn, "på rad", funnet_rad))
        break
      }
    }
    
    # Hvis ikke funnet, legg til rad (enten som ny eller på første tomme plass)
    if (is.null(funnet_rad)) {
      # Søk etter tom rad
      for (i in 1:length(første_kolonne)) {
        if (is.na(første_kolonne[i]) || første_kolonne[i] == "") {
          funnet_rad <- i
          message(paste("Bruker tom rad", funnet_rad, "for", søkenavn))
          break
        }
      }
      
      # Hvis ingen tom rad, legg til ny rad
      if (is.null(funnet_rad)) {
        funnet_rad <- nrow(mål_data) + 1
        message(paste("Legger til ny rad", funnet_rad, "for", søkenavn))
      }
    }
  } else {
    # Hvis datarammen ikke har kolonner, bruk rad 1
    funnet_rad <- 1
    message(paste("Bruker rad", funnet_rad, "i dataramme uten kolonner for", søkenavn))
  }
  
  # Sikre at vi alltid returnerer et gyldig radnummer
  if (is.null(funnet_rad) || is.na(funnet_rad)) {
    funnet_rad <- 1
    message(paste("Fallback: Bruker rad 1 for", søkenavn))
  }
  
  return(funnet_rad)
}