#' Analyser regnskapstall fra fagskole
#'
#' @param kilde_fil Filnavn for kildefilen
#' @param output_fil Filnavn for målfilen
#' @param søkenavn Navn på fagskolen som skal søkes etter i målarket
#' @param søke_ark Navn på arket i målfilen der verdiene skal plasseres
#' @return Usynlig retur, skriver data til målfil
#' @export
fagskole_analyse <- function(kilde_fil, output_fil, søkenavn, søke_ark = "Nøkkeltall") {
  
  # Spesifiser kolonner for output
  verdi_kolonner <- c(4, 6, 8, 10, 12, 14, 16)
  
  # Definer nøkkelord for fagskole
  alle_nøkkelord <- c(
    "Driftsinntekter",
    "Driftskostnader",
    "Driftsresultat",          
    "Årsresultat",
    "Omløpsmidler",
    "Egenkapital",
    "Kortsiktig gjeld"
  )
  
  # Gruppere nøkkelord etter hvilke ark de finnes i
  resultat_nøkkelord <- c("Driftsinntekter", "Driftskostnader", "Driftsresultat", "Årsresultat")
  eiendeler_nøkkelord <- c("Omløpsmidler")
  gjeld_ek_nøkkelord <- c("Egenkapital", "Kortsiktig gjeld", "Totalkapital")
  
  # Komponenter for beregning av omløpsmidler om nødvendig
  omlopsmidler_komponenter <- c(
    "Sum varer", 
    "Sum fordringer", 
    "Sum investeringer", 
    "Sum bankinnskudd, kontanter og lignende"
  )
  
  # Utfør regnskapsanalysen
  utfør_regnskapsanalyse(
    kilde_fil = kilde_fil,
    output_fil = output_fil,
    søkenavn = søkenavn,
    søke_ark = søke_ark,
    verdi_kolonner = verdi_kolonner,
    alle_nøkkelord = alle_nøkkelord,
    resultat_nøkkelord = resultat_nøkkelord,
    eiendeler_nøkkelord = eiendeler_nøkkelord,
    gjeld_ek_nøkkelord = gjeld_ek_nøkkelord,
    omlopsmidler_komponenter = omlopsmidler_komponenter,
    søk_totalkapital = TRUE
  )
}