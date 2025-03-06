#' Analyser regnskapstall fra studentsamskipnad
#'
#' @param kilde_fil Filnavn for kildefilen
#' @param output_fil Filnavn for målfilen
#' @param søkenavn Navn på studentsamskipnaden som skal søkes etter i målarket
#' @param søke_ark Navn på arket i målfilen der verdiene skal plasseres
#' @return Usynlig retur, skriver data til målfil
#' @export
studentsamskipnad_analyse <- function(kilde_fil, output_fil, søkenavn, søke_ark = "Nøkkeltall") {
  
  # Spesifiser kolonner for output
  verdi_kolonner <- c(4, 6, 8, 10, 12, 14, 16, 18, 20, 22)
  
  # Definer nøkkelord for studentsamskipnad
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
  
  # Gruppere nøkkelord etter hvilke ark de finnes i
  resultat_nøkkelord <- c("Driftsinntekter", "Driftskostnader", "Driftsresultat", "Årsresultat")
  eiendeler_nøkkelord <- c("Omløpsmidler")
  gjeld_ek_nøkkelord <- c("Egenkapital", "Avsetning for forpliktelser", "Annen langsiktig gjeld", 
                          "Kortsiktig gjeld", "Sum gjeld", "Totalkapital")
  
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
    søk_totalkapital = TRUE,
    beregn_totalkapital_fra_gjeld = TRUE
  )
}