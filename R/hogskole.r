#' Analyser regnskapstall fra høgskole
#'
#' @param kilde_fil Filnavn for kildefilen
#' @param output_fil Filnavn for målfilen
#' @param søkenavn Navn på høgskolen som skal søkes etter i målarket
#' @param søke_ark Navn på arket i målfilen der verdiene skal plasseres
#' @return Usynlig retur, skriver data til målfil
#' @export
hogskole_analyse <- function(kilde_fil, output_fil, søkenavn, søke_ark = "Nøkkeltall") {
  
  # Spesifiser kolonner for output
  verdi_kolonner <- c(4, 6, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29)
  
  # Definer nøkkelord for høgskole
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
  
  # Gruppere nøkkelord etter hvilke ark de finnes i
  resultat_nøkkelord <- c("Sum driftsinntekter", "Sum driftskostnader", "Driftsresultat", "Årsresultat")
  eiendeler_nøkkelord <- c("Sum finansielle anleggsmidler", "SUM EIENDELER", "Omløpsmidler")
  gjeld_ek_nøkkelord <- c("Sum opptjent egenkapital", "Sum egenkapital", "Sum avsetning for forpliktelser", 
                        "Sum annen langsiktig gjeld", "Sum kortsiktig gjeld", "SUM EGENKAPITAL OG GJELD")
  
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
    omlopsmidler_komponenter = omlopsmidler_komponenter
  )
}