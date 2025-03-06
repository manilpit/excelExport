# excelExport
excelExport er en R-pakke som forenkler prosessen med å trekke ut nøkkeltall fra regnskaper for utdanningsinstitusjoner og lagre disse i oversiktlige Excel-rapporter.

## Installasjon
Du kan installere pakken direkte fra GitHub med devtools:
```r
# Installer devtools hvis du ikke har det
# install.packages("devtools")

# Installer pakken fra GitHub
devtools::install_github("manilpit/excelExport")
```

Eller ved å klone repositoret og installere lokalt:
```r
devtools::install_local("path/to/excelExport")
```

## Funksjonalitet
Pakken tilbyr funksjonalitet for å:

Trekke ut spesifikke nøkkeltall fra Excel-baserte regnskaper
1. Lagre disse verdiene strukturert i målfiler
2. Støtte tre typer institusjoner: høgskoler, fagskoler og studentsamskipnader
3. Automatisk gjenkjenne relevant informasjon i forskjellige regnskapsoppsett
4. Bevare eksisterende data i målfiler
## Brukseksempler
### Grunnleggende bruk
```r
library(excelExport)

# Analyser regnskapsdata for en høgskole
analyser_regnskap(
  kilde_fil = "8208_2024_VID.xlsx",        # Kildefil med regnskapsdata
  output_fil = "VID_Resultatark.xlsx",     # Målfil for nøkkeltall
  søkenavn = "VID_SKOLE",                 # Navn på institusjonen i målarket
  type = "hogskole",                      # Type institusjon
  søke_ark = "Nøkkeltall"                 # Navnet på arket der data skal plasseres
)
```

Analysere regnskaper for ulike institusjonstyper
For høgskole:

```r
analyser_regnskap(
  kilde_fil = "8208_2024_VID.xlsx",
  output_fil = "Høgskoler_Oversikt.xlsx",
  søkenavn = "VID Vitenskapelige Høgskole",
  type = "hogskole" 
)
```
For fagskole:

```r
analyser_regnskap(
  kilde_fil = "Fagskole_Regnskap_2024.xlsx",
  output_fil = "Fagskoler_Oversikt.xlsx",
  søkenavn = "Noroff Fagskole",
  type = "fagskole"
)
```
For studentsamskipnad:

```r
analyser_regnskap(
  kilde_fil = "SiO_Regnskap_2024.xlsx",
  output_fil = "Studentsamskipnader_Oversikt.xlsx",
  søkenavn = "Studentsamskipnaden i Oslo",
  type = "studentsamskipnad"
)
```
## Nøkkeltall som trekkes ut
### Avhengig av institusjonstype hentes følgende nøkkeltall:

Høgskoler:
1. Sum driftsinntekter
2. Sum driftskostnader
3. Driftsresultat
4. Årsresultat
5. Sum finansielle anleggsmidler
6. SUM EIENDELER
7. Omløpsmidler
8. Sum opptjent egenkapital
9. Sum egenkapital
10. Sum avsetning for forpliktelser
11. Sum annen langsiktig gjeld
12. Sum kortsiktig gjeld
13. SUM EGENKAPITAL OG GJELD

Fagskoler
1. Driftsinntekter
2. Driftskostnader
3. Driftsresultat
4. Årsresultat
5. Omløpsmidler
6. Egenkapital
7. Kortsiktig gjeld
8. Totalkapital (beregnes om ikke funnet)

Studentsamskipnader
1. Driftsinntekter
2. Driftskostnader
3. Driftsresultat
4. Årsresultat
5. Omløpsmidler
6. Egenkapital
7. Avsetning for forpliktelser
8. Annen langsiktig gjeld
9. Kortsiktig gjeld
10. Sum gjeld
11. Totalkapital (beregnes om ikke funnet)

# Lisens
Denne pakken er lisensiert under MIT-lisensen.

# Bidrag
Bidrag til pakken er velkomne! Send inn issues eller pull requests på GitHub.