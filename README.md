# Valutakursberegning av Erasmus+-stipender
Valutakursberegning for utbetaling av Erasmus+-stipend skoleåret 2024/2025 ved Enhet for internasjonale relasjoner på NTNU.

Programmet kjøres på lokal datamaskin enten som en .exe fil (nedlasting finner du under "Releases"), eller ved å kjøre .py-filen du kan laste ned ovenfor. Sistnevnte krever program til å kunne kjøre Python-filer, samt installere avhengigheter (dependencies) programmet bruker som pandas, tkinter, requests, io, PyInstaller om dette ikke allerede er installert. Ved "kjøring" av program, uavhengig om det er .exe eller .py, se stegene nedenfor.

Det ligger også en testfil: ``test.xlsx``, du kan selv teste programmet på.


## Hvordan funker stipendutbetaling ved NTNU?
For skoleåret 2024/2025 betaler NTNU ut Erasmus+-stipend i to runder, i henholdsvis en første og en siste utbetaling. Tildeling (totalt, første, siste) skjer i euro i Mobility Online og stipendene hentes ut i Excel-lister med studentenes informasjon. En fast eurokurs på 11,7 brukes for første utbetaling og som "manuelt" regnes om i Excel. Programmet kjøres ved andre/siste utbetaling ved å hente en liste over alle dagskurser fra den Europeiske sentralbanken (ECB) fra og med 1. januar 1999 og til og med dagens dato, deretter finner den gjennomsnittet mellom dagene studenten var ute og gir en snittkurs per opphold/student. Snittkursen ganges så med total tildeling og beløpet studenten fikk i første utbetaling trekkes fra den totale tildelingen i kroner for å få siste utbetaling. Listen(e) sendes så til utbetaling.

Formel for andre utbetaling blir derfor:
Total tildeling * Snittkurs for opphold - første utbetaling * 11,7

### Før du starter
Før du kjører programmet må du ha en utbetalingsliste (Excel-liste, lagret som .xlsx). I tillegg til vanlig utbetalingsinformasjon må du ha med fire kolonner:
* ``First payment``
* ``Total grant``
* ``Stay from``
* ``Stay to``

``First payment`` og ``Total grant`` er henholdsvis beløpene studenten har fått i første utbetaling, og det studenten totalt har fått tildelt i stipend. Begge beløpene er i euro. ``Stay from`` og ``Stay to`` er henholdsvis start- og sluttdato på oppholdet.

Se ``test.xlsx`` for et eksempel for hvordan en utbetalingsliste kan se ut før programmet er kjørt. Det er viktig at navnene på de fire kolonnene over er nøyaktig det som står over, hvis ikke vil programmet feile. Rekkefølgen på kolonnene har ikke noe å si. Navnene på de andre kolonnene kan være hva som helst annet enn det samme som de fire kolonnene over, og heller ikke ``Snittkurs`` eller ``Siste utbetaling``. I tillegg er det anbefalt at du kjører programmet på en kopi av utbetalingslisten i tilfelle noe skjer under kjøringen.

## Kjøring av program
Du har nå gjort klar en Excel-liste lagret et sted du finner den, og du har en måte å kjøre programmet på din egen maskin. 

* Lukk Excel-filen du skal beregne stipend på. Programmet vil ikke kunne lagre dokumentet dersom det er åpent. 
* Kjør programmet
* Et svart konsollvindu åpnes. Her vil eventuelle feilmeldinger vises, samt om det fullførte uten problemer.
* Programmet henter så dagskursen for hver dag fra og med 1999 til og med dagens dato. Dette kan ta litt tid og gjøres hver gang programmet kjøres.
* Du vil deretter bli bedt om å velge en fil. Her velger du Excel-listen for andre utbetaling.
* Om alt går etter planen vil du få en beskjed i konsollvinduet: "Success!"

Programmet har nå lagt inn en ny kolonne for eurokursen for hvert av oppholdene, og en kolonne for siste utbetaling regnet ut i kroner.

## Typiske feil
Sjekk at alle punktene under er gjort dersom ting ikke funker

* Du har ikke lukket Excel-filen programmet brukes på. Det er derfor ikke mulig for programmet å lagre filen.
* Excel-filen er lagret som "read only". Lagre filen på nytt og sørg for at du også har mulighet til å redigere.
* Programmet klarer ikke hente .csv filen fra ECB. Du må være tilkoblet internett. Hvis ikke kan det hende du må bruke en proxy. Epost med første utkast fra ECB inneholder dette.
* Dersom startdato er før 1999 (bør ikke skje) vil første dag som brukes i snittkursberegningen være 1. januar 1999.
* Dersom sluttdato er senere enn dagens dato vil siste dag brukt i beregning være den siste datoen det er hentet snittkurs på (dagens dato eller et par dager tidligere)
* Dersom hele perioden er frem i tid vil det ikke genereres noen snittkurs.

## Endring av kode
For å gjøre endringer i programmet må du kunne redigere og kjøre Python-filer. Du laster i så fall ned .py-filen og gjør nødvendige endringer.

For å lage en .exe fil med oppdaterte endringer kjøres følgende kode i konsollen:

``python -m PyInstaller --onefile valutakursberegning.py``
