Vi ho caricato i nuovi codici e i nuovi input file per Metamap

Nella cartella "Codici VBA" trovate i tre script VBA usati per generare i file di input per Metamap.
Questi tre script vanno eseguiti sequenzialmente per ogni database in quest'ordine:
- Extraction&cleaning (in Excel)
- DetectLanguageScript (in Word)
- FilteringSpuriousCharacters (in Excel)

--> Vi ho riportato i codici che ho usato per il database "Medical".
--> Per usarli con "Health&Fitness" cambiate i nomi dei file che vengono generati e il numero dell'ultima riga da considerare nel ciclo for di "Extraction&cleaning"


All'interno di "Extraction&cleaning" e di "DetectLanguageScript" ho generato anche due file in cui vengono riportati gli ID delle app che quei due script stampano
e il numero total di questi ID, così sappiamo quante app abbiamo passo passo.

A questo proposito, vi riporto qui i dettagli che ho ottenuto eseguendo:

MEDICAL database
 - 33858 app totali
 - 26738 app dopo aver applicato "Extraction&cleaning"
 - 21265 app dopo aver applicato "DetectLanguageScript" (questo è il numero di app finali di "Medical" che Metamap dovrà processare)

HEALTH & FITNESS database
 - 51994 app totali
 - 42749 app dopo aver applicato "Extraction&cleaning"
 - 34910 app dopo aver applicato "DetectLanguageScript" (questo è il numero di app finali di "Health & Fitness" che Metamap dovrà processare)


Infine, nella cartella "File di Input per Metamap" vi ho riportato i due file di input generati, pronti da dare a Metamap.