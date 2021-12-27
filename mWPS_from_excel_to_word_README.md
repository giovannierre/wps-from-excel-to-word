# *mWPS_from_excel_to_word*

Un sistema per automatizzare la scrittura di WPS partendo da un file di Excel.

###### Sommario

[TOC]

# Introduzione

Il sistema legge i dati da Excel e li memorizza nelle *CustomProperties* di un documento di Word. In Word, tramite codici di campo, è possibile visualizzare i rispettivi valori. Il sistema è "autosettante" nel senso che si basa sulla coincidenza tra i nomi delle colonne di Excel e i nomi dei campi di Word. Per ottimizzare parzialmente le risorse, si è stabilito di escludere dalla lettura i campi che iniziano con "_" (underscore).

Inoltre è previsto anche l'inserimento di una immagine con dimensioni limite specificate. L'immagine è individuata dal nome del file e percorso.

# Riferimenti necessari

Il modulo utilizza altri moduli esterni che devono essere resi disponibili:

* **mCreateCustomProperties**:  mette a disposizione le funzioni che servono per verificare se una *CustomProperty* esiste e per crearla nel caso non esista.

* **mAuxiliaryFunctions**: mette a disposizione almeno la funzione *IsFileOpen* che serve per verificare se il template di Word è già aperto evitando di aprirne istanze non necessarie.  

  

# Elenco delle routine

* [**process_wps**](#process_wps): questa routine riconosce la selezione corrente e lancia la routine di interpretazione dati e scrittura documento (*read_wps_data*) basata sulla selezione
* [**read_wps_data**](#read_wps_data): è la routine principale, legge i dati dal foglio di Excel e li scrive nel documento di Word 
* **ShowUserForm**:
* **UpdateUserForm1**:

*[[index](#sommario)]*

## process_wps

```vb
Sub process_wps()
```

Scorre tutte le righe dell'intervallo selezionato e per ognuna di esse lancia la procedura di scrittura della WPS (*read_wps_data*)

*[[index](#sommario)]*

## read_wps_data

```vb
Sub read_wps_data(MySheet As Excel.Worksheet, CurrentCell As Excel.Range, AutomaticSave As Boolean)
```

La routine è strutturata in questo modo:

1) Una sezione SETTINGS dove sono definiti i nomi delle celle che contengono i percorsi per recuperare modelli e scrivere file. Questa è l'unica sezione che potrebbe essere necessario adattare all'applicazione specifica.

   ```vb
   '****************
   '****SETTINGS****
   '****************
       'Dimensioni massime immagine giunto, in cm
       IMAGE_HEIGHT = 3.5
       IMAGE_WIDTH = 8.3
       'Nomi delle cella che contengono i percorsi dei vari file:
       NAMED_CELL_FOR_IMAGE_PATH = "ImagePath"
       NAMED_CELL_FOR_TEMPLATE_PATH = "TemplateFullPath"
       NAMED_CELL_FOR_PDF_EXPORT_PATH = "SavePdfPath"
       NAMED_CELL_FOR_WORD_EXPORT_PATH = "SaveWordPath"
       'Nomi di alcuni campi "speciali":
       TABLE_FIELD_FOR_IMAGE_FILENAME = "joint_sketch_file"
   '***FINE SETTINGS***
   ```

   

2) **Individuazione della tabella dei dati**: viene individuato il primo oggetto *ListObject* nel foglio passato come argomento; si suppone che sia l'unica tabella nel foglio.

3) **Lettura delle intestazioni della tabella e memorizzazione in una collection** (*PropertyName*). Per evitare di leggere intestazioni non necessarie, vengono escluse quelle che iniziano con "_" (underscore).

4) **Lettura dei valori della riga della cella passata come argomento** (iterazione sulle celle della riga) e memorizzazione in una collection (*PropertyValue*). Il *key* della collection di ogni elemento memorizzato coincide con la rispettiva intestazione di colonna (nome del campo), in questo modo nella fase di scrittura iterando sui *PropertyName* e passandoli come *key* si potranno ricavare i valori dei rispettivi *PropertyValue*.

   > Nota: attualmente il sistema limita i *PropertyName* escludendo i campi con underscore, ma legge e memorizza comunque tutti i campi; di questi saranno poi utilizzati solo quelli il cui *key* coincide con i valori memorizzati in *PropertyName*. Su questo punto c'è spazio per un'ottimizzazione.

5) **Apertura del modello di Word**

6) **Creazione e scrittura delle *CustomProperties*** di Word (NB: tramite funzione in modulo esterno) nelle quali vengono memorizzati i valori letti da Excel. Se una *CustomProperty* non esiste, viene creata automaticamente.

7) **Inserimento dell'immagine**. L'immagine viene specifica come file e viene ridimensionata per rispettare larghezza e altezza massima come specificato nei settings.

8) **Salvataggio del file generato**. Il file può essere salvato sia in pdf ce in word.



