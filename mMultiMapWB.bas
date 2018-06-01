Attribute VB_Name = "mMultiMapWB"
Option Explicit

Sub ElaborateWB()
Attribute ElaborateWB.VB_ProcData.VB_Invoke_Func = "e\n14"
'Questa routine serve per elaborare welding book composti da molte Welding Map,
'per i quali nel foglio WPS sono specificati, in un'unica cella, Welding Map e Joint No.,
'esempio: "WM001: W1, W2; WM002: W1, W3"
'Le diverse welding map devono essere separate da ";" (punto e virgola).
'Il nome della welding map deve essere seguito dai ":" (due punti) e poi dai nomi dei giunti.

'La routine legge i dati dalla tabella sorgente e sul campo che viene specificato nei settings come
'quello da splittare, effettua lo split dei valori e con un ciclo for sul numero di valori ottenuti
'duplica la riga, andando a scrivere in ogni riga duplicata uno dei valori splittati.
'I dati sono gestiti in collection:
'- la collection RowCollection rappresenta le righe ed è una collection di collection di tipo 'RowValuesCollection'
'- la collection RowValuesCollection contiene i valori dei campi per ogni riga
'Ho utilizzato le collection anche se probabilmente non sono il metodo più veloce (la velocità per ora
'non è un vincolo) perchè consentono di iterare facilmente sugli item e di gestire i valori come coppie
'key-value, con vantaggi dal punto di vista della comprensibilità del codice.
    
    Dim NAMED_CELL_FOR_TARGET_WB, _
        SOURCE_SHEET_NAME, _
        SOURCE_FIELD_WB, _
        SOURCE_FIELD_WELDING_MAP, _
        SOURCE_FIELD_JOINT_NO, _
        SOURCE_FIELD_WPS_NUMBER, _
        SOURCE_FIELD_WPS_REV, _
        SOURCE_FIELD_TO_BE_SPLITTED, _
        TARGET_SHEET_NAME, _
        TARGET_FIELD_WB, _
        TARGET_FIELD_WPS_NUMBER, _
        TARGET_FIELD_WPS_REV, _
        TARGET_FIELD_WELDING_MAP, _
        TARGET_FIELD_JOINT_NO As String
    Dim DELIMITER1, DELIMITER2 As String
    Dim SourceSheet, TargetSheet As Worksheet
    Dim SourceTable, TargetTable As Excel.ListObject
    Dim RowsCount As Integer
    Dim r, f, sf, tf, tf2, v, MyCell, MyCell2 As Variant
    Dim i As Integer
    Dim TargetWB As String
    Dim RowCollection, RowValuesCollection, CollectionClone As Collection
    Dim MyCellText As String
    Dim MyRow As Range
    Dim SourceFields As Collection
    Dim MyItem As String
    Dim StringToSplit, split1, split2 As Variant
    
    On Error GoTo ErrHandler

    '****SETTINGS****
    NAMED_CELL_FOR_TARGET_WB = "TargetWB" 'Nome della cella che contiene il nome del Welding Book da elaborare
    'Da dove prendere i dati (nome foglio + nomi dei campi della tabella):
    SOURCE_SHEET_NAME = "WPS"
    SOURCE_FIELD_WB = "_Welding_Book"
    SOURCE_FIELD_WELDING_MAP = "_Welding_map"
    SOURCE_FIELD_JOINT_NO = "_Joint_No."
    SOURCE_FIELD_WPS_NUMBER = "wps_number"
    SOURCE_FIELD_WPS_REV = "wps_rev"
    SOURCE_FIELD_TO_BE_SPLITTED = "_Welding_map" 'Deve essere uguale a uno dei campi elencati sopra
    'Dove scrivere i dati (nome foglio + nomi dei campi della tabella):
    TARGET_SHEET_NAME = "RiepilogoWBMultiMap"
    TARGET_FIELD_WB = "_Welding_Book"
    TARGET_FIELD_WPS_NUMBER = "wps_number"
    TARGET_FIELD_WPS_REV = "wps_rev"
    TARGET_FIELD_WELDING_MAP = "_Welding_map"
    TARGET_FIELD_JOINT_NO = "_Joint_No."
    'Delimitatori di primo e secondo livello utilizzati nel campo da splittare
    DELIMITER1 = ";"
    DELIMITER2 = ":"
    '****END SETTINGS****
    
    
    Set SourceSheet = ActiveWorkbook.Sheets(SOURCE_SHEET_NAME)
    Set SourceTable = SourceSheet.ListObjects(1)
    'Crea una collection con i campi di interesse, sui quali iterare in lettura
    Set SourceFields = New Collection
    SourceFields.Add SOURCE_FIELD_WB
    SourceFields.Add SOURCE_FIELD_WELDING_MAP
    SourceFields.Add SOURCE_FIELD_JOINT_NO
    SourceFields.Add SOURCE_FIELD_WPS_NUMBER
    SourceFields.Add SOURCE_FIELD_WPS_REV
    
    Set TargetSheet = ActiveWorkbook.Sheets(TARGET_SHEET_NAME)
    Set TargetTable = TargetSheet.ListObjects(1)
  
'**********************************************
'***LETTURA DEI DATI DALLA TABELLA SORGENTE****
'**********************************************
'I dati vengono memorizzati in una collection

    TargetWB = Range(NAMED_CELL_FOR_TARGET_WB)
    
    'Crea delle collection dove memorizzare i valori
    Set RowCollection = New Collection
    Set RowValuesCollection = New Collection
    
    'Scorre le righe della tabella sorgente per intercettare il WB di interesse
    For Each MyCell In SourceTable.ListColumns(SOURCE_FIELD_WB).Range.Cells
        Debug.Print MyCell.Text
        If MyCell.Text = TargetWB Then
            'Se la riga corrisponde al criterio di ricerca, la setta come oggetto
            Set MyRow = SourceTable.Range.Rows(MyCell.Row - SourceTable.Range.Row + 1)
            'Memorizza in una collection i campi di interesse per la riga selezionata
            For Each f In SourceFields
                MyItem = MyRow.Columns(SourceTable.ListColumns(f).Range.Column).Cells(1, 1).Text
                RowValuesCollection.Add key:=f, Item:=MyItem
            Next f
            'Fa lo split a due livelli del campo da splittare e memorizza i valori splittati
            'in diversi item della RowCollection
            StringToSplit = RowValuesCollection(SOURCE_FIELD_TO_BE_SPLITTED)
            StringToSplit = (Replace(StringToSplit, " ", "")) 'Rimuove spazi vuoti
            StringToSplit = (Replace(StringToSplit, Chr(10), "")) 'Rimuove gli "a capo"
            split1 = Split(StringToSplit, DELIMITER1)
            For i = LBound(split1) To UBound(split1)
                'Crea un clone della collection, la modifica, e la aggiunge come item di RowCollection
                Set CollectionClone = New Collection
                For Each f In SourceFields
                    CollectionClone.Add key:=f, Item:=RowValuesCollection(f)
                Next f
                If split1(i) <> "" Then
                    split2 = Split(split1(i), DELIMITER2)
                    Set CollectionClone = UpdateStringCollection(CollectionClone, SOURCE_FIELD_WELDING_MAP, split2(0))
                    If UBound(split2) > 0 Then
                        Set CollectionClone = UpdateStringCollection(CollectionClone, SOURCE_FIELD_JOINT_NO, split2(1))
                    End If
                    RowCollection.Add CollectionClone
                End If
            Next i
            
            'Svuota la collection coi valori della riga, sarà ripopolata alla prossima iterazione
            Set RowValuesCollection = New Collection
        End If
    Next MyCell
        
    
'********************************************************
'****SCRITTURA DEI DATI NELLA TABELLA DI DESTINAZIONE****
'********************************************************
'Scrive i dati nella tabella di destinazione effettuando un opportuno split

    'Cacella tutte le righe della tabella tranne la prima, per conservare le formule
    While TargetTable.ListRows.Count > 1
        TargetTable.ListRows(TargetTable.ListRows.Count).Delete
    Wend
    
    'Scansiona le righe memorizzate nella collection
    For Each r In RowCollection
        'Setta come riga di destinazione l'ultima riga della tabella target
        Set MyRow = TargetTable.DataBodyRange.Rows(TargetTable.ListRows.Count)
        'Scorre i campi di interesse
        For Each sf In SourceFields
            'Questo costrutto select setta la corrispondenza tra i campi sorgente e target, ci sono modi più intelligenti
            'ma per ora questo è il più veloce
            tf2 = ""
            Select Case sf
                Case SOURCE_FIELD_WB
                    tf = TARGET_FIELD_WB
                Case SOURCE_FIELD_WELDING_MAP
                    tf = TARGET_FIELD_WELDING_MAP
                Case SOURCE_FIELD_JOINT_NO
                    tf = TARGET_FIELD_JOINT_NO
                Case SOURCE_FIELD_WPS_NUMBER
                    tf = TARGET_FIELD_WPS_NUMBER
                Case SOURCE_FIELD_WPS_REV
                    tf = TARGET_FIELD_WPS_REV
                Case Else
                    tf = ""
            End Select
            If tf <> "" Then
                Set MyCell = MyRow.Columns(TargetTable.ListColumns(tf).Range.Column).Cells(1, 1)
                With MyCell
                 .NumberFormat = "@"
                 .value = r(sf)
                End With
            End If
        Next sf
        
        'Aggiunge una riga alla tabella
        TargetTable.ListRows.Add AlwaysInsert:=True
    Next r
    
    'Cancella l'ultima riga della tabella, che per come è impostato il ciclo sopra rimane sempre vuota
    TargetTable.ListRows(TargetTable.ListRows.Count).Delete
    
    'Fa un po' di pulizia
    Set RowCollection = Nothing
    Set RowValuesCollection = Nothing
    Set CollectionClone = Nothing
    
MyExit:
    Exit Sub

ErrHandler:
    MsgBox "Ops, si è verifcato un errore!" & vbCrLf & _
           "Err. n." & Err.Number & ": " & Err.Description
    Resume MyExit
    
End Sub
