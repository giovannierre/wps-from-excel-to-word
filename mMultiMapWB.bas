Attribute VB_Name = "mMultiMapWB"
Option Explicit

'Le routine in questo modulo servono per elaborare welding book composti da molte Welding Map,
'per i quali nel foglio WPS sono specificati in un'unica cella Welding Map e Join No.,
'esempio: "WM001: W1, W2; WM002: W1, W3"
'Le diverse welding map devono essere separate da ";" (punto e virgola).

Sub ElaborateWB()
Attribute ElaborateWB.VB_ProcData.VB_Invoke_Func = "e\n14"
    
    Dim NAMED_CELL_FOR_TARGET_WB, _
        SOURCE_SHEET_NAME, _
        SOURCE_FIELD_WB, _
        SOURCE_FIELD_WELDING_MAP_AND_JOINT_NO, _
        SOURCE_FIELD_WPS_NUMBER, _
        SOURCE_FIELD_WPS_REV, _
        TARGET_SHEET_NAME, _
        TARGET_FIELD_WB, _
        TARGET_FIELD_WPS_NUMBER, _
        TARGET_FIELD_WPS_REV, _
        TARGET_FIELD_WELDING_MAP, _
        TARGET_FIELD_JOINT_NO As String
    Dim SourceSheet, TargetSheet As Worksheet
    Dim SourceTable, TargetTable As Excel.ListObject
    Dim RowsCount As Integer
    Dim r, f, sf, tf, tf2, v, MyCell, MyCell2 As Variant
    Dim i As Integer
    Dim TargetWB As String
    Dim RowCollection, RowValuesCollection As Collection
    Dim MyCellText As String
    Dim MyRow As Range
    Dim SourceFields As Collection
    Dim MyItem As String


    '****SETTINGS****
    NAMED_CELL_FOR_TARGET_WB = "TargetWB" 'Nome della cella che contiene il nome del Welding Book da elaborare
    SOURCE_SHEET_NAME = "WPS"
    SOURCE_FIELD_WB = "_Welding_Book"
    SOURCE_FIELD_WELDING_MAP_AND_JOINT_NO = "_Welding_map"
    SOURCE_FIELD_WPS_NUMBER = "wps_number"
    SOURCE_FIELD_WPS_REV = "wps_rev"
    TARGET_SHEET_NAME = "RiepilogoWBMultiMap"
    TARGET_FIELD_WB = "_Welding_Book"
    TARGET_FIELD_WPS_NUMBER = "wps_number"
    TARGET_FIELD_WPS_REV = "wps_rev"
    TARGET_FIELD_WELDING_MAP = "_Welding_map"
    TARGET_FIELD_JOINT_NO = "_Joint_No."
    
    Set SourceSheet = ActiveWorkbook.Sheets(SOURCE_SHEET_NAME)
    Set SourceTable = SourceSheet.ListObjects(1)
    'Crea una collection con i campi di interesse, sui quali iterare in lettura
    Set SourceFields = New Collection
    SourceFields.Add SOURCE_FIELD_WB
    SourceFields.Add SOURCE_FIELD_WELDING_MAP_AND_JOINT_NO
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
            'Memorizza i campi di interesse per la riga selezionata in una collection
            For Each f In SourceFields
                MyItem = MyRow.Columns(SourceTable.ListColumns(f).Range.Column).Cells(1, 1).Text
                RowValuesCollection.Add key:=f, Item:=Replace(MyItem, Chr(10), Chr(13))
            Next f
            RowCollection.Add RowValuesCollection
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
                Case SOURCE_FIELD_WELDING_MAP_AND_JOINT_NO
                    tf = TARGET_FIELD_WELDING_MAP
                    tf2 = TARGET_FIELD_JOINT_NO
                Case SOURCE_FIELD_WPS_NUMBER
                    tf = TARGET_FIELD_WPS_NUMBER
                Case SOURCE_FIELD_WPS_REV
                    tf = TARGET_FIELD_WPS_REV
                Case Else
                    tf = ""
            End Select
            If tf <> "" And tf2 = "" Then
                Set MyCell = MyRow.Columns(TargetTable.ListColumns(tf).Range.Column).Cells(1, 1)
                With MyCell
                 .NumberFormat = "@"
                 .Value = r(sf)
                End With
                
                'Se è stato specificato un secondo campo, effettua lo split
             'ElseIf tf2 <> "" Then
                
             
            End If
        Next sf
        
        'Aggiunge una riga alla tabella
        TargetTable.ListRows.Add AlwaysInsert:=True
    Next r
    
    'Cancella l'ultima riga della tabella, che per come è impostato il ciclo sopra rimane sempre vuota
    TargetTable.ListRows(TargetTable.ListRows.Count).Delete
    
    
End Sub
