Attribute VB_Name = "mInsertImage"
Option Explicit
Sub ProvaInsertImageInCell()
Attribute ProvaInsertImageInCell.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim TargetCell As Range
    Dim ImagePath As String
    
    Set TargetCell = ActiveCell
    ImagePath = "J:\180139-CMA_Spa-COLLABORAZIONE\PQR_e_WPS\JointSketchRepository\BWdissThk_PP+z1z2.jpg"
    
    InsertImageInCell TargetCell, ImagePath
    
End Sub
Sub PasteWPSImageAndData()
Attribute PasteWPSImageAndData.VB_ProcData.VB_Invoke_Func = "m\n14"
'NB: questa è una routine molto personalizzata sul foglio, sviluppata rapidamente e non pretende
'di essere fatta bene o di essere universale

'Attenzione: fa uso del modulo mRegexUtils creato da me che deve essere inserito nel file

Dim SourceSheetName As String, SourceLookUpColumnName As String
Dim SourceValueColumnName As String 'Per l'immagine
Dim SourceValue2ColumnName As String 'Per i dati del cianfrino
Dim TargetSheetName As String, TargetLookUpColumnName As String, TargetColumnName As String
Dim TargetLookUpColumnNumber As Integer, TargetColumnNumber As Integer
Dim TargetTableHeader As Range
Dim ColumnName As String, ColumnNumber As Integer
Dim ImageFolderPath As String
Dim SourceLookUpColumn As Range
Dim SourceValueColumn As Range, SourceValue2Column As Range
Dim Value2 As String
Dim TargetLookUpColumn As Range, TargetColumn As Range
Dim SourceRow As Variant
Dim MyCell As Range, TargetCell As Range
Dim ImageName As Variant 'può assumere un valore di errore se non va a buon fine il metodo Match
Dim ImagePath As String
Dim MyImage As Variant
Dim MyImageName As String
Dim TargetSheet As Excel.Worksheet
Dim UpdateSelected As Boolean, SelectedRange As Range
Dim FirstSelectedRow As Integer, LastSelectedRow As Integer

'On Error Resume Next

'************
'**SETTINGS**
'************
SourceSheetName = "WPS"
SourceLookUpColumnName = "wps_number"
SourceValueColumnName = "joint_sketch_file"
SourceValue2ColumnName = "joint_sketch_text_left"
ImageFolderPath = "J:\180139-CMA_Spa-COLLABORAZIONE\PQR_e_WPS\JointSketchRepository\"

'Per la destinazione è i nomi di colonna sono complicati, per cui è sufficiente inserirne una
'parte poi con le regular expression troverà la colonna corrispondente
TargetSheetName = "H217-21" 'E' solo una parte del nome, verrà confrontato con l'activesheet
TargetLookUpColumnName = "WPS-Nr."
TargetColumnName = "weld details"

'****************************
'**Individua il TargetSheet**
'****************************
Set TargetSheet = ActiveSheet
'Fa un confronto con il nome passato nelle impostazioni
If StrComp(regexMatch(TargetSheet.Name, TargetSheetName, False), TargetSheetName, vbTextCompare) <> 0 Then
    MsgBox "Hai tentato di lanciare al procedura su un foglio che non assomiglia a '" & TargetSheetName & _
           "', posizionati prima sul foglio di interesse oppure modifica le impostazioni nel codice VBA."
    GoTo MyExit
End If

'***********************************************************************
'**Definisce un parametro per gestire l'aggiornamento di una sola riga**
'***********************************************************************
'Se si fa una selezione multipla, allora viene aggiornata solo la selezione
UpdateSelected = False 'di default aggiorna tutto
If Selection.Cells.Count > 1 Then
    UpdateSelected = True
    Set SelectedRange = Selection
    FirstSelectedRow = SelectedRange.Row
    LastSelectedRow = FirstSelectedRow + SelectedRange.Rows.Count - 1
End If

'**********************************
'**FA PULIZIA PRIMA DI COMINCIARE**
'**********************************
'Cancella tutto solo se non c'è un aggiornamento parziale
If Not UpdateSelected Then
    'Scorre tutte le immagini nel foglio e le cancella
    For Each MyImage In TargetSheet.Shapes
        'conserva solo le immagini del template
        Debug.Print MyImage.Name
        MyImageName = MyImage.Name
        If MyImageName <> "Gruppieren 16" And MyImageName <> "Gruppieren 11" And MyImageName <> "Grafik 2" Then
            Debug.Print MyImageName
            MyImage.Delete
        End If
    Next MyImage
End If

'****************
'**ELABORAZIONE**
'****************
Set SourceLookUpColumn = Sheets(SourceSheetName).ListObjects(1).ListColumns(SourceLookUpColumnName).DataBodyRange
Set SourceValueColumn = Sheets(SourceSheetName).ListObjects(1).ListColumns(SourceValueColumnName).DataBodyRange
Set SourceValue2Column = Sheets(SourceSheetName).ListObjects(1).ListColumns(SourceValue2ColumnName).DataBodyRange

Set TargetTableHeader = TargetSheet.ListObjects(1).HeaderRowRange
For Each MyCell In TargetTableHeader.Cells
    'MyCell.Select 'solo per debug
    'Verifica se c'è corrispondenza e individua le colonnne
    If StrComp(regexMatch(MyCell.value, TargetLookUpColumnName, False), TargetLookUpColumnName, vbTextCompare) = 0 Then
        TargetLookUpColumnNumber = MyCell.Column - (TargetTableHeader.Cells(1.1).Column - 1)
    End If
    If StrComp(regexMatch(MyCell.value, TargetColumnName, False), TargetColumnName, vbTextCompare) = 0 Then
        TargetColumnNumber = MyCell.Column - (TargetTableHeader.Cells(1.1).Column - 1)
    End If
Next MyCell
Set TargetLookUpColumn = TargetSheet.ListObjects(1).ListColumns(TargetLookUpColumnNumber).DataBodyRange
Set TargetColumn = TargetSheet.ListObjects(1).ListColumns(TargetColumnNumber).DataBodyRange

'Scorre tutte le celle per cercare il valore di riferimento e inserire il contenuto
For Each MyCell In TargetLookUpColumn.Cells
    'Debug.Print MyCell.value
    'Cerca solo valori non nulli per evitare errori della funzione Match
    If Not MyCell.value = vbNullString Then
        SourceRow = Application.Match(MyCell.value, SourceLookUpColumn, 0)
        'Gestisce l'errore di valore non trovato
        If Not IsError(SourceRow) Then
            'Aggiorna immagine e dati sempre (se non c'è una selezione multipla)
            'oppure solo se la riga è nell'intervallo selezionato
            If Not UpdateSelected Or (MyCell.Row >= FirstSelectedRow And MyCell.Row <= LastSelectedRow) Then
                'Inserisce immagine
                ImageName = Application.Index(SourceValueColumn, SourceRow)
                ImagePath = ImageFolderPath & ImageName
                'Debug.Print TargetColumn.EntireColumn.Rows(MyCell.Row).Row
                Set TargetCell = TargetColumn.EntireColumn.Rows(MyCell.Row)
                InsertImageInCell TargetCell, ImagePath, 0.5
                
                'Inserisce i dati
                Value2 = Application.Index(SourceValue2Column, SourceRow)
                TargetCell.value = Value2
            End If
        End If
        
        
    End If
Next MyCell

MyExit:
    Exit Sub

End Sub

Sub InsertImageInCell(TargetCell As Range, ImagePath As String, Optional HCrop As Single = 0, Optional VCrop As Single = 0)
Attribute InsertImageInCell.VB_ProcData.VB_Invoke_Func = "m\n14"
'Inserisce un'immagine in una cella
'HCrop e VCrop sono Horizontal Crop e Vertical Crop

    Dim CellHeight As Single, CellWidth As Single
    Dim MyImage As ShapeRange
    Dim ImageWidth As Single
    
    CellHeight = TargetCell.Height
    CellWidth = TargetCell.Width
    
    'Seleziona la cella
    TargetCell.Select
    'Inserisce l'immagine
    ActiveSheet.Pictures.Insert(ImagePath).Select
    'Ridimensiona l'immagine sulla base dell'altezza della cella
    Set MyImage = Selection.ShapeRange
    MyImage.Height = CellHeight
    'Se l'immagine è più larga della cella, allora la ridimensiona anche in base alla larghezza della cella
    If MyImage.Width > CellWidth Then
        MyImage.Width = CellWidth
    End If
    
    'Taglia l'immagine in base al fattore di larghezza passato in ingresso
    If HCrop <> 0 Then
        ImageWidth = MyImage.Width
        'Le operazioni che seguono sono copiate e adattate da una registrazione macro, non ne ho comprensione
        'completa ma funziona
        MyImage.LockAspectRatio = msoFalse
        MyImage.IncrementLeft ImageWidth * HCrop
        MyImage.ScaleWidth HCrop, msoFalse, msoScaleFromTopLeft
        MyImage.PictureFormat.Crop.PictureWidth = ImageWidth
        MyImage.PictureFormat.Crop.PictureOffsetX = ImageWidth * HCrop / 2
        'Sposta leggermente l'immagine dai bordi della cella
        MyImage.IncrementTop 1
        MyImage.IncrementLeft -0.5
        'Blocca nuovamente le proporzioni
        MyImage.LockAspectRatio = msoTrue
      
    End If
   
End Sub
