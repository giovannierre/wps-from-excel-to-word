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
Sub PasteWPSImage()
Attribute PasteWPSImage.VB_ProcData.VB_Invoke_Func = "m\n14"
'NB: questa è una routine molto personalizzata sul foglio, sviluppata rapidamente e non pretende
'di essere fatta bene o di essere universale

Dim SourceSheetName As String, SourceLookUpColumnName As String, SourceValueColumnName As String
Dim TargetSheetName As String, TargetLookUpColumnNumber As Integer, TargetColumnNumber As Integer
Dim ImageFolderPath As String
Dim SourceLookUpColumn As Range, SourceValueColumn As Range
Dim TargetLookUpColumn As Range, TargetColumn As Range
Dim MyCell As Range
Dim ImageName As Variant 'può assumere un valore di errore se non va a buon fine il metodo Match
Dim ImagePath As String

'On Error Resume Next

'************
'**SETTINGS**
'************
SourceSheetName = "WPS"
SourceLookUpColumnName = "wps_number"
SourceValueColumnName = "joint_sketch_file"
ImageFolderPath = "J:\180139-CMA_Spa-COLLABORAZIONE\PQR_e_WPS\JointSketchRepository\"

'Per la destinazione è più facile definire i numeri di colonna che i nomi, che son troppo complicati
TargetSheetName = "H217-21_110"
TargetLookUpColumnNumber = 3
TargetColumnNumber = 2

'****************
'**ELABORAZIONE**
'****************
Set SourceLookUpColumn = Sheets(SourceSheetName).ListObjects(1).ListColumns(SourceLookUpColumnName).DataBodyRange
Set SourceValueColumn = Sheets(SourceSheetName).ListObjects(1).ListColumns(SourceValueColumnName).DataBodyRange

Set TargetLookUpColumn = Sheets(TargetSheetName).ListObjects(1).ListColumns(TargetLookUpColumnNumber).DataBodyRange
Set TargetColumn = Sheets(TargetSheetName).ListObjects(1).ListColumns(TargetColumnNumber).DataBodyRange

For Each MyCell In TargetLookUpColumn.Cells
    'Debug.Print MyCell.value
    
    'Cerca solo valori non nulli per evitare errori della funzione Match
    If Not MyCell.value = vbNullString Then
        ImageName = Application.Index(SourceValueColumn, _
                    Application.Match(MyCell.value, SourceLookUpColumn, 0))
        'Gestisce l'errore di valore non trovato
        If Not IsError(ImageName) Then
            Debug.Print ImageName
            ImagePath = ImageFolderPath & ImageName
            Debug.Print TargetColumn.EntireColumn.Rows(MyCell.Row).Row
            InsertImageInCell TargetColumn.EntireColumn.Rows(MyCell.Row), ImagePath
        End If
    End If
Next MyCell

End Sub

Sub InsertImageInCell(TargetCell As Range, ImagePath As String)
Attribute InsertImageInCell.VB_ProcData.VB_Invoke_Func = "m\n14"
'Inserisce un'immagine in una cella
    Dim CellHeight As Single, CellWidth As Single
    
    CellHeight = TargetCell.Height
    CellWidth = TargetCell.Width
    
    'Seleziona la cella
    TargetCell.Select
    'Inserisce l'immagine
    ActiveSheet.Pictures.Insert(ImagePath).Select
    'Ridimensiona l'immagine sulla base dell'altezza della cella
    Selection.ShapeRange.Height = CellHeight
    'Se l'immagine è più larga della cella, allora la ridimensiona anche in base alla larghezza della cella
    If Selection.ShapeRange.Width > CellWidth Then
        Selection.ShapeRange.Width = CellWidth
    End If
Exit Sub
    Selection.ShapeRange.Height = 263.6220472441
    Selection.ShapeRange.Height = 260.7874015748
    Selection.ShapeRange.Height = 257.9527559055
    Selection.ShapeRange.Height = 255.1181102362
    Selection.ShapeRange.Height = 252.2834645669
    Selection.ShapeRange.Height = 249.4488188976
    Selection.ShapeRange.Height = 246.6141732283
    Selection.ShapeRange.Height = 243.7795275591
    Selection.ShapeRange.Height = 240.9448818898
    Selection.ShapeRange.Height = 238.1102362205
    Selection.ShapeRange.Height = 235.2755905512
    Selection.ShapeRange.Height = 232.4409448819
    Selection.ShapeRange.Height = 229.6062992126
    Selection.ShapeRange.Height = 226.7716535433
    Selection.ShapeRange.Height = 223.937007874
    Selection.ShapeRange.Height = 221.1023622047
    Selection.ShapeRange.Height = 218.2677165354
    Selection.ShapeRange.Height = 215.4330708661
    Selection.ShapeRange.Height = 212.5984251969
    Selection.ShapeRange.Height = 209.7637795276
    Selection.ShapeRange.Height = 206.9291338583
    Selection.ShapeRange.Height = 204.094488189
    Selection.ShapeRange.Height = 201.2598425197
    Selection.ShapeRange.Height = 198.4251968504
    Selection.ShapeRange.Height = 195.5905511811
    Selection.ShapeRange.Height = 192.7559055118
    Selection.ShapeRange.Height = 189.9212598425
    Selection.ShapeRange.Height = 187.0866141732
    Selection.ShapeRange.Height = 184.2519685039
    Selection.ShapeRange.Height = 181.4173228346
    Selection.ShapeRange.Height = 178.5826771654
    Selection.ShapeRange.Height = 175.7480314961
    Selection.ShapeRange.Height = 172.9133858268
    Selection.ShapeRange.Height = 170.0787401575
    Selection.ShapeRange.Height = 167.2440944882
    Selection.ShapeRange.Height = 164.4094488189
    Selection.ShapeRange.Height = 161.5748031496
    Selection.ShapeRange.Height = 158.7401574803
    Selection.ShapeRange.Height = 155.905511811
    Selection.ShapeRange.Height = 153.0708661417
    Selection.ShapeRange.Height = 150.2362204724
    Selection.ShapeRange.Height = 147.4015748031
    Selection.ShapeRange.Height = 144.5669291339
    Selection.ShapeRange.Height = 141.7322834646
    Selection.ShapeRange.Height = 138.8976377953
    Selection.ShapeRange.Height = 136.062992126
    Selection.ShapeRange.Height = 133.2283464567
    Selection.ShapeRange.Height = 130.3937007874
    Selection.ShapeRange.Height = 127.5590551181
    Selection.ShapeRange.Height = 124.7244094488
    Selection.ShapeRange.Height = 121.8897637795
    Selection.ShapeRange.Height = 119.0551181102
    Selection.ShapeRange.Height = 116.2204724409
    Selection.ShapeRange.Height = 113.3858267717
    Selection.ShapeRange.Height = 110.5511811024
    Selection.ShapeRange.Width = 238.1102362205
    Selection.ShapeRange.Width = 235.2755905512
    Selection.ShapeRange.Width = 232.4409448819
    Selection.ShapeRange.Width = 229.6062992126
    Selection.ShapeRange.Width = 226.7716535433
    Selection.ShapeRange.Width = 223.937007874
    Selection.ShapeRange.Width = 221.1023622047
    Selection.ShapeRange.Width = 218.2677165354
    Selection.ShapeRange.Width = 215.4330708661
    Selection.ShapeRange.Width = 212.5984251969
    Selection.ShapeRange.Width = 209.7637795276
    Selection.ShapeRange.Width = 206.9291338583
    Selection.ShapeRange.Width = 204.094488189
    Selection.ShapeRange.Width = 201.2598425197
    Selection.ShapeRange.Width = 198.4251968504
    Selection.ShapeRange.Width = 195.5905511811
    Selection.ShapeRange.Width = 192.7559055118
    Selection.ShapeRange.Width = 189.9212598425
    Selection.ShapeRange.Width = 187.0866141732
    Selection.ShapeRange.Width = 184.2519685039
    Selection.ShapeRange.Width = 181.4173228346
    Selection.ShapeRange.Width = 178.5826771654
    Selection.ShapeRange.Width = 175.7480314961
    Selection.ShapeRange.Width = 172.9133858268
    Selection.ShapeRange.Width = 170.0787401575
    Selection.ShapeRange.Width = 167.2440944882
    Selection.ShapeRange.Width = 164.4094488189
    Selection.ShapeRange.Width = 161.5748031496
    Selection.ShapeRange.Width = 158.7401574803
    Selection.ShapeRange.Width = 155.905511811
    Selection.ShapeRange.Width = 153.0708661417
    Selection.ShapeRange.Width = 150.2362204724
    Selection.ShapeRange.Width = 147.4015748031
    Selection.ShapeRange.Width = 144.5669291339
    Selection.ShapeRange.Width = 141.7322834646
    Selection.ShapeRange.Width = 138.8976377953
    Selection.ShapeRange.Width = 136.062992126
    Selection.ShapeRange.Width = 133.2283464567
    Selection.ShapeRange.Width = 130.3937007874
    Selection.ShapeRange.Width = 127.5590551181
    Selection.ShapeRange.Width = 124.7244094488
    Selection.ShapeRange.Width = 121.8897637795
    Selection.ShapeRange.Width = 119.0551181102
    Selection.ShapeRange.Width = 116.2204724409
    Selection.ShapeRange.Width = 113.3858267717
    Selection.ShapeRange.Width = 110.5511811024
    Selection.ShapeRange.Width = 107.7165354331
    Selection.ShapeRange.Width = 104.8818897638
    Selection.ShapeRange.IncrementTop 0.6522047244
    Selection.ShapeRange.IncrementTop 0.6522047244
    Selection.ShapeRange.IncrementTop 0.6522047244
    Selection.ShapeRange.IncrementTop 0.6522047244
    Selection.ShapeRange.IncrementTop 0.6522047244
    Selection.ShapeRange.IncrementTop 0.6522047244
    Selection.ShapeRange.IncrementTop 0.6522047244
    Selection.ShapeRange.IncrementTop 0.6522047244
    Selection.ShapeRange.IncrementTop 0.6522047244
    Selection.ShapeRange.IncrementTop 0.6522047244
    Selection.ShapeRange.IncrementTop 0.6522047244
    Selection.ShapeRange.IncrementTop 0.6522047244
    Selection.ShapeRange.IncrementTop 0.6522047244
    Selection.ShapeRange.IncrementTop 0.6522047244
    Selection.ShapeRange.IncrementTop 0.6522047244
    Selection.ShapeRange.IncrementTop 0.6522047244
    Selection.ShapeRange.IncrementTop 0.6522047244
    Selection.ShapeRange.IncrementTop 0.6522047244
    Selection.ShapeRange.IncrementTop 0.6522047244
    Selection.ShapeRange.IncrementTop 0.6522047244
    Selection.ShapeRange.IncrementTop 0.6522047244
    Selection.ShapeRange.IncrementTop 0.6522047244
    Selection.ShapeRange.IncrementTop 0.6522047244
    Selection.ShapeRange.IncrementTop 0.6522047244
    Selection.ShapeRange.IncrementTop 0.6522047244
End Sub
