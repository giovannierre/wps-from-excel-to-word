Attribute VB_Name = "mWPS_from_excel_to_word"
Option Explicit
Global UserFormUpdaterRow As Integer


Sub read_wps_data()
Attribute read_wps_data.VB_ProcData.VB_Invoke_Func = "w\n14"

    Dim MySheet As Excel.Worksheet
    Dim MyTable As Excel.ListObject
    Dim MyCell, MyRow As Range
    Dim PropertyName As Collection
    Dim PropertyValue As Collection
    Dim TargetDocument As Object
    Dim WordApp As Object
    Dim TargetDocumentPath As String
    Dim StartTime
    Dim MyCellText As String
    Dim IMAGE_HEIGHT, IMAGE_WIDTH As Single
    Dim FileNameExport As String
    Dim MyAnswer As Variant
    Dim fld As Variant
    
    On Error GoTo ErrHandler
    
    '****SETTINGS****
    'Dimensioni massime immagine giunto
    IMAGE_HEIGHT = 3.5
    IMAGE_WIDTH = 8.3
    
        
    Set MySheet = ActiveWorkbook.Worksheets("WPS")
    
    Set MyTable = MySheet.ListObjects(1)
    
    ' Legge le intestazioni di colonna della tabella e le mette in un array
    Set PropertyName = New Collection
    For Each MyCell In MyTable.HeaderRowRange.Cells
     MyCellText = MyCell.Text
     'Esclude i campi che iniziano con "_" (underscore)
     If InStr(1, MyCellText, "_") <> 1 Then
        PropertyName.Add MyCell.Text
     End If
    Next
    
    Set PropertyValue = New Collection
    Set MyRow = MyTable.Range.Rows(ActiveCell.Row - MyTable.Range.Row + 1)
    
    For Each MyCell In MyRow.Cells
        PropertyValue.Add key:=MyTable.HeaderRowRange.Columns(MyCell.Column).Text, _
                          Item:=Replace(MyCell.Text, Chr(10), Chr(13)) 'la sostituzione serve per fare andare a capo correttamente word
    Next
        
    'Crea un oggetto Applicazione di Word
    Set WordApp = CreateObject("Word.Application")
    
    'Seleziona il file di template dalla cella "TemplateFullPath" oppure, se vuoto, lo fa scegliere all'utente
    TargetDocumentPath = MySheet.Range("TemplateFullPath")
    If TargetDocumentPath = "" Then
        With Application.FileDialog(msoFileDialogOpen)
            .AllowMultiSelect = False
            .InitialFileName = ActiveWorkbook.Path & Application.PathSeparator
            .Show
            TargetDocumentPath = .SelectedItems(1)
        End With
    End If
       
    'Apre il file non è già aperto
    'NB: usa una funzione ausiliaria che non è built-in ma si trova in un altro modulo (copiata da sito MS)
    If Not IsFileOpen(TargetDocumentPath) Then
            WordApp.documents.Open FileName:=TargetDocumentPath, ReadOnly:=True
        Else
            Set WordApp = GetObject(TargetDocumentPath).Application
    End If
    'Debug.Print WordApp.documents(0).Name
    WordApp.documents(TargetDocumentPath).Activate
    WordApp.Visible = True
    
    Set TargetDocument = WordApp.ActiveDocument
    
    StartTime = Timer
    
    'Inserisce i valori nel file di Word, attraverso la creazione (se serve) e popolamento
    'delle CustomProperties
    Call CreateCustomProperties(TargetDocument, PropertyName, PropertyValue)
    
    'Inserisce l'immagine mediante un content control di tipo Picture
    Dim cc As ContentControl
    Dim ImageFilePath As String
    
    'Se la cella "joint_sketch_file" contiene i ":" (due punti) significa che è indicato il percorso completo
    'e si tiene buono quello, altrimenti si aggiunge il path specificato nella cella "ImagePath"
    ImageFilePath = PropertyValue("joint_sketch_file")
    If InStr(1, ImageFilePath, ":") < 1 Then
        ImageFilePath = MySheet.Range("ImagePath").Text & ImageFilePath
    End If
    
    Set cc = TargetDocument.ContentControls(1)
    If cc.Type = wdContentControlPicture Then
        If cc.Range.InlineShapes.Count > 0 Then
            cc.Range.InlineShapes(1).Delete
        End If
        TargetDocument.InlineShapes.AddPicture _
            FileName:=ImageFilePath, _
            linktofile:=False, Range:=cc.Range
        
        'Riconsidera l'oggetto per ridimensionarlo:
        Set cc = TargetDocument.ContentControls(1)
        
        Dim DesiredHeight, DesiredWidth As Single
        Dim FactorH, FactorW, Factor As Single
        
        DesiredHeight = Application.CentimetersToPoints(IMAGE_HEIGHT)
        DesiredWidth = Application.CentimetersToPoints(IMAGE_WIDTH)
        With cc.Range.InlineShapes(1)
            FactorH = DesiredHeight / .Height
            FactorW = DesiredWidth / .Width
            Factor = IIf(FactorH < FactorW, FactorH, FactorW)
            .Height = .Height * Factor
            .Width = .Width * Factor
        End With
    End If
    
    TargetDocument.Fields.Update
    
    Debug.Print "Elapsed time: " & Timer - StartTime
    
    'Esporta il file in pdf:
    'Attenzione: perchè funzionino le costanti di Word bisogna aggiungere
    'ai riferimenti la libreria "Microsoft Word x.x Object Library"
    Dim FileName As String
    
    'Definisce il nome del file con numero e revisione della WPS
    FileName = Replace("WPS_" & PropertyValue("wps_number") & "_rev" & PropertyValue("wps_rev"), _
                            "/", "-")
    'Cerca un percorso di default memorizzato nel foglio
    FileNameExport = MySheet.Range("SavePdfPath")
        'Se non è specificato un percorso, allora utilizza il percorso del template
        If FileNameExport = "" Then
            FileNameExport = Replace(TargetDocument.FullName, TargetDocument.Name, "")
        End If
        FileNameExport = FileNameExport & FileName
    
    'Chiede se salvare o no il file pdf
    MyAnswer = MsgBox("Vuoi salvare il documento in pdf?", vbYesNo)
        
    If MyAnswer = vbYes Then
        TargetDocument.ExportAsFixedFormat OutputFileName:= _
            FileNameExport & ".pdf", _
            ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
            wdExportOptimizeForPrint, Range:=wdExportAllDocument, Item:= _
            wdExportDocumentContent, IncludeDocProps:=False, KeepIRM:=True, _
            CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
            BitmapMissingFonts:=True, UseISO19005_1:=False
    End If
    
    MyAnswer = MsgBox("Vuoi appiattire e salvare il documento Word?", vbYesNo)
    
    If MyAnswer = vbYes Then
        'Appiattisce i codici di campo
        For Each fld In TargetDocument.Fields
            fld.Unlink
        Next
        Set fld = Nothing
        TargetDocument.SaveAs2 FileName:=FileNameExport & ".docx", FileFormat:=wdFormatXMLDocument
    End If
    

MyExit:
    Exit Sub
    
ErrHandler:
    MsgBox "Ops, si è verifcato un errore!" & vbCrLf & _
           "Err. n." & Err.Number & ": " & Err.Description
    Resume MyExit

End Sub

Sub ShowUserForm()
    UserForm1.Show vbModeless
End Sub

Sub UpdateUserForm1(CurrentSelection As Range)
    
    On Error GoTo ErrHandler
    
    'Se non è selezionato l'update continuo, allora evidenzia con il colore
    'il fatto che l'immagine mostrata corrisponda a quella delle riga selezionata
    If Not UserForm1.chkAutomaticUpdate Then
        If CurrentSelection.Row = UserFormUpdaterRow Then
            UserForm1.BackColor = vbGreen
        Else
            UserForm1.BackColor = vbRed
        End If
    Else
    'Altrimenti se è selezionato l'update continuo, allora chiama la procedura di update
    'dello userform1
        Call UserForm1.cmdUpdate_Click
    End If

Exit Sub

ErrHandler:
    MsgBox "Ops! Si è verificato un errore nella routine 'SelectionChange' del foglio" & vbCrLf & _
            "Err. n." & Err.Number & ": " & Err.Description

End Sub
