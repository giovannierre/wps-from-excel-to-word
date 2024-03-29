Attribute VB_Name = "mWPS_from_excel_to_word"
Option Explicit
Global UserFormUpdaterRow As Integer
Sub process_wps()
Attribute process_wps.VB_ProcData.VB_Invoke_Func = "w\n14"
    
    Dim MySheet As Excel.Worksheet
    Dim SelectedRange As Excel.Range
    Dim MyCell As Excel.Range
    Dim MyAnswer As Variant
    Dim MultipleRows As Boolean
    Dim AllowDoEvents As Boolean
    
    On Error GoTo ErrHandler
    
    'Definisco un parametro per poter interrompere loop molto lunghi
    AllowDoEvents = False
    
    Set MySheet = ActiveWorkbook.Worksheets("WPS")
    
    Set SelectedRange = Selection
    If SelectedRange.Rows.Count > 1 Then
        MyAnswer = MsgBox("Hai selezionato un intervallo multiplo, saranno processate tutte le righe, potrebbe richiedere " & _
                         "molto tempo, sei sicuro di voler proseguire?", vbOKCancel)
        If MyAnswer = vbCancel Then GoTo MyExit
        If SelectedRange.Rows.Count > 10 Then
            MyAnswer = MsgBox("Il numero di righe selezionato � superiore a 10, potrebbe essere un errore di selezione, " & _
                       "prosegui solo se sei sicuro di quello che stai facendo!" & vbCrLf & _
                       "[in caso di problemi, premere ESC per uscire forzatamente dal loop]", vbOKCancel)
            If MyAnswer = vbCancel Then GoTo MyExit
            AllowDoEvents = True
        End If
        MultipleRows = True
     Else
        MultipleRows = False
    End If
    
    'Scorre le righe dell'intervallo selezionato
    For Each MyCell In SelectedRange.Columns(1).Cells
        Call read_wps_data(MySheet, MyCell, MultipleRows)
        If AllowDoEvents Then DoEvents
    Next MyCell
    
    Set MySheet = Nothing
    
MyExit:
    Exit Sub

ErrHandler:
    MsgBox "Ops, si � verifcato un errore!" & vbCrLf & _
           "Err. n." & Err.Number & ": " & Err.Description
    Resume MyExit
    
End Sub

Sub read_wps_data(MySheet As Excel.Worksheet, CurrentCell As Excel.Range, AutomaticSave As Boolean)
Attribute read_wps_data.VB_ProcData.VB_Invoke_Func = "w\n14"

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
    Dim FullNamePdfExport, FullNameWordExport As String
    Dim MyAnswer As Variant
    Dim FileName As String
    Dim NAMED_CELL_FOR_IMAGE_PATH, _
        NAMED_CELL_FOR_TEMPLATE_PATH, _
        NAMED_CELL_FOR_PDF_EXPORT_PATH, _
        NAMED_CELL_FOR_WORD_EXPORT_PATH As String
    Dim TABLE_FIELD_FOR_IMAGE_FILENAME As String
    Dim DirToCheck As String
    Dim DirSplit As Variant
    Dim cc As ContentControl
    Dim ImageFilePath As String
    Dim DesiredHeight, DesiredWidth As Single
    Dim FactorH, FactorW, Factor As Single
        
    On Error GoTo ErrHandler

'##################################################################
'               ************
'               **SETTINGS**
'               ************
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
'             **FINE SETTINGS**
'##################################################################
    
'********************
'**LETTURA DEI DATI**
'********************
    'Definisce l'oggeto ListObject (tabella di Excel) sul quale lavorare
    Set MyTable = MySheet.ListObjects(1)
    
    ' Legge le intestazioni di colonna della tabella e le mette in una collection
    Set PropertyName = New Collection
    For Each MyCell In MyTable.HeaderRowRange.Cells
     MyCellText = MyCell.Text
     'Esclude i campi che iniziano con "_" (underscore)
     If InStr(1, MyCellText, "_") <> 1 Then
        PropertyName.Add MyCell.Text
     End If
    Next
    
    Set PropertyValue = New Collection
    Set MyRow = MyTable.Range.Rows(CurrentCell.Row - MyTable.Range.Row + 1)
    
    For Each MyCell In MyRow.Cells
        PropertyValue.Add key:=MyTable.HeaderRowRange.Columns(MyCell.Column).Text, _
                          Item:=Replace(MyCell.Text, Chr(10), Chr(13)) 'la sostituzione serve per fare andare a capo correttamente word
    Next
   
'**********************
'**SCRITTURA DEI DATI**
'**********************
     
    'Seleziona il file di template del documento Word dalla cella NAMED_CELL_FOR_TEMPLATE_PATH oppure, _
    'se vuoto, lo fa scegliere all'utente
    TargetDocumentPath = Range(NAMED_CELL_FOR_TEMPLATE_PATH)
    If TargetDocumentPath = "" Then
        With Application.FileDialog(msoFileDialogOpen)
            .AllowMultiSelect = False
            .InitialFileName = ActiveWorkbook.Path & Application.PathSeparator
            .Show
            TargetDocumentPath = .SelectedItems(1)
        End With
    End If
       
    'Apre il file se non � gi� aperto
    'NB: usa una funzione ausiliaria che non � built-in ma si trova in un altro modulo (copiata da sito MS)
    If IsFileOpen(TargetDocumentPath) Then
        'File gi� aperto, ne prende il controllo
        Set WordApp = GetObject(TargetDocumentPath).Application
    Else
        'File non aperto, crea un nuovo oggetto Applicazione di Word e apre il file
        Set WordApp = CreateObject("Word.Application")
        WordApp.documents.Open FileName:=TargetDocumentPath, ReadOnly:=True
    End If
    'Debug.Print WordApp.documents(0).Name
    WordApp.documents(TargetDocumentPath).Activate
    WordApp.Visible = True
    
    Set TargetDocument = WordApp.ActiveDocument
    
    StartTime = Timer
    
    'Inserisce i valori nel file di Word, attraverso la creazione (se serve) e popolamento
    'delle CustomProperties
    Call CreateCustomProperties(TargetDocument, PropertyName, PropertyValue)
    
'*****************************
'**INSERIMENTO DELL'IMMAGINE**
'*****************************
    
    'Inserisce l'immagine mediante un content control di tipo Picture
    'Se la cella nel campo TABLE_FIELD_FOR_IMAGE_FILENAME contiene i ":" (due punti) significa che �
    'indicato il percorso completo e si tiene buono quello, altrimenti si aggiunge il path specificato
    'nella cella NAMED_CELL_FOR_IMAGE_PATH
    ImageFilePath = PropertyValue(TABLE_FIELD_FOR_IMAGE_FILENAME)
    If InStr(1, ImageFilePath, ":") < 1 Then
        ImageFilePath = Range(NAMED_CELL_FOR_IMAGE_PATH).Text & ImageFilePath
    End If
    
    Set cc = TargetDocument.ContentControls(1)
    If cc.Type = wdContentControlPicture Then
        If cc.Range.InlineShapes.Count > 0 Then
            cc.Range.InlineShapes(1).Delete
        End If
        TargetDocument.InlineShapes.AddPicture _
            FileName:=ImageFilePath, _
            LinkToFile:=False, Range:=cc.Range
        
        'Riconsidera l'oggetto per ridimensionarlo:
        Set cc = TargetDocument.ContentControls(1)
        
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
    
'****************************
'****SALVATAGGIO DEI FILE****
'****************************
    
    'Attenzione: perch� funzionino le costanti di Word bisogna aggiungere
    'ai riferimenti la libreria "Microsoft Word x.x Object Library"
    
    'Definisce il nome del file con numero e revisione della WPS, per eventual salvataggio
    FileName = Replace("WPS_" & PropertyValue("wps_number") & "_rev" & PropertyValue("wps_rev"), _
                            "/", "-")
        
'******************************
'**SALVATAGGIO IN FORMATO pdf**
'******************************

    'Chiede se salvare o no il file pdf
    If AutomaticSave Then
        MyAnswer = vbYes
    Else
        MyAnswer = MsgBox("Vuoi salvare il documento in pdf?", vbYesNo)
    End If
        
    If MyAnswer = vbYes Then
        'Cerca un percorso di default memorizzato nel foglio per il salvataggio del pdf
        FullNamePdfExport = Range(NAMED_CELL_FOR_PDF_EXPORT_PATH)
        'Se non � specificato un percorso, allora utilizza il percorso del template
        If FullNamePdfExport = "" Then
            FullNamePdfExport = Replace(TargetDocument.FullName, TargetDocument.Name, "")
        End If
        'Controlla che la directory esista, tramite una funzione ausiliaria, che provveda anche a crearla, se confermato.
        If Not DirExists(FullNamePdfExport, False, False, True) Then
            MsgBox "Procedura annullata!"
            GoTo MyExit
        End If
        
        FullNamePdfExport = FullNamePdfExport & FileName
        'Salva il file in pdf
        TargetDocument.ExportAsFixedFormat OutputFileName:= _
            FullNamePdfExport & ".pdf", _
            ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
            wdExportOptimizeForPrint, Range:=wdExportAllDocument, Item:= _
            wdExportDocumentContent, IncludeDocProps:=False, KeepIRM:=True, _
            CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
            BitmapMissingFonts:=True, UseISO19005_1:=False
    End If

'*******************************
'**SALVATAGGIO IN FORMATO Word**
'*******************************
    'Chiede se salvare o no il file word compilato
    If AutomaticSave Then
        MyAnswer = vbYes
    Else
        MyAnswer = MsgBox("Vuoi appiattire e salvare il documento Word?", vbYesNo)
    End If
    
    If MyAnswer = vbYes Then
        'Cerca un percorso di default memorizzato nel foglio per il salvataggio del file di Word
        FullNameWordExport = Range(NAMED_CELL_FOR_WORD_EXPORT_PATH)
        'Se non � specificato un percorso, allora utilizza il percorso del template
        If FullNameWordExport = "" Then
            FullNameWordExport = Replace(TargetDocument.FullName, TargetDocument.Name, "")
        End If
        'Controlla che la directory esista, tramite una funzione ausiliaria, che provveda anche a crearla, se confermato.
        If Not DirExists(FullNameWordExport, False, False, True) Then
            MsgBox "Procedura annullata!"
            GoTo MyExit
        End If
        
        FullNameWordExport = FullNameWordExport & FileName
        'Appiattisce tutti i codici di campo nel documento
        TargetDocument.Fields.Unlink
        TargetDocument.SaveAs2 FileName:=FullNameWordExport & ".docx", FileFormat:=wdFormatXMLDocument
    End If
    
    If AutomaticSave Then
        TargetDocument.Close
        WordApp.Quit
    End If
    Set TargetDocument = Nothing
    Set WordApp = Nothing
    
MyExit:
    Exit Sub
    
ErrHandler:
    MsgBox "Ops, si � verifcato un errore!" & vbCrLf & _
           "Err. n." & Err.Number & ": " & Err.Description
    Resume MyExit

End Sub

Sub ShowUserForm()
    UserForm1.Show vbModeless
End Sub

Sub UpdateUserForm1(CurrentSelection As Range)
    
    On Error GoTo ErrHandler
    
    'Se non � selezionato l'update continuo, allora evidenzia con il colore
    'il fatto che l'immagine mostrata corrisponda a quella delle riga selezionata
    If Not UserForm1.chkAutomaticUpdate Then
        If CurrentSelection.Row = UserFormUpdaterRow Then
            UserForm1.BackColor = vbGreen
        Else
            UserForm1.BackColor = vbRed
        End If
    Else
    'Altrimenti se � selezionato l'update continuo, allora chiama la procedura di update
    'dello userform1, solo se si cambia riga
        If CurrentSelection.Row <> UserFormUpdaterRow Then
            Call UserForm1.cmdUpdate_Click
        End If
    End If

Exit Sub

ErrHandler:
    MsgBox "Ops! Si � verificato un errore nella routine 'SelectionChange' del foglio" & vbCrLf & _
            "Err. n." & Err.Number & ": " & Err.Description

End Sub
