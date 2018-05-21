VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Show Joint Sketch"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6930
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CheckBox1_Click()

End Sub

Private Sub chkAutomaticUpdate_Change()
    
    Dim MyAnswer As Variant
    
    'Chiede conferma all'utente e se non c'è conferma esce dalla routine
    If chkAutomaticUpdate Then
        MyAnswer = MsgBox("Attenzione! L'update automatico dell'immagine può rallentare notevolmente la funzionalità " & _
                          "del foglio, preseguire?", vbYesNo)
        If MyAnswer = vbNo Then
            chkAutomaticUpdate = False
            GoTo MyExit
        End If
    End If
    
    If chkAutomaticUpdate Then
        chkAutomaticUpdate.BackColor = vbRed
        cmdUpdate.Enabled = False
    Else
        chkAutomaticUpdate.BackColor = &H8000000F
        cmdUpdate.Enabled = True
    
    End If

MyExit:
    Exit Sub
                         
End Sub

Public Sub cmdUpdate_Click()
    Dim MyTable As Object
    Dim MyTableHeaderRange As Range
    Dim MySheet As Excel.Worksheet
    Dim MyRow As Range
    Dim ImageFilePath As String
    Dim MyCell As Range
    Dim IsPivot As Boolean
    Dim IMAGE_PATH_FIELD As String
    Dim vf As Variant 'vf= visible field (per tabella pivot)
    
    On Error GoTo ErrHandler
    
    IMAGE_PATH_FIELD = "joint_sketch_file"
    IsPivot = False
    
    Set MySheet = ActiveSheet
    
    'Verifica se nel foglio c'è una tabella Pivot, nel caso farà riferimento a questa
    If MySheet.PivotTables.Count > 0 Then IsPivot = True
    
'GESTISCE IN MODO SEPARATO IL CASO DI TABELLA PIVOT O TABELLA
Select Case IsPivot
    'Caso di tabella (oggetto ListObject)
    Case False
        Set MyTable = MySheet.ListObjects(1)
        Set MyTableHeaderRange = MyTable.HeaderRowRange
            
        Set MyRow = MyTable.Range.Rows(ActiveCell.Row - MyTable.Range.Row + 1)
        
        ImageFilePath = ""
        For Each MyCell In MyTableHeaderRange.Cells
            If MyCell.Text = IMAGE_PATH_FIELD Then
                ImageFilePath = MyRow.Columns(MyCell.Column).Text
                'Se la cella "joint_sketch_file" contiene i ":" (due punti) significa che è indicato il percorso completo
                'e si tiene buono quello, altrimenti si aggiunge il path specificato nella cella "ImagePath"
                If InStr(1, ImageFilePath, ":", vbTextCompare) < 1 Then
                    ImageFilePath = MySheet.Range("ImagePath").Text & ImageFilePath
                End If
             GoTo Proceed
            End If
        Next
    
    'Caso di tabella pivot (oggetto PivotTable):
    Case True
        Set MyTable = MySheet.PivotTables(1)
        
        If Not PivotFieldIsVisible(MyTable, IMAGE_PATH_FIELD) Then
            MsgBox "Necessario visualizzare il campo '" & IMAGE_PATH_FIELD & "' per attivare la funzione!"
            Exit Sub
        End If
               
        Set MyTableHeaderRange = MyTable.PivotFields(IMAGE_PATH_FIELD).LabelRange
        Set MyRow = MyTable.PivotFields(IMAGE_PATH_FIELD).DataRange.Rows(ActiveCell.Row - MyTableHeaderRange.Row)
        ImageFilePath = MyRow.Cells(1, 1).Text
        'Se la cella "joint_sketch_file" contiene i ":" (due punti) significa che è indicato il percorso completo
        'e si tiene buono quello, altrimenti si aggiunge il path specificato nella cella "ImagePath"
        If InStr(1, ImageFilePath, ":", vbTextCompare) < 1 Then
            ImageFilePath = MySheet.Range("PivotImagePath").Text & ImageFilePath
        End If
        GoTo Proceed
        
End Select

Proceed:
    'MsgBox "ImageFilePath= " & ImageFilePath

    With Me.Image1
      .Picture = LoadPicture(ImageFilePath)
      .PictureSizeMode = fmPictureSizeModeZoom
    End With
    
    Me.BackColor = vbGreen
    
    UserFormUpdaterRow = MyRow.Row

MyExit:
    Exit Sub

ErrHandler:
    Select Case Err
        Case 53
            MsgBox "Nome di file non valido, verificare" & vbCrLf & _
            "Nome completo rilevato: " & ImageFilePath
        Case Else
            MsgBox "Ops, si è verificato un errore." & vbCrLf & "Error n." & Err.Number & ": " & Err.Description
    End Select
    Err.Clear
    Resume MyExit
    
End Sub

Private Sub UserForm_Activate()
    Call cmdUpdate_Click
End Sub

