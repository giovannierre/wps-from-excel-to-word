VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Show Joint Sketch"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8460
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
    Dim SourceTable As Object
    Dim SourceTableHeaderRange As Range
    Dim MySheet As Excel.Worksheet
    Dim MyRow As Range
    Dim ImageFilePath As String
    Dim MyCell As Range
    Dim IsPivot As Boolean
    Dim IMAGE_PATH_FIELD As String
    Dim vf As Variant 'vf= visible field (per tabella pivot)
    Dim NAMED_CELL_FOR_IMAGE_PATH, _
        JOINT_NUMBER_FIELD, _
        WPS_NUMBER_FIELD, _
        JOINT_DETAILS_FIELD As String
    Dim SourceFields, FieldValues As Collection
    Dim f As Variant
    Dim MyItem As String
    Dim JointNumber, WPSNumber, JointDetails As String
    
    On Error GoTo ErrHandler
    
'***********************
'*******SETTINGS********
'***********************
    'Nome della cella che contiene il path delle immagini
    NAMED_CELL_FOR_IMAGE_PATH = "ImagePath"
    'Nomi dei campi della tabella che contengono le informazioni desiderate
    IMAGE_PATH_FIELD = "joint_sketch_file"
    JOINT_NUMBER_FIELD = "_Joint_No."
    WPS_NUMBER_FIELD = "wps_number"
    JOINT_DETAILS_FIELD = "joint_sketch_text_left"
    
    'Definisce una collection dei campi, sulla quale iterare per leggere i valori desiderati dalla tabella
    Set SourceFields = New Collection
    SourceFields.Add IMAGE_PATH_FIELD
    SourceFields.Add JOINT_NUMBER_FIELD
    SourceFields.Add WPS_NUMBER_FIELD
    SourceFields.Add JOINT_DETAILS_FIELD
    
'**************************
'*******END SETTINGS*******
'**************************

    'Questa collection servirà per memorizzare i valori dei campi letti dalla tabella
    Set FieldValues = New Collection
    
    'Setta a stringa vuota alcuni valori
    JointNumber = ""
    WPSNumber = ""
    JointDetails = ""
        
    'Di default considera che il foglio contenga una tabella (ListObject) e non una tabella pivot
    IsPivot = False
    
    Set MySheet = ActiveSheet
    
    'Verifica se nel foglio c'è una tabella Pivot, nel caso farà riferimento a questa
    If MySheet.PivotTables.Count > 0 Then IsPivot = True
    
'GESTISCE IN MODO SEPARATO IL CASO DI TABELLA PIVOT O TABELLA (ListObject)
Select Case IsPivot
    'Caso di tabella (oggetto ListObject)
    Case False
        Set SourceTable = MySheet.ListObjects(1)
        'Set SourceTableHeaderRange = SourceTable.HeaderRowRange
            
        Set MyRow = SourceTable.Range.Rows(ActiveCell.Row - SourceTable.Range.Row + 1)
        
        ImageFilePath = ""
                
        For Each f In SourceFields
                MyItem = MyRow.Columns(SourceTable.ListColumns(f).Range.Column).Cells(1, 1).Text
                If f = IMAGE_PATH_FIELD Then
                    'Se non trova un percorso completo nella cella, allora gli aggiunge il percorso
                    'generale specificato nel foglio
                    If InStr(1, MyItem, ":", vbTextCompare) < 1 Then
                        MyItem = Range(NAMED_CELL_FOR_IMAGE_PATH).Text & MyItem
                    End If
                End If
                FieldValues.Add key:=f, Item:=MyItem
        Next f

        ImageFilePath = FieldValues(IMAGE_PATH_FIELD)
        JointNumber = FieldValues(JOINT_NUMBER_FIELD)
        WPSNumber = FieldValues(WPS_NUMBER_FIELD)
        JointDetails = FieldValues(JOINT_DETAILS_FIELD)
        
        GoTo Proceed
                    
    'Caso di tabella pivot (oggetto PivotTable):
    Case True
        Set SourceTable = MySheet.PivotTables(1)
        
        If Not PivotFieldIsVisible(SourceTable, IMAGE_PATH_FIELD) Then
            MsgBox "Necessario visualizzare il campo '" & IMAGE_PATH_FIELD & "' per attivare la funzione!"
            Exit Sub
        End If
               
        Set SourceTableHeaderRange = SourceTable.PivotFields(IMAGE_PATH_FIELD).LabelRange
        Set MyRow = SourceTable.PivotFields(IMAGE_PATH_FIELD).DataRange.Rows(ActiveCell.Row - SourceTableHeaderRange.Row)
        ImageFilePath = MyRow.Cells(1, 1).Text
        'Se la cella "joint_sketch_file" contiene i ":" (due punti) significa che è indicato il percorso completo
        'e si tiene buono quello, altrimenti si aggiunge il path specificato nella cella "ImagePath"
        If InStr(1, ImageFilePath, ":", vbTextCompare) < 1 Then
            ImageFilePath = Range(NAMED_CELL_FOR_IMAGE_PATH).Text & ImageFilePath
        End If
        GoTo Proceed
        
End Select

Proceed:
    'MsgBox "ImageFilePath= " & ImageFilePath

    With Me.Image1
      .Picture = LoadPicture(ImageFilePath)
      .PictureSizeMode = fmPictureSizeModeZoom
    End With
    
    If JointNumber <> "" Then Me.lblJointNo = "Joint No: " & JointNumber
    If WPSNumber <> "" Then Me.lblWPSNo = "WPS No.: " & WPSNumber
    If JointDetails <> "" Then Me.lblJointDetails = JointDetails
    
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
    UserFormUpdaterRow = MyRow.Row
    Err.Clear
    Resume MyExit
    
End Sub

Private Sub UserForm_Activate()
    Call cmdUpdate_Click
End Sub

