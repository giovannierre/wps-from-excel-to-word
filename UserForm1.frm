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



Private Sub cmdUpdate_Click()
    Dim MyTable As Excel.ListObject
    Dim MyTableHeaderRange As Range
    Dim MyRow As Range
    Dim ImageFilePath As String
    Dim MyCell As Range
    
    
    Set MyTable = ActiveSheet.ListObjects(1)
    Set MyTableHeaderRange = MyTable.HeaderRowRange
    
    Set MyRow = MyTable.Range.Rows(ActiveCell.Row - MyTable.Range.Row + 1)
    
    ImageFilePath = ""
    For Each MyCell In MyTableHeaderRange.Cells
        If MyCell.Text = "joint_sketch_file" Then
         ImageFilePath = MyRow.Columns(MyCell.Column).Text
         GoTo Proceed
        End If
    Next

Proceed:
    'MsgBox "ImageFilePath= " & ImageFilePath

    With Me.Image1
      .Picture = LoadPicture(ImageFilePath)
      .PictureSizeMode = fmPictureSizeModeZoom
    End With
    
    Me.BackColor = vbGreen
    
    UserFormUpdaterRow = MyRow.Row

End Sub

Private Sub UserForm_Activate()
    Call cmdUpdate_Click
End Sub

Private Sub UserForm_Click()

End Sub
