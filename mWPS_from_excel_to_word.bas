Attribute VB_Name = "mWPS_from_excel_to_word"
Option Explicit

Sub read_wps_data()
Attribute read_wps_data.VB_ProcData.VB_Invoke_Func = "w\n14"

    Dim MySheet As Excel.Worksheet
    Dim MyTable As Excel.ListObject
    Dim MyCell, MyRow As Range
    Dim PropertyName As Collection
    Dim PropertyValue As Collection
    Dim TargetDocument As Object
    Dim TargetApplication As Object
    Dim TargetDocumentPath As String
    
    Set MySheet = ActiveWorkbook.Worksheets("WPS")
    
    Set MyTable = MySheet.ListObjects(1)
    
    ' Legge le intestazioni di colonna della tabella e le mette in un array
    Set PropertyName = New Collection
    For Each MyCell In MyTable.HeaderRowRange.Cells
     PropertyName.Add MyCell.Text
    Next
    
    Set PropertyValue = New Collection
    Set MyRow = MyTable.Range.Rows(ActiveCell.Row - MyTable.Range.Row + 1)
    
    For Each MyCell In MyRow.Cells
        PropertyValue.Add Item:=MyCell.Text, Key:=MyTable.HeaderRowRange.Columns(MyCell.Column).Text
    Next
        
    Set TargetApplication = CreateObject("Word.Application")
    
    'Seleziona il file di destinazione
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        .Show
        TargetDocumentPath = .SelectedItems(1)
    End With
    
    TargetApplication.documents.Open Filename:=TargetDocumentPath, ReadOnly:=False
    
    TargetApplication.Visible = True
    
    Set TargetDocument = TargetApplication.activedocument
    
    Call CreateCustomProperties(TargetDocument, PropertyName, PropertyValue)

End Sub
