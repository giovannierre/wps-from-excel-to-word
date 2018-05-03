Attribute VB_Name = "mAuxiliaryFunction"
Option Explicit

Public Function IsInArray(MyVal As Variant, arr As Variant) As Boolean

Dim i As Integer

On Error GoTo ErrHandler
   
    For i = LBound(arr) To UBound(arr)
        If MyVal = arr(i) Then
            IsInArray = True
            GoTo MyExit
        End If
    Next
    IsInArray = False

MyExit:
    Exit Function

ErrHandler:
    MsgBox "Errore nel calcolo della funzione 'IsInArray' (Err. n.)" & Err.Number & ": " & _
            Err.Description & ")"

End Function

Public Function IsInArray2(stringToBeFound As String, arr As Variant) As Boolean
'Funzione trovata qui:
'http://stackoverflow.com/questions/11109832/how-to-find-if-an-array-contains-a-string

    IsInArray2 = (UBound(Filter(arr, stringToBeFound)) > -1)

End Function

Public Function IsInCollection(col As Collection, ByVal key As String) As Boolean

On Error GoTo incol
col.Item key

incol:
IsInCollection = (Err.Number = 0)

End Function

Public Function SplitCellContent(CellContent As String, Delimiter As String, Pos As Integer) As String
   'Definisce una funzione di Split da usare nel foglio di lavoro
    
    Dim SplitArray As Variant
    
    'Se non c'è il delimitatore specificato allora restituisce
    'il valore della cella per la prima posizione e stringa vuota "" per le altre
    If InStr(1, CellContent, Delimiter) < 1 Then
        If Pos = 1 Then
            SplitCellContent = CellContent
        Else
            SplitCellContent = ""
        End If
    Else
        'Se trova il delimitatore allora fa lo split e restituisce la posizione specificata, finchè ce n'è,
        'dopo restituisce stringa vuoata ""
        SplitArray = Split(CellContent, Delimiter)
        If (Pos - 1) > UBound(SplitArray) Then
            SplitCellContent = ""
        Else
            SplitCellContent = RTrim(LTrim(SplitArray(Pos - 1)))
        End If
    End If
    
End Function
