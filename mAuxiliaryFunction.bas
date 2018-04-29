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
