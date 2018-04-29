Attribute VB_Name = "mCreateCustomProperties"

Sub CreateCustomProperties(TargetDocument As Variant, PropertyName As Collection, PropertyValue As Collection, _
                            Optional ShowPropertyList As Boolean = False)

'TargetDocument può essere un documento Word o Excel

Dim AddedProperty() As Variant
Dim Property As Variant
Dim PropertyList As String
Dim p As Variant
Dim i As Integer
Dim MyMsg
Dim ExistingProperties As Collection

On Error GoTo ErrHandler

ReDim AddedProperty(1 To 1) As Variant

'Crea una collection delle proprietà esistenti nel documento per decidere se scrivere
'solo il valore o se creare la nuova proprietà
Set ExistingProperties = New Collection
For Each p In TargetDocument.CustomDocumentProperties
    ExistingProperties.Add key:=p.Name, Item:=1
Next p

For Each p In PropertyName
 If IsInCollection(ExistingProperties, p) Then
    TargetDocument.CustomDocumentProperties(p).Value = PropertyValue(p)
  Else
    TargetDocument.CustomDocumentProperties.Add _
        Name:=p, _
        LinkToContent:=False, _
        Type:=msoPropertyTypeString, _
        Value:=PropertyValue(p)
    If ShowPropertyList Then
        AddedProperty(UBound(AddedProperty)) = p
        ReDim Preserve AddedProperty(1 To (UBound(AddedProperty) + 1))
    End If
 End If
 
Next p

If ShowPropertyList Then
    MyMsg = ""
    For Each Property In TargetDocument.CustomDocumentProperties
     PropertyList = PropertyList & vbCrLf & IIf(ValueInArray(Property.Name, AddedProperty), "*", " ") & Property.Name
    Next
    
    MyMsg = "Ecco le proprietà personalizzate del documento (contrassegnate con * quelle aggiunte):"
    
    MsgBox MyMsg & vbCrLf & PropertyList
End If

Exit Sub

ErrHandler:
    MsgBox "An error occured in sub 'CreateCustomeProperties': " & Err.Description & " (" & Err.Number & ")"

End Sub
Function PropertyExists(TargetDocument As Variant, ByVal PropertyName As String) As Boolean
 Dim p As Variant
 On Error Resume Next
 For Each p In TargetDocument.CustomDocumentProperties
    If p.Name = PropertyName Then
        PropertyExists = True
        Exit Function
    End If
 Next p
 PropertyExists = False
End Function

Public Function CustomProperty(ByVal prop As String)
    'Valido per Excel
    'Si può utilizzare in una cella per visualizzare il valore di una proprietà personalizzata
    'del documento
    On Error Resume Next
    CustomProperty = ActiveWorkbook.CustomDocumentProperties(prop)
End Function

Function ValueInArray(ByVal MyValue As Variant, ByRef MyArray As Variant) As Boolean
 Dim v As Variant
 On Error Resume Next
 For Each v In MyArray
    If v = MyValue Then
        ValueInArray = True
        Exit Function
    End If
 Next v
 ValueInArray = False
End Function

