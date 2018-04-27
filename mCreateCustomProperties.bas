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

'On Error Resume Next

ReDim AddedProperty(1 To 1) As Variant

For Each p In PropertyName
 
 If Not PropertyExists(TargetDocument, p) Then
    TargetDocument.CustomDocumentProperties.Add _
        Name:=p, _
        LinkToContent:=False, _
        Type:=msoPropertyTypeString, _
        Value:=PropertyValue(p)
    AddedProperty(UBound(AddedProperty)) = p
    ReDim Preserve AddedProperty(1 To (UBound(AddedProperty) + 1))
  Else
    TargetDocument.CustomDocumentProperties(p).Value = PropertyValue(p)
 End If
 
Next p

If ShowPropertyList = True Then
    MyMsg = ""
    For Each Property In TargetDocument.CustomDocumentProperties
     PropertyList = PropertyList & vbCrLf & IIf(ValueInArray(Property.Name, AddedProperty), "*", " ") & Property.Name
    Next
    
    MyMsg = "Ecco le proprietà personalizzate del documento (contrassegnate con * quelle aggiunte):"
    
    MsgBox MyMsg & vbCrLf & PropertyList
End If

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

