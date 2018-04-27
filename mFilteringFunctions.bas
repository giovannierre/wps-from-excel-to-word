Attribute VB_Name = "mFilteringFunctions"
'Option Explicit
Public FilterCache() As Variant
'GR2015-02-04: copiato pari pari da
'http://stackoverflow.com/questions/9489126/in-excel-vba-how-do-i-save-restore-a-user-defined-filter

'~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Sub:      SaveListObjectFilters
' Purpose:  Save filter on worksheet
' Returns:  wks.AutoFilterMode when function entered
' Source: http://stackoverflow.com/questions/9489126/in-excel-vba-how-do-i-save-        restore-a-user-defined-filter
'
' Arguments:
'   [Name]      [Type]  [Description]
'   wks         I/P     Worksheet that filter may reside on
'   FilterRange O/P     Range on which filter is applied as string; "" if no filter
'   FilterCache O/P     Variant dynamic array in which to save filter
'
' Author:   Based on MS Excel AutoFilter Object help file
'
' Modifications:
' 2006/12/11 Phil Spencer: Adapted as general purpose routine
' 2007/03/23 PJS: Now turns off .AutoFilterMode
' 2013/03/13 PJS: Initial mods for XL14, which has more operators
' 2013/05/31 P.H.: Changed to save list-object filters


Function SaveListObjectFilters(lo As ListObject, FilterCache()) As Boolean
Dim ii As Long

filterRange = ""
    With lo.AutoFilter
        filterRange = .Range.Address
        With .Filters
            ReDim FilterCache(1 To .Count, 1 To 3)
            For ii = 1 To .Count
                With .Item(ii)
                    If .On Then
#If False Then ' XL11 code
                        FilterCache(ii, 1) = .Criteria1
                        If .Operator Then
                            FilterCache(ii, 2) = .Operator
                            FilterCache(ii, 3) = .Criteria2
                        End If
#Else   ' first pass XL14
                        Select Case .Operator

                        Case 1, 2   'xlAnd, xlOr
                            FilterCache(ii, 1) = .Criteria1
                            FilterCache(ii, 2) = .Operator
                            FilterCache(ii, 3) = .Criteria2

                        Case 0, 3 To 7 ' no operator, xlTop10Items, _
xlBottom10Items, xlTop10Percent, xlBottom10Percent, xlFilterValues
                            FilterCache(ii, 1) = .Criteria1
                            FilterCache(ii, 2) = .Operator

                        Case Else    ' These are not correctly restored; there's someting in Criteria1 but can't save it.
                            FilterCache(ii, 2) = .Operator
                            ' FilterCache(ii, 1) = .Criteria1   ' <-- Generates an error
                            ' No error in next statement, but couldn't do restore operation
                            ' Set FilterCache(ii, 1) = .Criteria1

                        End Select
#End If
                    End If
                End With ' .Item(ii)
            Next
        End With ' .Filters
    End With ' wks.AutoFilter
End Function


'~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Sub:      RestoreListObjectFilters
' Purpose:  Restore filter on listobject
' Source: http://stackoverflow.com/questions/9489126/in-excel-vba-how-do-i-save-restore-a-user-defined-filter
' Arguments:
'   [Name]      [Type]  [Description]
'   wks         I/P     Worksheet that filter resides on
'   FilterRange I/P     Range on which filter is applied
'   FilterCache I/P     Variant dynamic array containing saved filter
'
' Author:   Based on MS Excel AutoFilter Object help file
'
' Modifications:
' 2006/12/11 Phil Spencer: Adapted as general purpose routine
' 2013/03/13 PJS: Initial mods for XL14, which has more operators
' 2013/05/31 P.H.: Changed to restore list-object filters
'
' Comments:
'----------------------------
Sub RestoreListObjectFilters(lo As ListObject, FilterCache())
Dim col As Long

If lo.Range.Address <> "" Then
    For col = 1 To UBound(FilterCache(), 1)

#If False Then  ' XL11
        If Not IsEmpty(FilterCache(col, 1)) Then
            If FilterCache(col, 2) Then
                lo.AutoFilter Field:=col, _
                    Criteria1:=FilterCache(col, 1), _
                        Operator:=FilterCache(col, 2), _
                    Criteria2:=FilterCache(col, 3)
            Else
                lo.AutoFilter Field:=col, _
                    Criteria1:=FilterCache(col, 1)
            End If
        End If
#Else

        If Not IsEmpty(FilterCache(col, 2)) Then
            Select Case FilterCache(col, 2)

            Case 0  ' no operator
                lo.Range.AutoFilter Field:=col, _
                    Criteria1:=FilterCache(col, 1) ' Do NOT reload 'Operator'

            Case 1, 2   'xlAnd, xlOr
                lo.Range.AutoFilter Field:=col, _
                    Criteria1:=FilterCache(col, 1), _
                    Operator:=FilterCache(col, 2), _
                    Criteria2:=FilterCache(col, 3)

            Case 3 To 6 ' xlTop10Items, xlBottom10Items, xlTop10Percent,     xlBottom10Percent
#If True Then
                lo.Range.AutoFilter Field:=col, _
                    Criteria1:=FilterCache(col, 1) ' Do NOT reload 'Operator' , it doesn't work
                ' wks.AutoFilter.Filters.Item(col).Operator = FilterCache(col, 2)
#Else ' Trying to restore Operator as well as Criteria ..
                ' Including the 'Operator:=' arguement leads to error.
                ' Criteria1 is expressed as if for a FALSE .Operator
                lo.Range.AutoFilter Field:=col, _
                    Criteria1:=FilterCache(col, 1), _
                    Operator:=FilterCache(col, 2)
#End If

            Case 7  'xlFilterValues
                lo.Range.AutoFilter Field:=col, _
                    Criteria1:=FilterCache(col, 1), _
                    Operator:=FilterCache(col, 2)

#If False Then ' Switch on filters on cell formats
' These statements restore the filter, but cannot reset the pass Criteria, so the filter hides all data.
' Leave it off instead.
            Case Else   ' (Various filters on data format)
                lo.RangeAutoFilter Field:=col, _
                    Operator:=FilterCache(col, 2)
#End If ' Switch on filters on cell formats

            End Select
        End If

#End If     ' XL11 / XL14
    Next col
End If
End Sub

Sub SaveFilter()
 Call SaveListObjectFilters(ActiveSheet.ListObjects(1), FilterCache())
End Sub

Sub RestoreFilter()
 Call RestoreListObjectFilters(ActiveSheet.ListObjects(1), FilterCache())
End Sub

Sub CancellaFiltri_Click()
'
' CancellaFiltri_Click Macro

    On Error Resume Next
    
    ActiveSheet.ListObjects(1).Range(1, 1).Select
    
    ActiveSheet.ShowAllData

End Sub

Sub FilterField(ByVal FieldNumber As Integer, ByVal Criteria As String)
    
    On Error Resume Next
    
    'ActiveSheet.ListObjects(1).Range(1, 1).Select
    
    Criteria = "=*" & Criteria & "*"
    
    ActiveSheet.ListObjects(1).Range.AutoFilter Field:=FieldNumber, Criteria1 _
        :=Criteria, Operator:=xlAnd

End Sub
