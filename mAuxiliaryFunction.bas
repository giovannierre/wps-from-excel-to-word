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

Function IsFileOpen(FileName As String)
'Copiata pari pari da: https://support.microsoft.com/it-it/help/291295/macro-code-to-check-whether-a-file-is-already-open
    Dim filenum As Integer, errnum As Integer

    On Error Resume Next   ' Turn error checking off.
    filenum = FreeFile()   ' Get a free file number.
    ' Attempt to open the file and lock it.
    Open FileName For Input Lock Read As #filenum
    Close filenum          ' Close the file.
    errnum = Err           ' Save the error number that occurred.
    On Error GoTo 0        ' Turn error checking back on.

    ' Check to see which error occurred.
    Select Case errnum

        ' No error occurred.
        ' File is NOT already open by another user.
        Case 0
         IsFileOpen = False

        ' Error number for "Permission Denied."
        ' File is already opened by another user.
        Case 70
            IsFileOpen = True

        ' Another error occurred.
        Case Else
            Error errnum
    End Select

End Function

Function PivotFieldIsVisible(pt As Excel.PivotTable, FieldName As String) As Boolean
    
    On Error GoTo ErrHandler
    
    Debug.Print pt.PivotFields(FieldName).LabelRange.Address
    PivotFieldIsVisible = True
    Exit Function
    
ErrHandler:
    PivotFieldIsVisible = False
    Err.Clear
    
End Function

Public Function UpdateStringCollection(coll As Collection, ByVal MyKey As String, ByVal MyValue As String) As Collection
    coll.Remove MyKey
    coll.Add key:=MyKey, Item:=MyValue
    Set UpdateStringCollection = coll
End Function


Function DirExists(ByVal FullPath As String, Optional FullPathIsDir As Boolean = False, _
                   Optional CreateDir As Boolean = False, Optional MsgToUser As Boolean = False) As Boolean
'Controlla se una directory esiste e, se specificato nei parametri di ingresso, la crea
'Parametri:
'   FullPath:      percorso completo da verificare, può terminare con una directory o con un file
'                  Es.: "C:\la\mia\directory" oppure "C:\il\mio\file.pdf"
'   FullPathIsDir: definisce se il percorso è una directory o un file (vd. esempi sopra), se è un file la routine
'                  ne estrae la directory
'   CreateDir:     specifica se creare la directory, nel caso in cui la directory non esista
'   MsgToUser:     specifica se chiedere conferma all'utente per la creazione della directory, nel caso in cui non esista
    
    Dim Msg As String
    Dim DirSplit As Variant
    Dim TextToRemove As String
    
    On Error Resume Next
    
    'Se il FullPath non è una directory, allora toglie l'ultima parte del percorso e tiene solo la directory.
    'La funzione Split mi sembra la più rapida per elaborare la stringa: definisce tutti i pezzetti del percorso,
    'successivamente seleziono la stringa togliendo l'ultimo pezzetto.
    If Not FullPathIsDir Then
        DirSplit = Split(FullPath, "\")
        'La parte da rimuovere è l'ultima che segue lo slash "\", includendo lo slash
        TextToRemove = "\" & DirSplit(UBound(DirSplit))
        FullPath = Left(FullPath, Len(FullPath) - Len(TextToRemove))
    End If
    
    If Dir(FullPath, vbDirectory) = "" Then
        DirExists = False
        If CreateDir Then
            GoTo CreateDir
        Else
            If MsgToUser Then
                Msg = "La directory '" & FullPath & "'" & vbCrLf & "non esiste, vuoi crearla?"
                If (MsgBox(Msg, vbOKCancel)) = vbOK Then
                    GoTo CreateDir
                End If
            End If
        End If
    Else
        DirExists = True
    End If
Exit Function

CreateDir:
    MkDir Path:=FullPath
    'Controlla che la directory sia stata creata con successo
    If Not Dir(FullPath, vbDirectory) = "" Then
        DirExists = True
    Else
        If MsgToUser Then MsgBox "Ops! Per qualche motivo non è stato possibile creare la directory: " & vbCrLf & vbCrLf & _
                                  FullPath & vbCrLf & vbCrLf & _
                                 "Controllare che il percorso sia corretto."
        DirExists = False
    End If

End Function
