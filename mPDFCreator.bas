Attribute VB_Name = "mPDFCreator"
Option Explicit

'Questo modulo è un tentativo di utilizzare il modello a oggetti di PDFCreator per
'gestire le stampe dei documenti nel modo più flessibile possibile
'NB: come prima cosa bisogna selezionare tra i riferimenti 'PDFCreator_COM' (Component Object Model)

Sub PDFCreatorPrint()
    Dim oPDFCreator As PdfCreatorObj
    Dim oQueue As Queue
    Dim oPrintJob As PrintJob
    Dim FullPath As String
        
    Set oPDFCreator = New PdfCreatorObj
    Set oQueue = New Queue
    
    'Se c'è una istanza attiva la elimina, per poter ricominciare da capo
    If oPDFCreator.IsInstanceRunning Then oQueue.ReleaseCom
        
    'Inizializza la coda di stampa
    oQueue.Initialize
    
    'Application.ActivePrinter = "PDFCreator" 'Per qualche motivo non funziona, da verificare
    ActiveSheet.PrintOut 1, 1
       
    FullPath = ActiveWorkbook.Path & "\ProvePDFCreator\" & "TestPage.pdf"
    
    If Not oQueue.WaitForJob(10) Then
        MsgBox "The print job did not reach the queue within " & " 10 seconds”"
    Else
        MsgBox "Currently there are " & oQueue.Count & " job(s) in the queue”"
    End If
    
        
    'Setta un oggetto PrintJob
    Set oPrintJob = oQueue.NextJob
    
    'Definisce alcuni setting del PrintJob
    oPrintJob.SetProfileByGuid ("DefaultGuid")
    oPrintJob.SetProfileSetting "AuthorTemplate", "GR"
    
    oPrintJob.ConvertTo (FullPath)
    
    If (Not oPrintJob.IsFinished Or Not oPrintJob.IsSuccessful) Then
        MsgBox "Ops! Some problem occured, could not convert the file: " & FullPath
    Else
        MsgBox "Job finished successfully"
    End If
    
    oQueue.ReleaseCom

End Sub
