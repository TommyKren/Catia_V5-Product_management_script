'*********************************
'*********** Developed ***********
'******** by Tomas Krenek ********
'*********************************

' Toto je vylepsena verze makra pro export vykresu do pdf
' Exportuje vsechny vykresy navazane na dily, ktere jsou soucasti hlavni sestavy

Option Explicit

Sub CATMain()
 Dim fileSystem
 Dim openWindows
 Dim activeWindow
 Dim documentList
 
 Dim folderPath
 Dim fileFolder
 
 Dim i As Integer
 
 Dim findingDrawing As String
 
 Dim drawingDocument
 Dim drawingName As String
 Dim pdfName As String
 
 Dim questionState
 
 Dim drawingState As Boolean
 
 ' kontrola zda je otevrena jen hlavni sestava (mozno doplnit o podminku s overenim)
 questionState = MsgBox("Mate otevrenu pouze hlavni sestavu?!", vbYesNo, "Varovani")
  
 If questionState <> 6 Then
    Exit Sub
 End If
 
 ' vlozeni cilove slozky pro ulozeni
 folderPath = InputBox("Vlozit cilovou slozku pro export:", "Cilova slozka pro vykresy")
 
 ' kontrola vstupu
 If folderPath = "" Then
    MsgBox "Cesta k souboru nezadana", vbOKOnly, "Error"
    Exit Sub
 End If

 ' nacteni otevrenych souboru v catia a file systemu
 Set documentList = CATIA.Documents
 Set fileSystem = CATIA.fileSystem
 
 ' vyhledani vykresu v souborech
 For i = 1 To documentList.Count
    ' nacteni kompletni cesty k dilu
    findingDrawing = documentList.Item(i).Path + "\" + documentList.Item(i).Name
    
    ' extrakce cesty bez koncovky
    findingDrawing = Left(findingDrawing, InStrRev(findingDrawing, "."))
    
    ' pridani koncovky CATDrawing
    findingDrawing = findingDrawing + "CATDrawing"
    
    ' kontrola existence souboru a jeho otevreni
    If fileSystem.FileExists(findingDrawing) Then
    
        documentList.Open findingDrawing
    
    End If
 
 Next
 
 ' nacteni otevrenych oken
 Set openWindows = CATIA.Windows
 
 ' loop pro prochazeni oken s jednotlivymi soubory
 Do
    
    drawingState = False
    
    ' loop pro vyhledavani okna s vykresem
    For i = 1 To openWindows.Count
    
        If InStr(openWindows.Item(i).Name, "CATDrawing") <> 0 Then
        
            openWindows.Item(i).Activate
            drawingState = True
            Exit For
            
        End If
        
    Next
    
    ' podminka pro ukonceni pri absenci nebo po dokonceni exportu vsech vykresu
    If drawingState = False Then
        
        MsgBox "Akce probehla uspesne!", vbOKOnly, "Oznameni"
        Exit Sub
        
    End If
    
    ' nacteni aktivniho okna
    Set activeWindow = CATIA.activeWindow
    
    ' nacteni dokumentu vykresu z aktivniho okna
    Set drawingDocument = activeWindow.Parent
    
    ' nacteni slozky pro export
    Set fileFolder = fileSystem.GetFolder(folderPath)
    
    ' update vykresu
    drawingDocument.Update

    ' zpracovani nazvu souboru pro export
    drawingName = Len(CATIA.ActiveDocument.Name)
    pdfName = Left(CATIA.ActiveDocument.Name, drawingName - 11)

    'Export drawing (Alert for replacing document is off)
    CATIA.DisplayFileAlerts = False
    drawingDocument.ExportData folderPath & "\" & pdfName, "pdf"
 
    'save and close the open drawing document
    drawingDocument.Save
    CATIA.ActiveDocument.Close
    CATIA.DisplayFileAlerts = True
        
 Loop Until drawingState = False
 
End Sub


