'*********************************
'****** Copied and modified ******
'******** by Tomas Krenek ********
'*********************************

' Toto makro slozi k exportu vykresu Catia do pdf umistene ve stejne slozce

Option Explicit

Sub CATMain()
 Dim fileSys
 Dim folderPath
 Dim fileFolder
 Dim i As Integer
 Dim IFile
 Dim Doc
 Dim drawingName As String
 Dim sDocPath
 Dim PartDocument
 Dim pdfName

 CATIA.DisplayFileAlerts = True
 Set fileSys = CATIA.fileSystem

 'Folder path
 folderPath = InputBox("Vlozit cestu k souborum:", "Cesta k souboru s vykresy")

 If folderPath = "" Then
    MsgBox "Cesta k souboru nenalezena", vbOKOnly, "Error"
    Exit Sub
 End If

 Set fileFolder = fileSys.GetFolder(folderPath)
 'loop through all files in the folder
 For i = 1 To fileFolder.Files.Count
    Set IFile = fileFolder.Files.Item(i)

    'if the file is a CATDrawing, then open it in CATIA
    If InStr(IFile.Name, ".CATDrawing") <> 0 Then
        Set Doc = CATIA.Documents.Open(IFile.Path)
        Set PartDocument = CATIA.ActiveDocument

        'CATDrawing Update
        PartDocument.Update

        'Extracting drawing name
        drawingName = Len(CATIA.ActiveDocument.Name)
        pdfName = Left(CATIA.ActiveDocument.Name, drawingName - 11)

        'Export drawing (Alert for replacing document is off)
        CATIA.DisplayFileAlerts = False
        PartDocument.ExportData folderPath & "\" & pdfName, "pdf"
        CATIA.DisplayFileAlerts = True

        'save and close the open drawing document
        PartDocument.Save
        CATIA.ActiveDocument.Close
    End If
 Next 'go to the next drawing in the folder
 
 MsgBox "Akce probehla uspesne!", vbOKOnly, "Oznameni"
End Sub
