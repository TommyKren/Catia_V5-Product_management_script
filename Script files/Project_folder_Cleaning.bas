Attribute VB_Name = "Project_folder_Cleaning"
'*********************************
'*** Developed by Tomas Krenek ***
'*********************************

' Toto makro slouzi k vycisteni slozky CAD od nevyuzitych souboru (presunuty do slozky Kos)

Option Explicit
 
 Dim fileSystem
 Dim FolderCleanPath As String
 Dim FolderClean
 
 Dim documentList
 Dim MainProduct
 
 Dim TrashBinIncluded As Boolean
 Dim TrashBinPath As String
 
 Dim CheckedFileName As String
 Dim ReferenceFileName As String
 Dim CheckedFileIncluded As Boolean
 
 Dim FilesCount As Integer
 
 Dim i, j, k, l, m
 
 Dim StartCheck
 
 Dim ErrorMessage
 
Public Sub CATMain()
 
 ' Overeni pred spustenim
 StartCheck = MsgBox("Opravdu chcete vycistit slozku s Catia dokumenty", vbYesNo, "Varovani")
 
 If StartCheck <> 6 Then Exit Sub
 
 ' Nacteni dokument listu
 Set documentList = CATIA.Documents
 
 ' Kontrola, ze je aktivni dokument CATProduct
 If (InStr(CATIA.ActiveDocument.Name, "CATProduct") <> 0) Then
    Set MainProduct = CATIA.ActiveDocument
 Else
    'MsgBox "Active document has to be .CATProduct"
    MsgBox "Aktivni dokument musi byt .CATProduct", vbOKOnly, "Error"
    Exit Sub
 End If
 
 ' Nacteni file systemu
 Set fileSystem = CATIA.fileSystem

 ' Extrakce cesty k souborum z CATProductu
 FolderCleanPath = MainProduct.Path

 If FolderCleanPath = "" Then
    MsgBox "Cesta k souboru nenalezena", vbOKOnly, "Error"
    Exit Sub
 End If
 
 ' Slozka pro vycisteni
 Set FolderClean = fileSystem.GetFolder(FolderCleanPath)
 
 'Error detection
 'On Error Resume Next
 
 ' kontrola pritomnosti slozky Kos
 For i = 1 To FolderClean.SubFolders.Count
    If FolderClean.SubFolders.Item(i).Name = "Kos" Then
        TrashBinIncluded = True
        Exit For
    Else
        TrashBinIncluded = False
    End If
 Next
 
 ' Cesta ke slozce Kos
 TrashBinPath = FolderCleanPath + "\Kos"
 
 ' Vytvoreni slozky kos pokud neexistuje
 If TrashBinIncluded = False Then
    fileSystem.CreateFolder (TrashBinPath)
    
    If fileSystem.FolderExists(TrashBinPath) Then
    Else
        MsgBox "Doslo k chybe pri vytvoreni kose", vbOKOnly, "Error"
        Exit Sub
    End If
 End If
 
  
 FilesCount = FolderClean.Files.Count
 
 ' Jeden po druhem projdu soubory ve slozce a porovnam s otevrenymi soubory
 j = 1
 While j <= FilesCount
    CheckedFileName = FolderClean.Files.Item(j).Name
    
    'Check Catia file
    If InStr(CheckedFileName, ".CATDrawing") <> 0 Or InStr(CheckedFileName, ".CATPart") <> 0 _
    Or InStr(CheckedFileName, ".CATProduct") <> 0 Or InStr(CheckedFileName, ".pdf") <> 0 Then
        'Comparing file to document list
        For k = 1 To documentList.Count
            ReferenceFileName = documentList.Item(k).Name
            ReferenceFileName = Left(ReferenceFileName, InStrRev(ReferenceFileName, ".") - 1)
            
            ' Podminka detekce dokumentu podle nazvu
            If InStr(CheckedFileName, ReferenceFileName) <> 0 Then
                CheckedFileIncluded = True
                Exit For
            Else
                CheckedFileIncluded = False
            End If
        Next
        
        ' Vycisteni prebytecneho souboru
        If CheckedFileIncluded = False Then
            fileSystem.CopyFile FolderCleanPath + "\" + CheckedFileName, TrashBinPath + "\" + CheckedFileName, True
            
            If fileSystem.FileExists(TrashBinPath + "\" + CheckedFileName) Then
                fileSystem.DeleteFile FolderCleanPath + "\" + CheckedFileName
                
                ' Toto pripocitani resi problem s loopem po vymazani souboru (nedokonale, obcas je treba pustit makro opakovane)
                j = j - 2
            Else
                MsgBox "Doslo k chybe pri presunu prebytecneho souboru", vbOKOnly, "Error"
                Exit Sub
            End If
        Else
        End If
    Else
    End If
    
    ' Prechod na dalsi soubor
    j = j + 1
    FilesCount = FolderClean.Files.Count
 Wend
 
 ' Tato cast neni spustena (detekce padu)
 If Err.Number <> 0 Then 'Any number other than zero is an error
    'Add your own custom error message to the user
    ErrorMessage = "!Chyba pracovani souboru - opakujte!"
    MsgBox ErrorMessage, 16, "Error"
 End If
 
 MsgBox "Akce probehla uspesne!", vbOKOnly, "Oznameni"
End Sub

