'*********************************
'*** Developed by Tomas Krenek ***
'*********************************

' Toto makro slouzi k prejmenovani jednoho dilu ve strome Catie a ulozeni stare verze dilu do slozky kos

Option Explicit
 Dim CatiaSelection As Selection
 Dim SelectedDocument As Product
 
 Dim RenamingDocument As Object
 
 Dim fileSystem
 Dim DocumentFolderPath As String
 Dim DocumentFolder
 
 Dim documentList
 Dim MainProduct
 
 Dim TrashBinPath As String
 Dim TrashBinIncluded As Boolean
 
 Public OldFileName As String
 Public NewFileName As String
 Dim RenamedFileIncluded As Boolean
 
 Dim FilesCount As Integer
 
 Public IsRenamingInitializationOK As Boolean
 
 Dim i, j, k, l, m
 
 Dim ErrorMessage
 
Public Sub CATMain()
 
 ' nacteni dokument listu a vyberu
 Set documentList = CATIA.Documents
 
 Set CatiaSelection = CATIA.ActiveDocument.Selection
 
 ' kontrola ze je ve vyberu pouze jeden dokument ve strome
 If CatiaSelection.Count > 1 Then
    MsgBox "Je zvolen vice nez jeden dokument ve strome!", vbOKOnly, "Error"
    Exit Sub
 ElseIf CatiaSelection.Count < 1 Then
    MsgBox "Neni zvolen zadny dokument pro prejmenovani!", vbOKOnly, "Error"
    Exit Sub
 Else
 End If

 ' kontrola zda je zvoleny dokument v sestave
 If (InStr(CATIA.ActiveDocument.Name, ".CATProduct") <> 0) Then
    Set MainProduct = CATIA.ActiveDocument
 Else
    'MsgBox "Active document has to be .CATProduct"
    MsgBox "Dokument pro prejmenovani musi byt zvolen v .CATProduct", vbOKOnly, "Error"
    Exit Sub
 End If
 
 ' nastaveni zvoleneho dokumentu
 Set SelectedDocument = CatiaSelection.Item(1).LeafProduct
 
 ' Error detection
 ' On Error Resume Next
 
 ' kontrola vhodnosti dokumentu k prejmenovani
 If (InStr(SelectedDocument.ReferenceProduct.Parent.Name, ".CATProduct") <> 0 Or _
     InStr(SelectedDocument.ReferenceProduct.Parent.Name, ".CATPart") <> 0) Then
    
    ' hlaska pokud se neshoduje nazev souboru a nazev v catii
    If InStr(SelectedDocument.ReferenceProduct.Parent.Name, SelectedDocument.PartNumber) = 0 Then
        MsgBox "Upozorneni - nazev souboru a nazev v Catia se lisi!", vbOKOnly, "Error"
    Else
    End If
    
    ' nastaveni dokumentu pro prejmenovani
    Set RenamingDocument = SelectedDocument.ReferenceProduct.Parent
    
    ' extrakce cesty k souboru
    DocumentFolderPath = SelectedDocument.ReferenceProduct.Parent.Path
    
    ' ulozeni nazvu dilu pred prejmenovanim
    OldFileName = SelectedDocument.ReferenceProduct.Parent.Name
 Else
    'MsgBox "Active document has to be .CATProduct"
    MsgBox "! Tento vyber neni mozno prejmenovat !", vbOKOnly, "Error"
    Exit Sub
 End If
 
 ' detekce padu (neaktivni)
 If Err.Number <> 0 Then 'Any number other than zero is an error
    'Add your own custom error message to the user
    MsgBox "!Chyba nacteni dilu z Catia!", 16, "Error"
    Exit Sub
 End If
 On Error GoTo 0
 
 ' otevreni uzivatelskeho prostredi pro prejmenovani
 Call Document_renaming_Form.Show
 
 ' navratova hodnota z formulare
 If IsRenamingInitializationOK = False Then Exit Sub
 
 ' uprava vstupni hodnoty
 NewFileName = UCase(NewFileName)
 NewFileName = NewFileName + Right(OldFileName, Len(OldFileName) - InStrRev(OldFileName, ".") + 1)
 
 ' poravnani stary nazev vs. novy
 If NewFileName = OldFileName Then
    MsgBox "Puvodni a novy nazev jsou totozne opakujte!", vbOKOnly, "Error"
    Exit Sub
 Else
 End If
 
 ' nacteni file system
 Set fileSystem = CATIA.fileSystem
 
 ' kontrola kolize noveho nazvu dokumentu se existujicimi soubory
 If fileSystem.FileExists(DocumentFolderPath + "\" + NewFileName) Then
    MsgBox "Dil pod timto jmenem jiz existuje!", vbOKOnly, "Error"
    Exit Sub
 Else
 End If
 
 ' kontrola existence slozky pro ulozeni
 If DocumentFolderPath = "" Then
    MsgBox "Cesta k souboru nenalezena", vbOKOnly, "Error"
    Exit Sub
 ElseIf fileSystem.FolderExists(DocumentFolderPath) = False Then
    MsgBox "Cesta k souboru neexistuje", vbOKOnly, "Error"
    Exit Sub
 End If
 
 ' nacteni slozky
 Set DocumentFolder = fileSystem.GetFolder(DocumentFolderPath)
 
 ' deaktivovana kontrola padu
 'Error detection
 'On Error Resume Next
 
 'For i = 1 To DocumentFolder.SubFolders.Count
    'If DocumentFolder.SubFolders.Item(i).Name = "Kos" Then
        'TrashBinIncluded = True
        'Exit For
    'Else
        'TrashBinIncluded = False
    'End If
 'Next
 
 ' vytvoreni cesty ke slozce Kos
 TrashBinPath = DocumentFolderPath + "\Kos"
 
 ' kontrola zda slozka kos uz neexistuje, pripadne jeji zalozeni
 If fileSystem.FolderExists(TrashBinPath) = False Then
    fileSystem.CreateFolder (TrashBinPath)
    
    If fileSystem.FolderExists(TrashBinPath) Then
    Else
        MsgBox "Doslo k chybe pri vytvoreni kose!", vbOKOnly, "Error"
        Exit Sub
    End If
 End If
 
 ' ulozeni dokumentu pod novym nazvem a prejmenovani v catia
 RenamingDocument.SaveAs DocumentFolderPath + "\" + NewFileName
 RenamingDocument.Product.PartNumber = Left(NewFileName, InStrRev(NewFileName, ".") - 1)
 
 ' kontrola, ze dokument se ulozil spravne
 If fileSystem.FileExists(DocumentFolderPath + "\" + NewFileName) = False Then
    MsgBox "Nedoslo k vytvoreni dilu pod novym jmenem!", vbOKOnly, "Error"
    Exit Sub
 Else
 End If
 
 ' zaloha stareho dilu do slozky kos
 fileSystem.CopyFile DocumentFolderPath + "\" + OldFileName, TrashBinPath + "\" + OldFileName, True
 
 ' kontrola zalohy souboru
 If fileSystem.FileExists(TrashBinPath + "\" + OldFileName) = False Then
    MsgBox "Nedoslo k presunuti dilu do Kose!", vbOKOnly, "Error"
    Exit Sub
 Else
 End If
 
 ' vymazani dokumentu z puvodni slozky
 fileSystem.DeleteFile DocumentFolderPath + "\" + OldFileName
 
 'If Err.Number <> 0 Then 'Any number other than zero is an error
    'Add your own custom error message to the user
    'ErrorMessage = "!Chyba zpracovani souboru - opakujte!"
    'MsgBox ErrorMessage, 16, "Error"
 'End If
End Sub



