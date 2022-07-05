'*********************************
'*** Developed by Tomas Krenek ***
'*********************************

' Toto makro slouzi k prejmenovani nazvu dilu a jeho instance v Catii podle jmena souboru
' Makro vyuziva cast kodu z makra pro prejmenovani instanci

Sub CATMain()

 Dim documentList
 Dim MainProduct As Product
 Dim MainDocument As ProductDocument

 Set documentList = CATIA.Documents
 Set MainDocument = CATIA.ActiveDocument
 Set TopProduct = MainDocument.Product
 
 ' Overeni ze dokument pro prejmenovani je CATDocument
 If (InStr(MainDocument.Name, "CATProduct") <> 0) Then
    
    ' Volani funkce pro prejmenovani podle nazvu souboru
    Call RenamePartNameFromFileName(documentList)
    
    ' Volani funkce pro prejmenovani instanci
    Call Instance_name_Renaming.RenameInstanceNames(TopProduct)
 Else
    MsgBox "Aktivni dokument musi byt CATProduct"
    Exit Sub
 End If
 
 MsgBox "Akce probehla uspesne!", vbOKOnly, "Oznameni"
 
 End Sub

 Sub RenamePartNameFromFileName(documentList)
 
 'Define variables
 Dim ProductToRename As ProductDocument
 Dim PartToRename As PartDocument
 Dim FileName As String
 Dim NumberOfFiles As Long
 Dim myArray(500) As String
 Dim i, j As Integer
 Dim CatiaDocumentState As Boolean
 
 'Set variables
 CatiaDocumentState = False
 NumberOfFiles = documentList.Count
 
 'Cycle file after file
 For i = 1 To NumberOfFiles
 
 'Condition for CATProduct
 If (InStr(documentList.Item(i).Name, ".CATProduct") <> 0) Then
    CatiaDocumentState = True
 
    Set ProductToRename = documentList.Item(i)
 
    FileName = ProductToRename.Name
    FileName = Left(FileName, Len(FileName) - 11)
 
    If ProductToRename.Product.PartNumber <> FileName Then
        ProductToRename.Product.PartNumber = FileName
    Else
    End If
 
 Else
    'Condition for CATPart
    If (InStr(documentList.Item(i).Name, ".CATPart") <> 0) Then
    CatiaDocumentState = True
    
    Set PartToRename = documentList.Item(i)
    
    FileName = PartToRename.Name
    FileName = Left(FileName, Len(FileName) - 8)
    
    If PartToRename.Product.PartNumber <> FileName Then
        PartToRename.Product.PartNumber = FileName
    Else
    End If
 Else
    'Not Catia Document -> ignore
 End If
 End If
 Next

 End Sub

