'*********************************
'*** Developed by Tomas Krenek ***
'*********************************

' Toto makro slouzi k prejmenovani instanci name v Catia (pouziva dummy jmeno pro vyreseni kolizi)

Sub CATMain()

 Dim documentList
 Dim MainProduct As Product
 Dim MainDocument As Object

 Set documentList = CATIA.Documents
 Set MainDocument = CATIA.ActiveDocument
 Set TopProduct = MainDocument.Product

 If (InStr(MainDocument.Name, "CATProduct") <> 0) Then
    'Call RenameSingleLevelProduct(oTopProduct)
    Call RenameInstanceNames(TopProduct)
 Else
    MsgBox "Active document should be a Product"
    Exit Sub
 End If
 
 MsgBox "Akce probehla uspesne!", vbOKOnly, "Oznameni"
End Sub

'Instances renaming
Public Sub RenameInstanceNames(TopProduct)

 Dim ItemToRename As Product
 Dim ItemToRenamePartNumber As String
 Dim NumberOfProducts As Long
 Dim myArray(3000) As String
 Dim i, j, k As Integer
 
 'Dummy rename
 k = 0
 ItemToRenamePartNumber = "DummyName"
 
 NumberOfProducts = TopProduct.Products.Count
 
 For j = 1 To NumberOfProducts
    Set ItemToRename = TopProduct.Products.Item(j)
 
    'Rename as dummy instance
    k = k + 1
    ItemToRename.Name = ItemToRenamePartNumber & "." & k
    If (ItemToRename.Products.Count <> 0) Then
        Call RenameInstanceNames(ItemToRename.ReferenceProduct)
    End If
 Next
 
 'Right rename instance
 NumberOfProducts = TopProduct.Products.Count
 For i = 1 To NumberOfProducts
    myArray(i) = ""
 Next
 For i = 1 To NumberOfProducts
    Set ItemToRename = TopProduct.Products.Item(i)
    k = 0
    'Rename Instance
    ItemToRenamePartNumber = ItemToRename.PartNumber
    myArray(i) = ItemToRenamePartNumber
 
    For j = 1 To i
        If myArray(j) = ItemToRenamePartNumber Then
            k = k + 1
        End If
    Next

    ItemToRename.Name = ItemToRenamePartNumber & "." & k
    If (ItemToRename.Products.Count <> 0) Then
        Call RenameInstanceNames(ItemToRename.ReferenceProduct)
    End If
 Next
 
End Sub

'Old version with reorder fails
Private Sub RenameSingleLevelProduct(TopProduct)

 Dim ItemToRename As Product
 Dim ItemToRenamePartNumber As String
 Dim NumberOfItems As Long
 Dim myArray(500) As String
 Dim i, j, k As Integer

 NumberOfItems = TopProduct.Products.Count
 For i = 1 To NumberOfItems
    myArray(i) = ""
 Next
 
 For i = 1 To NumberOfItems
    Set ItemToRename = TopProduct.Products.Item(i)
    k = 0
    'Rename Instance
    ItemToRenamePartNumber = ItemToRename.PartNumber
    myArray(i) = ItemToRenamePartNumber
    
    For j = 1 To i
        If myArray(j) = ItemToRenamePartNumber Then
            k = k + 1
        End If
    Next
 
    ItemToRename.Name = ItemToRenamePartNumber & "." & k
    If (ItemToRename.Products.Count <> 0) Then
        Call RenameSingleLevelProduct(ItemToRename.ReferenceProduct)
    End If
 Next
 
End Sub
