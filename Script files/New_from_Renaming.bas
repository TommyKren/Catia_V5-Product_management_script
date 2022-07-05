'*********************************
'*** Developed by Tomas Krenek ***
'*********************************

' Toto makro slouzi k prejmenovani dilu v sestave po New From a ulozeni vykresu pod novam nazvem

Option Explicit
 
 Dim documentList As Documents
 Dim MainDocument As Object
 Dim TopProduct As Object
 
 Dim ErrorMsg As String

'*************************************************************************
'Main function
Sub CATMain()
 Initialization
 
 MsgBox "Akce probehla uspesne!", vbOKOnly, "Oznameni"
End Sub

'*************************************************************************
'Initialization function for opening UserForm
Private Function Initialization() As Boolean
 Set documentList = CATIA.Documents
 Set MainDocument = CATIA.ActiveDocument

 If (InStr(MainDocument.Name, "CATProduct") <> 0) Then
    Initialization = True
    Set TopProduct = MainDocument.Product
    'MsgBox "Aktivni dokument je .CATProduct", vbOKOnly, "OK"
    New_from_renaming_Form.Show
 Else
    'MsgBox "Active document has to be .CATProduct"
    MsgBox "Aktivni dokument musi byt .CATProduct", vbOKOnly, "Error"
    Initialization = False
 End If

End Function

'*************************************************************************
'Initiation script Replacing text in Partname from input
Public Sub ReplaceTextInPartnameInitialization(TextOld, TextNew)
 Call ReplaceTextInPartname(documentList, TextOld, TextNew)
End Sub

'*************************************************************************
'Replacing text in Partname from input
Private Sub ReplaceTextInPartname(documentList, TextOld, TextNew)
 
 'Define variables
 Dim ItemToRename As Object
 Dim ItemName As String
 Dim ItemNameBeginning As String, ItemNameEnd As String
 Dim NumberOfItems As Long
 Dim SearchTextPosition As Integer
 
 Dim StateCatiaDocument As Boolean
 Dim StateCATDocument As Boolean
 Dim StateCATPart As Boolean
 
 Dim i, j As Integer
 
 NumberOfItems = documentList.Count

 'Cycle file after file
 For i = 1 To documentList.Count
 
    Set ItemToRename = documentList.Item(i)
 
    StateCatiaDocument = False
    'Condition for CATProduct
    If (InStr(ItemToRename.Name, ".CATProduct") <> 0) Then
        StateCatiaDocument = True
 
    Else
        'Condition for CATPart
        If (InStr(ItemToRename.Name, ".CATPart") <> 0) Then
            StateCatiaDocument = True
        Else
        End If
    End If
 
    If (StateCatiaDocument = True) Then
 
    SearchTextPosition = InStr(ItemToRename.Product.Name, TextOld)
 
        If (SearchTextPosition <> 0) Then
    
            ItemName = ItemToRename.Product.Name
            ItemNameBeginning = Left(ItemName, SearchTextPosition - 1)
            ItemNameEnd = Right(ItemName, Len(ItemName) - (SearchTextPosition + Len(TextOld)) + 1)
    
            ItemToRename.Product.PartNumber = ItemNameBeginning & TextNew & ItemNameEnd
        End If
    End If
 Next
 
End Sub

'*************************************************************************
'Initiation script Instances renaming
Public Sub RenameInstanceNamesInitialization()
 Call Instance_name_Renaming.RenameInstanceNames(TopProduct)
End Sub

'*************************************************************************
'Initiation script Export new from files
Public Sub SaveAsFilesInitialization()
 Call SaveAsFiles(documentList, TopProduct)
End Sub

'*************************************************************************
'Export new from files
Private Sub SaveAsFiles(documentList, TopProduct)

 Dim ItemToSave As Object
 Dim ProductName As String
 Dim NumberOfItems As Integer
 Dim i As Integer
 Dim folderPath As String
 Dim SavingPath As String
 Dim ItemToSaveName As String
 Dim SavingPathName As String
 Dim NumberOfViews As Integer
 Dim ViewObject As Object
 Dim SheetObject As Object
 
 'Target folder path
 'FolderPath = InputBox("Enter a folder path:", "Target folder for Save As")
 folderPath = InputBox("Vlozit cestu k cilove slozce:", "Slozka pro ulozeni")
 'FolderPath = "dummy"
 ProductName = TopProduct.Parent.Name
 
 'Saving main product
 SavingPathName = folderPath & "\" & ProductName
 CATIA.DisplayFileAlerts = False
 TopProduct.Parent.SaveAs (SavingPathName)
 CATIA.DisplayFileAlerts = True

 'Saving other catdrawings
 NumberOfItems = documentList.Count
 
 
 
 For i = 1 To NumberOfItems
 
 On Error Resume Next
 Set ItemToSave = documentList.Item(i)
 
 If (InStr(ItemToSave.Name, ".CATDrawing") <> 0) Then
    
    'Reset Number of items(Modification after debug error)
    NumberOfItems = documentList.Count

    ItemToSave.Update
    NumberOfViews = ItemToSave.Sheets.Item(1).Views.Count
    Dim j As Integer
        
    For j = 1 To NumberOfViews
        
        Dim CheckText As String
        CheckText = ItemToSave.Sheets.Item(1).Views.Item(j).Name
        
        'Condition for searching view with reference to part or assy (Front view or Iso view)
        If (ItemToSave.Sheets.Item(1).Views.Item(j).IsGenerative) Then
            ItemToSaveName = ItemToSave.Sheets.Item(1).Views.Item(j).GenerativeBehavior.Document.Parent.Name
        
            Dim IsScene As Boolean
            Dim IsPartBody As Boolean
    
            IsScene = (InStr(ItemToSave.Sheets.Item(1).Views.Item(j).GenerativeBehavior.Document.Parent.Name, "Scene") <> 0)
            IsPartBody = (InStr(ItemToSave.Sheets.Item(1).Views.Item(j).GenerativeBehavior.Document.Parent.Name, "Bodies") <> 0)
            
            'Debug.Print ItemToSave.Sheets.Item(1).Views.Item(j).GenerativeBehavior.Document.Parent.Name
            
            'Condition for view generated from scene
            If (IsScene Or IsPartBody) Then
                ItemToSaveName = ItemToSave.Sheets.Item(1).Views.Item(j).GenerativeBehavior.Document.Parent.Parent.Parent.Name
            Else
            End If
            
            'Condition delete end of name if view is generated from CATProduct
            If (InStr(ItemToSaveName, ".CATProduct") <> 0) Then
                ItemToSaveName = Left(ItemToSaveName, Len(ItemToSaveName) - 11)
                Exit For
            Else
                If (InStr(ItemToSaveName, ".CATPart") <> 0) Then
                    ItemToSaveName = Left(ItemToSaveName, Len(ItemToSaveName) - 8)
                    Exit For
                Else
                End If
            End If
        Else
        End If
    Next
    SavingPathName = folderPath & "\" & ItemToSaveName & ".CATDrawing"
    CATIA.DisplayFileAlerts = False
    ItemToSave.SaveAs (SavingPathName)
    CATIA.DisplayFileAlerts = True
    'ItemToSave.Close
 End If
 
 If Err.Number <> 0 Then 'Any number other than zero is an error
    'Add your own custom error message to the user
    ErrorMsg = "Chyba nacitani reference"
    MsgBox ErrorMsg, 16, "Error"
 End If
 Next
 
End Sub

'*************************************************************************
