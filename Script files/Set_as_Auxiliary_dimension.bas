'*********************************
'****** Copied and modified ******
'******** by Tomas Krenek ********
'*********************************

' Toto makro slouzi k uprave koty ve vykresu ( normalni vs. informativni )
' Toto makro je prevzate bez vetsich uprav

'write or delete (..) to drawing dimension
Sub CATMain()

' Nacteni zvolenych kot z catia
Dim MySelection As Selection
Set MySelection = CATIA.ActiveDocument.Selection

Dim MyDimension As DrawingDimension
Dim TextBefore As String
Dim TextAfter As String
Dim TextUpper As String
Dim TextLower As String

' Loop pro prochazeni zvolenych kot
For i = 1 To MySelection.Count
    
    ' Kontrola zda je zvolena kota
    If TypeName(MySelection.Item(i).Value) = "DrawingDimension" Then
    
        ' Nacteni objektu koty a vycteni hodnot
        Set MyDimension = MySelection.Item(i).Value
        MyDimension.GetValue.GetBaultText 1, TextBefore, TextAfter, TextUpper, TextLower
        
        ' Podminka pro zmenu mezi normalni a informativni kotou
        If (InStr(TextBefore, "(") <> 0 And InStr(TextAfter, ")") <> 0) Then
            MyDimension.GetValue.SetBaultText 1, "", "", TextUpper, TextLower
        Else
            MyDimension.GetValue.SetBaultText 1, "(", ")", TextUpper, TextLower
        End If
    Else
    End If
Next

End Sub
