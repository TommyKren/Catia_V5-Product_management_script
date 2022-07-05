'*********************************
'*** Developed by Tomas Krenek ***
'*********************************

' Toto makro slouzi k poloautomatickemu vyplnovani vykresu s posledni verzi predtisku NAF
' !Pozor! existuje starsi verze predtisku, ktera nefunguje korektne

Option Explicit
 
 Dim documentList As Documents
 Dim EditDocument As drawingDocument
 Dim TitleBlockView As DrawingView
 
 Public IsInitializationOK As Boolean
 
 Public TextProjectNumber As String
 Public TextProjectIndex As String
 Public TextPartNumberName As String
 Public TextPartIndex As String
 
 Public TextScale As String
 Public TextGeneralDimensions As String
 Public TextMaterial As String
 Public TextSurface As String
 Public TextHardness As String
 
 Public TextCreatedDate As String
 Public TextCreatedName As String
 
 Public TextChangedDate As String
 Public TextChangedName As String
 Public TextChangedDescription As String
 
 Public TextDrawingName As String
 
 Public ActiveUser As String
 Public ActualDate As String
 
 Public FromDrawingMaterial As String
 
 Public ArrayMaterial(0 To 16) As String
 Public ArraySurface(0 To 9) As String
 Public ArrayHardness(0 To 3) As String
 
 Public NewDrawingState As Boolean
 Public NewDrawingDummyBool As BoolParam
 
 
'*************************************************************************
'Main procedure
Sub CATMain()
 
 'Initialization User interface
 Initialization
 If IsInitializationOK = True Then
    'Set Textboxes in Title block
    ApplyToDrawing
 End If
 
 'End Main procedure
End Sub

'*************************************************************************
'Initialization function for opening UserForm
Public Function Initialization() As Boolean
 
 'Read documents from Catia
 Set documentList = CATIA.Documents
  
 'Arrays for ComboBoxes
 ArrayMaterial(1) = "EN AW 6060"
 ArrayMaterial(2) = "EN AW 6082"
 ArrayMaterial(3) = "EN AW 5083"
 ArrayMaterial(4) = "Al alloy"
 ArrayMaterial(5) = "Brass"
 ArrayMaterial(6) = "1.0036"
 ArrayMaterial(7) = "1.0038"
 ArrayMaterial(8) = "1.0060"
 ArrayMaterial(9) = "1.1191"
 ArrayMaterial(10) = "1.2210"
 ArrayMaterial(11) = "1.2516"
 ArrayMaterial(12) = "1.2842"
 ArrayMaterial(13) = "POM"
 ArrayMaterial(14) = "PET-G"
 ArrayMaterial(15) = "Tecadur GF30"
 ArrayMaterial(16) = "Tecapeek GF30"
 
 ArraySurface(1) = "Anodized-natural"
 ArraySurface(2) = "Anodized-harden"
 ArraySurface(3) = "Sand blasted"
 ArraySurface(4) = "Blackened"
 ArraySurface(5) = "Galvanized"
 ArraySurface(6) = "Chem. nickel 3-5um"
 ArraySurface(7) = "Chem. nickel 10um"
 ArraySurface(8) = "Varnish RAL 7035"
 ArraySurface(9) = "Varnish RAL 3020"
 
 ArrayHardness(1) = "40+-2 HRC"
 ArrayHardness(2) = "50+-2 HRC"
 ArrayHardness(3) = "60+-2 HRC"
 
 'Actual inputs (User, actual date)
 ActiveUser = Environ("USERNAME")
 ActualDate = CStr(Day(Date)) + "." + CStr(Month(Date)) + "." + CStr(Right((Year(Date)), 2))
 
 'If active document is drawing, read document object to macro
 If (InStr(documentList.Application.ActiveDocument.Name, "CATDrawing") <> 0) Then
    IsInitializationOK = True
    
    'Set active document
    Set EditDocument = CATIA.ActiveDocument
    
    'Read Drawing name
    TextDrawingName = EditDocument.Name
    TextDrawingName = Left(TextDrawingName, Len(TextDrawingName) - 11)
    
    'Read drawing from title block
    ReadParametersFromDrawing
    
    'Open User interface
    Drawing_filling_Form.Show
 Else
    IsInitializationOK = False
    
    'MsgBox "Active document has to be .CATdrawing"
    MsgBox "Aktivni dokument musi byt .CATdrawing", vbOKOnly, "Error"
 End If
 
End Function

'*************************************************************************
'Reading parameters from actual drawing title
Private Sub ReadParametersFromDrawing()
 'Define variables
 Dim i, j, k, l, m As Integer
 Dim CreateDrawingState As Boolean
 CreateDrawingState = True

 'Find View with title block text (Main view)
 For i = 1 To EditDocument.Sheets.Item(1).Views.Count
    If (EditDocument.Sheets.Item(1).Views.Item(i).Name = "Main View") Then
        Set TitleBlockView = EditDocument.Sheets.Item(1).Views.Item(i)
        Exit For
    End If
 Next
 
 'Find New drawing parameter
 For k = 1 To EditDocument.Parameters.RootParameterSet.DirectParameters.Count
    If (InStr(EditDocument.Parameters.RootParameterSet.DirectParameters.Item(k).Name, "NewDrawing") <> 0) Then
        NewDrawingState = EditDocument.Parameters.RootParameterSet.DirectParameters.Item(k).Value
        CreateDrawingState = False
        Exit For
    End If
 Next
 
 'Set New drawing parameter, if no parameter is there
 If (CreateDrawingState = True) Then
    Set NewDrawingDummyBool = EditDocument.Parameters.RootParameterSet.DirectParameters.CreateBoolean("NewDrawing", True)
    NewDrawingState = True
 End If
 
 'Read first scale from Front view
 Dim TextScaleDummy As String
 'Find Front view
 For m = 1 To EditDocument.Sheets.Item(1).Views.Count
    If (InStr(EditDocument.Sheets.Item(1).Views.Item(m).Name, "Front view") <> 0) Then
        If (EditDocument.Sheets.Item(1).Views.Item(m).Scale <= 1) Then
            TextScale = "1:" + CStr(1 / EditDocument.Sheets.Item(1).Views.Item(m).Scale)
        Else
            TextScale = CStr(EditDocument.Sheets.Item(1).Views.Item(m).Scale) + ":1"
        End If
        Exit For
    Else
        If (InStr(EditDocument.Sheets.Item(1).Views.Item(m).Name, "Isometric view") <> 0) Then
            If (EditDocument.Sheets.Item(1).Views.Item(m).Scale <= 1) Then
                TextScale = "1:" + CStr(1 / EditDocument.Sheets.Item(1).Views.Item(m).Scale)
            Else
                TextScale = CStr(EditDocument.Sheets.Item(1).Views.Item(m).Scale) + ":1"
            End If
        Else
        
        End If
    End If
 Next
 
 TextScale = TextScale + " ("
 
 'Read other scales from Views
 For l = 1 To EditDocument.Sheets.Item(1).Views.Count
    
    Dim IsGenerative As Boolean
    Dim IsFrontView As Boolean
    Dim IsReady As Boolean
    
    IsGenerative = EditDocument.Sheets.Item(1).Views.Item(l).IsGenerative
    IsFrontView = EditDocument.Sheets.Item(1).Views.Item(l).Name <> "Front view"
    IsReady = IsGenerative And IsFrontView
    
    If (IsReady) Then
        If (EditDocument.Sheets.Item(1).Views.Item(l).Scale <= 1) Then
            TextScaleDummy = "1:" + CStr(1 / EditDocument.Sheets.Item(1).Views.Item(l).Scale)
        Else
            TextScaleDummy = CStr(EditDocument.Sheets.Item(1).Views.Item(l).Scale) + ":1"
        End If
    End If
    
    If (InStr(TextScale, TextScaleDummy) = 0) Then
        TextScale = TextScale + TextScaleDummy + ","
        
    End If
    
 Next
 If Right(TextScale, 2) = " (" Then
    TextScale = Left(TextScale, Len(TextScale) - 2)
 Else
    TextScale = Left(TextScale, Len(TextScale) - 1) + ")"
 End If
 
 'Read TextBoxes from Title block to local variables
 For j = 1 To TitleBlockView.Texts.Count
    Select Case TitleBlockView.Texts.Item(j).Name
        'TextProjectNumber
        Case "Project_number"
            TextProjectNumber = TitleBlockView.Texts.Item(j).Text
        'TextProjectIndex
        Case "Project_index"
            TextProjectIndex = TitleBlockView.Texts.Item(j).Text
        'TextProjectPartNumberName
        Case "Part_name"
            TextPartNumberName = TitleBlockView.Texts.Item(j).Text
        'TextProjectPartIndex
        Case "Part_index"
            TextPartIndex = TitleBlockView.Texts.Item(j).Text
        'TextProjectPartIndex
        Case "Scale"
            'TextScale = TitleBlockView.Texts.Item(j).Text
        'TextGeneralDimensions
        Case "General_dimensions"
            TextGeneralDimensions = TitleBlockView.Texts.Item(j).Text
        'TextMaterial
        Case "Material"
            TextMaterial = TitleBlockView.Texts.Item(j).Text
        'TextSurface
        Case "Surface"
            TextSurface = TitleBlockView.Texts.Item(j).Text
        'TextHardness
        Case "Hardness"
            TextHardness = TitleBlockView.Texts.Item(j).Text
        'TextCreatedDate
        Case "Create_date"
            TextCreatedDate = TitleBlockView.Texts.Item(j).Text
        'TextCreatedName
        Case "Create_name"
            TextCreatedName = TitleBlockView.Texts.Item(j).Text
        'TextChangedDate
        Case "Changed_date"
            TextChangedDate = TitleBlockView.Texts.Item(j).Text
        'TextChangedName
        Case "Changed_name"
            TextChangedName = TitleBlockView.Texts.Item(j).Text
        'TextChangedDescription
        Case "Changed_description"
            TextChangedDescription = TitleBlockView.Texts.Item(j).Text
    End Select
 Next

End Sub

'*************************************************************************
'Applying changed informations to drawing title
Private Sub ApplyToDrawing()
 'Define variables
 Dim i, j, k, l As Integer
 Dim GenerativeDocument As String
  
 'Set New drawing parameter as false
 For i = 1 To EditDocument.Parameters.RootParameterSet.DirectParameters.Count
    If (InStr(EditDocument.Parameters.RootParameterSet.DirectParameters.Item(i).Name, "NewDrawing") <> 0) Then
        EditDocument.Parameters.RootParameterSet.DirectParameters.Item(i).Value = False
        Exit For
    End If
 Next
 
 'Apply to part description
 For k = 1 To EditDocument.Sheets.Item(1).Views.Count
    
    Dim ReferenceDocument As Object
    Dim IsGenerative As Boolean
    Dim IsScene As Boolean
    Dim IsPartBody As Boolean
    
    IsGenerative = EditDocument.Sheets.Item(1).Views.Item(k).IsGenerative
    
    If (IsGenerative) Then
        
        IsScene = (InStr(EditDocument.Sheets.Item(1).Views.Item(k).GenerativeBehavior.Document.Parent.Name, "Scene") <> 0)
        IsPartBody = (InStr(EditDocument.Sheets.Item(1).Views.Item(k).GenerativeBehavior.Document.Parent.Name, "Bodies") <> 0)
        
        If (IsScene Or IsPartBody) Then
            For l = 1 To documentList.Count
            
                'Debug.Print DocumentList.Item(l).Name
                'Debug.Print EditDocument.Sheets.Item(1).Views.Item(k).GenerativeBehavior.Document.Parent.Parent.Parent.Name
                'Debug.Print ""
                
                If InStr(documentList.Item(l).Name, EditDocument.Sheets.Item(1).Views.Item(k).GenerativeBehavior.Document.Parent.Parent.Parent.Name) <> 0 And InStr(documentList.Item(l).Name, ".CATDrawing") = 0 Then
                    
                    Set ReferenceDocument = documentList.Item(l)
                    'Debug.Print ReferenceDocument.Product.Revision
                    
                    ReferenceDocument.Product.Revision = TextGeneralDimensions
                    ReferenceDocument.Product.Definition = TextSurface
                    ReferenceDocument.Product.Nomenclature = TextMaterial
                    ReferenceDocument.Product.DescriptionRef = TextHardness
                    Exit For
                Else
                End If
            Next
            Exit For
        Else
        End If
        
        EditDocument.Sheets.Item(1).Views.Item(k).GenerativeBehavior.Document.Revision = TextGeneralDimensions
        EditDocument.Sheets.Item(1).Views.Item(k).GenerativeBehavior.Document.Definition = TextSurface
        EditDocument.Sheets.Item(1).Views.Item(k).GenerativeBehavior.Document.Nomenclature = TextMaterial
        EditDocument.Sheets.Item(1).Views.Item(k).GenerativeBehavior.Document.DescriptionRef = TextHardness
        
        Exit For
    End If
 Next
 
 If (IsGenerative = False) Then
    MsgBox "Vykres neni spojeny s dilem"
 End If
 
 'Put text from variables to Title block
 For j = 1 To TitleBlockView.Texts.Count
    Select Case TitleBlockView.Texts.Item(j).Name
        'TextProjectNumber
        Case "Project_number"
            TitleBlockView.Texts.Item(j).Text = TextProjectNumber
        'TextProjectIndex
        Case "Project_index"
            TitleBlockView.Texts.Item(j).Text = TextProjectIndex
        'TextProjectPartNumberName
        Case "Part_name"
            TitleBlockView.Texts.Item(j).Text = TextPartNumberName
        'TextProjectPartIndex
        Case "Part_index"
            TitleBlockView.Texts.Item(j).Text = TextPartIndex
        'TextProjectPartIndex
        Case "Scale"
            TitleBlockView.Texts.Item(j).Text = TextScale
        'TextGeneralDimensions
        Case "General_dimensions"
            TitleBlockView.Texts.Item(j).Text = TextGeneralDimensions
        'TextMaterial
        Case "Material"
            TitleBlockView.Texts.Item(j).Text = TextMaterial
        'TextSurface
        Case "Surface"
            TitleBlockView.Texts.Item(j).Text = TextSurface
        'TextHardness
        Case "Hardness"
            TitleBlockView.Texts.Item(j).Text = TextHardness
        'TextCreatedDate
        Case "Create_date"
            TitleBlockView.Texts.Item(j).Text = TextCreatedDate
        'TextCreatedName
        Case "Create_name"
            TitleBlockView.Texts.Item(j).Text = TextCreatedName
        'TextChangedDate
        Case "Changed_date"
            TitleBlockView.Texts.Item(j).Text = TextChangedDate
        'TextChangedName
        Case "Changed_name"
            TitleBlockView.Texts.Item(j).Text = TextChangedName
        'TextChangedDescription
        Case "Changed_description"
            TitleBlockView.Texts.Item(j).Text = TextChangedDescription
        Case "Mass"
            TitleBlockView.Texts.Item(j).Text = ""
    End Select
 Next

End Sub

Public Function ExtractProductIndex(InputString As String) As String
 Dim dummy As String
 dummy = Trim(InputString)
 If dummy <> "" And IsNumeric(Left(dummy, 9)) Then
    dummy = Right(dummy, Len(dummy) - InStrRev(dummy, "-00"))
    ExtractProductIndex = dummy
 ElseIf InStr(dummy, "-00") = 0 Then
    ExtractProductIndex = "NONE"
 Else
    ExtractProductIndex = ""
 End If
End Function
