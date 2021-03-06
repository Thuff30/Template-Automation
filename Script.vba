Private Sub Document_Open()
    
    'Decalre all variables
    Dim contract As String
    Dim client As String
    Dim today As Date
    Dim servers As Integer
    Dim int_circuits As Integer
    Dim url As String
    Dim travel As Integer
    Dim basic_cost As Double
    Dim hrs_adv As Integer
    Dim adv_cost As Double
    Dim supp_users As Integer
    Dim eff_date As Date
    Dim locations As String
    Dim ult_fee As Double
    Dim myrange As Range
    Dim domains As String
    Dim myStoryRange As Range
    Dim rgePages As Range

    'Determine contract level
    contract = InputBox("Is the contract PPB, PPA, or PPU?", "Contract Selection", "<Contract>")
    
    'Select statement to remove unwanted contracts
    Select Case contract
        Case "PPB"
            'Delete PPA and PPU contracts
            Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=6
            Set rgePages = Selection.Range
            Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=21
            rgePages.End = Selection.Bookmarks("\Page").Range.End
            rgePages.Delete
            
        Case "PPA"
            'Delete PPB contact
            Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=1
            Set rgePages = Selection.Range
            Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=5
            rgePages.End = Selection.Bookmarks("\Page").Range.End
            rgePages.Delete
            
            'Delete PPU contact
            Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=13
            Set rgePages = Selection.Range
            Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=21
            rgePages.End = Selection.Bookmarks("\Page").Range.End
            rgePages.Delete
            
        Case "PPU"
            'Delete PPB and PPA contracts
            Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=1
            Set rgePages = Selection.Range
            Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=12
            rgePages.End = Selection.Bookmarks("\Page").Range.End
            rgePages.Delete
            
    End Select
    
    'Replace universal variables
    client = InputBox("Enter the client company name.", "Company Name", "<Client>")
        For Each myStoryRange In ActiveDocument.StoryRanges
            With myStoryRange.Find
                .Text = "<CLIENT>"
                .Replacement.Text = client
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
        Do While Not (myStoryRange.NextStoryRange Is Nothing)
            Set myStoryRange = myStoryRange.NextStoryRange
            With myStoryRange.Find
                .Text = "<CLIENT>"
                .Replacement.Text = client
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
        Loop
        Next myStoryRange
                
    today = InputBox("What is today's date?", "Date", "DD,MM,YYYY")
        For Each myStoryRange In ActiveDocument.StoryRanges
            With myStoryRange.Find
                .Text = "<DATE>"
                .Replacement.Text = today
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
        Do While Not (myStoryRange.NextStoryRange Is Nothing)
            Set myStoryRange = myStoryRange.NextStoryRange
            With myStoryRange.Find
                .Text = "<DATE>"
                .Replacement.Text = today
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
        Loop
        Next myStoryRange
        
    url = InputBox("What is the company website URL?", "URL", "<Web site URL>")
        For Each myStoryRange In ActiveDocument.StoryRanges
            With myStoryRange.Find
                .Text = "<Web site URL>"
                .Replacement.Text = url
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
            Do While Not (myStoryRange.NextStoryRange Is Nothing)
                Set myStoryRange = myStoryRange.NextStoryRange
                With myStoryRange.Find
                    .Text = "<Web site URL>"
                    .Replacement.Text = url
                    .Wrap = wdFindContinue
                    .Execute Replace:=wdReplaceAll
                End With
            Loop
            Next myStoryRange
                
    travel = InputBox("How many hours of travel will be billed for onsites?", "Travel Hours", "<Travel>")
        For Each myStoryRange In ActiveDocument.StoryRanges
            With myStoryRange.Find
                .Text = "<Travel>"
                .Replacement.Text = travel
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
            Do While Not (myStoryRange.NextStoryRange Is Nothing)
                Set myStoryRange = myStoryRange.NextStoryRange
                With myStoryRange.Find
                    .Text = "<Travel>"
                    .Replacement.Text = travel
                    .Wrap = wdFindContinue
                    .Execute Replace:=wdReplaceAll
                End With
            Loop
            Next myStoryRange
            
    'Select statement for contract specific variables
    Select Case contract
    
        'Peak Performance Basic contract variables
        Case "PPB"
            
            servers = InputBox("How many servers will be supported?", "Number of Servers", "<# Servers>")
                For Each myStoryRange In ActiveDocument.StoryRanges
                    With myStoryRange.Find
                        .Text = "<# Servers>"
                        .Replacement.Text = servers
                        .Wrap = wdFindContinue
                        .Execute Replace:=wdReplaceAll
                    End With
                    Do While Not (myStoryRange.NextStoryRange Is Nothing)
                        Set myStoryRange = myStoryRange.NextStoryRange
                        With myStoryRange.Find
                            .Text = "<# Servers>"
                            .Replacement.Text = servers
                            .Wrap = wdFindContinue
                            .Execute Replace:=wdReplaceAll
                        End With
                Loop
                Next myStoryRange
            
            int_circuits = InputBox("How many internet circuits will be supported?", "Number of Circuits", "<# Internet Access Circuits>")
                For Each myStoryRange In ActiveDocument.StoryRanges
                    With myStoryRange.Find
                        .Text = "<# Internet Access Circuits>"
                        .Replacement.Text = int_circuits
                        .Wrap = wdFindContinue
                        .Execute Replace:=wdReplaceAll
                    End With
                    Do While Not (myStoryRange.NextStoryRange Is Nothing)
                        Set myStoryRange = myStoryRange.NextStoryRange
                        With myStoryRange.Find
                            .Text = "<# Internet Access Circuits>"
                            .Replacement.Text = int_circuits
                            .Wrap = wdFindContinue
                            .Execute Replace:=wdReplaceAll
                        End With
                Loop
                Next myStoryRange
            
            basic_cost = InputBox("What is the total monthly rate?", "Basic Monthly Rate", "00.00")
                For Each myStoryRange In ActiveDocument.StoryRanges
                    With myStoryRange.Find
                        .Text = "<Basic Cost>"
                        .Replacement.Text = basic_cost
                        .Wrap = wdFindContinue
                        .Execute Replace:=wdReplaceAll
                    End With
                    Do While Not (myStoryRange.NextStoryRange Is Nothing)
                        Set myStoryRange = myStoryRange.NextStoryRange
                        With myStoryRange.Find
                            .Text = "<Basic Cost>"
                            .Replacement.Text = basic_cost
                            .Wrap = wdFindContinue
                            .Execute Replace:=wdReplaceAll
                        End With
                Loop
                Next myStoryRange
                
        'Peak Performance Advanced contract variables
        Case "PPA"
            
            servers = InputBox("How many servers will be supported?", "Number of Servers", "<# Servers>")
                For Each myStoryRange In ActiveDocument.StoryRanges
                    With myStoryRange.Find
                        .Text = "<# Servers>"
                        .Replacement.Text = servers
                        .Wrap = wdFindContinue
                        .Execute Replace:=wdReplaceAll
                    End With
                    Do While Not (myStoryRange.NextStoryRange Is Nothing)
                        Set myStoryRange = myStoryRange.NextStoryRange
                        With myStoryRange.Find
                            .Text = "<# Servers>"
                            .Replacement.Text = servers
                            .Wrap = wdFindContinue
                            .Execute Replace:=wdReplaceAll
                        End With
                Loop
                Next myStoryRange
            
            int_circuits = InputBox("How many internet circuits will be supported?", "Number of Circuits", "<# Internet Access Circuits>")
                For Each myStoryRange In ActiveDocument.StoryRanges
                    With myStoryRange.Find
                        .Text = "<# Internet Access Circuits>"
                        .Replacement.Text = int_circuits
                        .Wrap = wdFindContinue
                        .Execute Replace:=wdReplaceAll
                    End With
                    Do While Not (myStoryRange.NextStoryRange Is Nothing)
                        Set myStoryRange = myStoryRange.NextStoryRange
                        With myStoryRange.Find
                            .Text = "<# Internet Access Circuits>"
                            .Replacement.Text = int_circuits
                        .Wrap = wdFindContinue
                            .Execute Replace:=wdReplaceAll
                        End With
                Loop
                Next myStoryRange
    
            hrs_adv = InputBox("How many block hours will be purchased each month?", "Amount of Advances Block Hours", "<# Hrs-Advanced>")
                For Each myStoryRange In ActiveDocument.StoryRanges
                    With myStoryRange.Find
                        .Text = "<# Hrs-Advanced>"
                        .Replacement.Text = hrs_adv
                        .Wrap = wdFindContinue
                        .Execute Replace:=wdReplaceAll
                    End With
                    Do While Not (myStoryRange.NextStoryRange Is Nothing)
                        Set myStoryRange = myStoryRange.NextStoryRange
                        With myStoryRange.Find
                            .Text = "<# Hrs-Advanced>"
                            .Replacement.Text = hrs_adv
                            .Wrap = wdFindContinue
                            .Execute Replace:=wdReplaceAll
                        End With
                Loop
                Next myStoryRange
                
            adv_cost = InputBox("What is the total monthly rate?", "Monthly Rate", "00.00")
                For Each myStoryRange In ActiveDocument.StoryRanges
                    With myStoryRange.Find
                        .Text = "<Advanced Cost>"
                        .Replacement.Text = adv_cost
                        .Wrap = wdFindContinue
                        .Execute Replace:=wdReplaceAll
                    End With
                    Do While Not (myStoryRange.NextStoryRange Is Nothing)
                        Set myStoryRange = myStoryRange.NextStoryRange
                        With myStoryRange.Find
                            .Text = "<Advanced Cost>"
                            .Replacement.Text = adv_cost
                            .Wrap = wdFindContinue
                            .Execute Replace:=wdReplaceAll
                        End With
                Loop
                Next myStoryRange
                
        'Peak Performance Ultimate contract variables
        Case "PPU"
        
            supp_users = InputBox("How many users will be supported?", "Supported Users", "<# Supported Users>")
                For Each myStoryRange In ActiveDocument.StoryRanges
                    With myStoryRange.Find
                        .Text = "<# Supported Users>"
                        .Replacement.Text = supp_users
                        .Execute Replace:=wdReplaceAll
                    End With
                    Do While Not (myStoryRange.NextStoryRange Is Nothing)
                        Set myStoryRange = myStoryRange.NextStoryRange
                        With myStoryRange.Find
                            .Text = "<# Supported Users>"
                            .Replacement.Text = supp_users
                            .Wrap = wdFindContinue
                            .Execute Replace:=wdReplaceAll
                        End With
                    Loop
                    Next myStoryRange
                    
            eff_date = InputBox("What is the effective date for this contract?", "Effective Date", "<EFFECTIVE DATE>")
                For Each myStoryRange In ActiveDocument.StoryRanges
                    With myStoryRange.Find
                        .Text = "<EFFECTIVE DATE>"
                        .Replacement.Text = eff_date
                        .Execute Replace:=wdReplaceAll
                    End With
                    Do While Not (myStoryRange.NextStoryRange Is Nothing)
                        Set myStoryRange = myStoryRange.NextStoryRange
                        With myStoryRange.Find
                            .Text = "<EFFECTIVE DATE>"
                            .Replacement.Text = eff_date
                            .Wrap = wdFindContinue
                            .Execute Replace:=wdReplaceAll
                        End With
                    Loop
                    Next myStoryRange
            
            locations = InputBox("List the addresses of any other locations to be supported", "Alternate Locations", "<OTHER OFFICES>")
                For Each myStoryRange In ActiveDocument.StoryRanges
                    With myStoryRange.Find
                        .Text = "<OTHER OFFICES>"
                        .Replacement.Text = locations
                        .Execute Replace:=wdReplaceAll
                    End With
                    Do While Not (myStoryRange.NextStoryRange Is Nothing)
                        Set myStoryRange = myStoryRange.NextStoryRange
                        With myStoryRange.Find
                            .Text = "<OTHER OFFICES>"
                            .Replacement.Text = locations
                            .Wrap = wdFindContinue
                            .Execute Replace:=wdReplaceAll
                        End With
                    Loop
                    Next myStoryRange
                    
            domains = InputBox("What domains will be covered for registration, renewal, and DNS administration?", "Covered Domains", "<DOMAINS>")
                For Each myStoryRange In ActiveDocument.StoryRanges
                    With myStoryRange.Find
                        .Text = "<DOMAINS>"
                        .Replacement.Text = domains
                        .Wrap = wdFindContinue
                        .Execute Replace:=wdReplaceAll
                    End With
                    Do While Not (myStoryRange.NextStoryRange Is Nothing)
                        Set myStoryRange = myStoryRange.NextStoryRange
                        With myStoryRange.Find
                            .Text = "<DOMAINS>"
                            .Replacement.Text = domains
                            .Wrap = wdFindContinue
                            .Execute Replace:=wdReplaceAll
                        End With
                    Loop
                    Next myStoryRange
                    
            ult_fee = InputBox("What is the total monthly rate?", "Monthly Rate", "00.00")
                For Each myStoryRange In ActiveDocuments.StoryRanges
                    With myStoryRange.Find
                        .Text = "<Ultimate Fee>"
                        .Replacement.Text = ult_fee
                        .Wrap = wdFindContinue
                        .Execute Replace:=wdReplaceAll
                    End With
                    Do While Not (myStoryRange.NextStoryRange Is Nothing)
                        Set myStoryRange = myStoryRange.NextStoryRange
                        With myStoryRange.Find
                            .Text = "<Ultimate Fee>"
                            .Replacement.Text = ult_fee
                            .Wrap = wdFindContinue
                            .Execute Replace:=wdReplaceAll
                        End With
                    Loop
                    Next myStoryRange
                    
        End Select
            
End Sub


