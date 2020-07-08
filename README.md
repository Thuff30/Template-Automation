# Template-Automation
  This is a Visual Basic script that can be adapted to automate the filling out of any document template. The script displayed was developed to speed up the completion of an IT service contract. It can be added to an existing document by selecting the Developer tab in Microsoft Word then Visual Basic. This will open a separate window where this script can be copied and pasted into and modified to fit.
  
  The document template that it was created for housed three different forms of the same contract, based on the service level of the contract. Because of the differences between the three service levels, the script is broken into multiple sections using a Select Case statement (this is similar in syntax and functionality to a switch statement in most other languages), with the exception of lines 60 – 137 that applied to all three versions.
  
  The script functions by asking key questions with test input boxes, then applying the user's input to replace key words or phrases throughout the document. This was accomplished in this case by using descriptive terms surrounded by < and > to denote a variable meant to be replaced such as <DATE> or <CLIENT>. See below for a sample of a block of code performing this action.
  
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
  
  By modifying the question being asked in the first line as well as the text to be replaced (in this case) it is possible to modify this block to find and replace any word or phrase needed. Simply repeat this for each word/phrase that will need to be filled in or replaced in the template to apply it to the needed document. If more questions are needed simply copy and paste the 17-line block of code used to find and replace and modify it as needed. An additional variable will need to be declared as well to hold the user input (see lines 4-21 for examples of variable declarations).
  
  Once all needed changes are made, the document will need to be saved as a macro enabled template. When this is done, each time the document is opened it will progress through the programmed series of questions and apply the user's input as it does to complete the template.
