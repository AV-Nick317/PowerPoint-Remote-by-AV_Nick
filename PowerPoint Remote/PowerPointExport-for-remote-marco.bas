Attribute VB_Name = "Module1"
Sub SaveSlidesAsJPG()
    Dim oSlide As Slide
    Dim sImagePath As String
    Dim sImageName As String

    ' Set the desired image path (e.g., "C:\Images\")
    sImagePath = "C:\Users\first\Desktop\FilesForPowerpointRemote\"

    ' Loop through each slide in the presentation
    For Each oSlide In ActivePresentation.Slides
        ' Construct the image file name (e.g., "1.jpg")
        sImageName = sImagePath & oSlide.SlideIndex & ".jpg"

        ' Export the slide as a JPG image
        oSlide.Export sImageName, "JPG"
    Next oSlide
    
    sImageName = Replace(sImageName, sImagePath, "")
    sImageName = Replace(sImageName, ".jpg", "")
    
    MsgBox (sImageName)
    
    strFileName = sImagePath & "0.txt"
    intFileNum = FreeFile()
    Open strFileName For Output As intFileNum
    Print #intFileNum, sImageName
    Close #intFileNum
    
End Sub







Sub ExportNotes()
' Write each slide's notes to a text file
' in same directory as presentation itself
' Each file is named NNNN_Notes_Slide_xxx
' where NNNN is the name of the presentation
'       xxx is the slide number

Dim oSl As Slide
Dim oSh As Shape
Dim strFileName As String
Dim strNotesText As String
Dim intFileNum As Integer

' Get the notes text
For Each oSl In ActivePresentation.Slides
    For Each oSh In oSl.NotesPage.Shapes
        If oSh.PlaceholderFormat.Type = ppPlaceholderBody Then
            If oSh.HasTextFrame Then
                If oSh.TextFrame.HasText Then
                    ' now write the text to file
                    strFileName = ActivePresentation.Path _
                        & "\FilesForPowerpointRemote\" & CStr(oSl.SlideIndex) _
                        & ".txt"
                    intFileNum = FreeFile()
                    Open strFileName For Append As intFileNum
                    Print #intFileNum, oSh.TextFrame.TextRange.Text
                    Close #intFileNum
                End If
            End If
        End If
    Next oSh
Next oSl

End Sub




Sub TransitionTimes()
    Dim oSlide As Slide
    Dim slideNumber As Integer
    Dim transitionTime As Double
    
    For Each oSlide In ActivePresentation.Slides
        slideNumber = oSlide.SlideIndex
        transitionTime = oSlide.SlideShowTransition.AdvanceTime
        Debug.Print "Slide " & slideNumber & ": Transition time = " & transitionTime & " seconds"
        
        strFileName = ActivePresentation.Path _
            & "\FilesForPowerpointRemote\" & CStr(oSlide.SlideIndex) _
            & ".txt"
        intFileNum = FreeFile()
        Open strFileName For Output As intFileNum
        Print #intFileNum, oSlide.SlideShowTransition.AdvanceTime
        Close #intFileNum
        
    Next oSlide
End Sub




Sub ExportforRemote()


'SaveSlidesAsJPG()
'TransitionTimes()
    Dim oSlide As Slide
    Dim slideNumber As Integer
    Dim sImagePath As String
    Dim sImageName As String
    Dim transitionTime As Double

    ' Set the desired image path (e.g., "C:\Images\")
    sImagePath = "C:\Users\first\source\repos\PowerPoint-Remote 3002\PowerPoint Remote\bin\Debug\net5.0-windows\win-x64\wwwroot\powerpoint_parts\"

    ' Loop through each slide in the presentation
    For Each oSlide In ActivePresentation.Slides
        'get slide number
        slideNumber = oSlide.SlideIndex
        'get transitionTime
        transitionTime = oSlide.SlideShowTransition.AdvanceTime
        'create txt file for transition time
        strFileName = sImagePath & oSlide.SlideIndex & ".txt"
        intFileNum = FreeFile()
        Open strFileName For Output As intFileNum
        Print #intFileNum, oSlide.SlideShowTransition.AdvanceTime
        Close #intFileNum
        ' Construct the image file name (e.g., "1.jpg")
        sImageName = sImagePath & oSlide.SlideIndex & ".jpg"

        ' Export the slide as a JPG image
        oSlide.Export sImageName, "JPG"
    Next oSlide
    
    
'ExportNotes()
    Dim oSl As Slide
    Dim oSh As Shape
    Dim strNotesText As String
    
    ' Get the notes text
    For Each oSl In ActivePresentation.Slides
        For Each oSh In oSl.NotesPage.Shapes
            If oSh.PlaceholderFormat.Type = ppPlaceholderBody Then
                If oSh.HasTextFrame Then
                    If oSh.TextFrame.HasText Then
                        ' now write the text to file
                        strFileName = sImagePath & oSl.SlideIndex & ".txt"
                        intFileNum = FreeFile()
                        Open strFileName For Append As intFileNum
                        Print #intFileNum, oSh.TextFrame.TextRange.Text
                        Close #intFileNum
                    End If
                End If
            End If
        Next oSh
    Next oSl
    
    
End Sub
