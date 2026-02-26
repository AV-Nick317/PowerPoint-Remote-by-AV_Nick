Attribute VB_Name = "Module1"
Function EncodeToHTML2(sInput As String) As String

    sInput = Replace(sInput, vbTab, "&nbsp;&nbsp;&nbsp;&nbsp;")
    sInput = Replace(sInput, vbCr, "<br>")
    sInput = Replace(sInput, vbLf, "<br>")
    sInput = Replace(sInput, vbCrLf, "<br>")
    sInput = Replace(sInput, Chr(160), "&nbsp;")
    'sInput = Replace(sInput, "<BR>&nbsp;", "<BR>&nbsp;&nbsp;&nbsp;&nbsp;")
    sInput = Replace(sInput, "<BR>", "<br>")
    sInput = Replace(sInput, "‘", "&lsquo;")
    sInput = Replace(sInput, "’", "&rsquo;")
    sInput = Replace(sInput, "“", "&ldquo;")
    sInput = Replace(sInput, "”", "&rdquo;")
    
    EncodeToHTML2 = sInput
    
End Function


Sub RemoveAudioFromAllSlides()
    Dim oSl As slide
    Dim oSh As Shape
    Dim x As Long
    
    For Each oSl In ActivePresentation.Slides
        For x = oSl.Shapes.Count To 1 Step -1
            With oSl.Shapes(x)
                If .Type = msoMedia Then
                    If .MediaType = ppMediaTypeSound Then
                        .Delete
                    End If
                End If
            End With
        Next     ' x
    Next    ' Slide

End Sub

Sub saveAsS()

Dim opres As Presentation
Set opres = ActivePresentation
ActivePresentation.SaveCopyAs Environ("USERPROFILE") & "\Desktop\s.pptx"

End Sub

Sub saveAsL()

Dim opres As Presentation
Set opres = ActivePresentation
ActivePresentation.SaveCopyAs Environ("USERPROFILE") & "\Desktop\l.pptx"

End Sub


Sub ServiceExport()

'changeBGtoBlack()
ActivePresentation.SlideMaster.Background.Fill.ForeColor.RGB = RGB(0, 0, 0)

'ReplaceLetTheVineyardsFLC()

Dim slide As slide
Dim notesText As String
Dim slideFound As Boolean

slideFound = False

For Each slide In ActivePresentation.Slides
    notesText = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
    If InStr(1, notesText, "Let_the_Vineyards_be_Fruitful-FLC", vbTextCompare) > 0 Then
        slide.SlideShowTransition.AdvanceOnTime = msoTrue
        slide.SlideShowTransition.AdvanceTime = 0
        slideFound = True
        Exit For
    End If
Next slide

If slideFound < 0 Then
    
    Dim slideNum As Integer
    Dim slideNum2 As Integer
    
    slideNum = slide.slideIndex
    
    Dim CountSlidesNum As Integer
    Dim oPresentation As Presentation
    Set oPresentation = Presentations.Open("C:\Users\first\Documents\Presentation_Local\Let_the_Vineyards_be_Fruitful-FLC.pptx")
    
    slideNum2 = slideNum - oPresentation.Slides.Count + 1
    oPresentation.Close
    
    ActivePresentation.Slides.InsertFromFile _
    "C:\Users\first\Documents\Presentation_Local\Let_the_Vineyards_be_Fruitful-FLC.pptx", slideNum
    
    ActivePresentation.Slides.Range(MyRange(slideNum2, slideNum)).Delete
    
End If

'saveAsS

Dim opres As Presentation
Set opres = ActivePresentation

ActivePresentation.SaveCopyAs Environ("USERPROFILE") & "\Creative Cloud Files\Desktop\s.pptx"

'ExportforRemote()

    Dim j As Integer
    Dim baseName2 As String
    Dim charCode As Integer

    Randomize
    For j = 1 To 3
        ' Generate a random number between 97 ("a") and 122 ("z")
        charCode = Int((26 * Rnd) + 97)
        baseName2 = baseName2 & Chr(charCode)
    Next j



'SaveSlidesAsJPG()
'TransitionTimes()
    Dim oSlide As slide
    Dim slideNumber As Integer
    Dim sImagePath As String
    Dim sImagePath2 As String
    Dim sImageName As String
    Dim transitionTime As Double

    ' Set the desired image path (e.g., "C:\Images\")
    sImagePath = "C:\Users\first\source\repos\PowerPoint-Remote 3006\PowerPoint Remote\bin\Debug\net5.0-windows\win-x64\wwwroot\powerpoint_parts\"
    
    Kill sImagePath & "*.*"

    ' Loop through each slide in the presentation
    For Each oSlide In ActivePresentation.Slides
        'get slide number
        slideNumber = oSlide.slideIndex
        'get transitionTime
        transitionTime = oSlide.SlideShowTransition.AdvanceTime
        'create txt file for transition time
        strFileName = sImagePath & baseName2 & oSlide.slideIndex & ".txt"
        intFileNum = FreeFile()
        Open strFileName For Output As intFileNum
        Print #intFileNum, oSlide.SlideShowTransition.AdvanceTime
        Close #intFileNum
        ' Construct the image file name (e.g., "1.jpg")
        sImageName = sImagePath & baseName2 & oSlide.slideIndex & ".jpg"

        ' Export the slide as a JPG image
        oSlide.Export sImageName, "JPG"
    Next oSlide
    
'//create slide count number txt file
    sImagePath2 = sImagePath & baseName2
    sImageName = Replace(sImageName, sImagePath2, "")
    sImageName = Replace(sImageName, ".jpg", "")
    
    strFileName = sImagePath & "0.txt"
    intFileNum = FreeFile()
    Open strFileName For Output As intFileNum
    Print #intFileNum, sImageName
    Print #intFileNum, baseName2
    Close #intFileNum
    
    
'ExportNotes()
    Dim oSl As slide
    Dim oSh As Shape
    Dim strNotesText As String
    
    ' Get the notes text
    For Each oSl In ActivePresentation.Slides
        For Each oSh In oSl.NotesPage.Shapes
            If oSh.PlaceholderFormat.Type = ppPlaceholderBody Then
                If oSh.HasTextFrame Then
                    If oSh.TextFrame.HasText Then
                        ' now write the text to file
                        strFileName = sImagePath & baseName2 & oSl.slideIndex & ".txt"
                        intFileNum = FreeFile()
                        Open strFileName For Append As intFileNum
                        Print #intFileNum, EncodeToHTML2(oSh.TextFrame.TextRange.Text)
                        Close #intFileNum
                    End If
                End If
            End If
        Next oSh
    Next oSl



'RemoveAudioFromAllSlides

    Dim x As Long
    
    For Each oSl In ActivePresentation.Slides
        For x = oSl.Shapes.Count To 1 Step -1
            With oSl.Shapes(x)
                If .Type = msoMedia Then
                    If .MediaType = ppMediaTypeSound Then
                        .Delete
                    End If
                End If
            End With
        Next     ' x
    Next    ' Slide
    

'changeBGtoGreen()
ActivePresentation.SlideMaster.Background.Fill.ForeColor.RGB = RGB(0, 255, 0)

'SendTextToBottom()
ActivePresentation.SlideMaster.CustomLayouts(3).Shapes(1).TextFrame.VerticalAnchor = msoAnchorBottom


'ReplaceLivestreamFiles()
    Dim folderPath As String
    Dim fileName As String
    Dim baseName As String
    Dim lastRow As Long
        
    ' Change this to the path of your folder
    folderPath = "C:\Users\first\Documents\Presentation_Local\Livestream_Replacements\"
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    fileName = Dir(folderPath & "*.*")
    lastRow = 1
    
    Do While fileName <> ""
        baseName = Left(fileName, InStrRev(fileName, ".") - 1)
        
        slideFound = False
        
        For Each slide In ActivePresentation.Slides
            notesText = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
            If InStr(1, notesText, baseName, vbTextCompare) > 0 Then
                slide.SlideShowTransition.AdvanceOnTime = msoTrue
                slide.SlideShowTransition.AdvanceTime = 0
                slideFound = True
                Exit For
            End If
        Next slide
        
        If slideFound < 0 Then
            
            slideNum = slide.slideIndex
            
            Set oPresentation = Presentations.Open(folderPath & fileName)
            
            slideNum2 = slideNum - oPresentation.Slides.Count + 1
            oPresentation.Close
            
            ActivePresentation.Slides.InsertFromFile _
            folderPath & fileName, slideNum
            
            ActivePresentation.Slides.Range(MyRange(slideNum2, slideNum)).Delete
            
        End If
        
        lastRow = lastRow + 1
        fileName = Dir
    Loop

    
'FindSlideWithFLCoffering()
    
    slideFound = False
    
    For Each slide In ActivePresentation.Slides
        For Each shp In slide.Shapes
            If shp.HasTextFrame Then
                If InStr(1, shp.TextFrame.TextRange.Text, "OFFERING:") > 0 Then
                    slideNum = slide.slideIndex
                    slideFound = True
                    Exit For ' Exit loop if text is found on the slide
                End If
            End If
        Next shp
    Next slide
    
    If slideFound < 0 Then
    
        slideNum = slideNum - 1
        
        ActivePresentation.Slides.InsertFromFile _
        "C:\Users\first\Documents\Presentation_Local\FLC_offering.pptx", slideNum
        
        ActivePresentation.Slides(slideNum).Delete
    
    End If
    
    
'saveAsL

ActivePresentation.SaveCopyAs Environ("USERPROFILE") & "\Creative Cloud Files\Desktop\l.pptx"

End Sub


Sub changeBGtoBlack()
ActivePresentation.SlideMaster.Background.Fill.ForeColor.RGB = RGB(0, 0, 0)
End Sub

Sub changeBGtoGreen()
ActivePresentation.SlideMaster.Background.Fill.ForeColor.RGB = RGB(0, 255, 0)
End Sub


Sub SendTextToBottom()

ActivePresentation.SlideMaster.CustomLayouts(3).Shapes(1).TextFrame.VerticalAnchor = msoAnchorBottom

End Sub



Sub FindSlideWithPrayers()
    Dim slide As slide
    Dim slideNum As Integer
    
    For Each slide In ActivePresentation.Slides
        For Each shp In slide.Shapes
            If shp.HasTextFrame Then
                If InStr(1, shp.TextFrame.TextRange.Text, "PRAYERS OF INTERCESSION") > 0 Or InStr(1, shp.TextFrame.TextRange.Text, "PRAYERS OF THE CHURCH") > 0 Then
                    slideNum = slide.slideIndex
                    Exit For ' Exit loop if text is found on the slide
                End If
            End If
        Next shp
    Next slide
    
    ActivePresentation.Slides(slideNum).Select
    
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim filePath As String

    ' Path to the Word document
    filePath = "C:\Users\first\Dropbox\Pastor\Prayers\" & Format(Now, "YYYY-MM-DD") & ".doc"

    ' Create a new instance of Word application
    On Error Resume Next
    Set wdApp = CreateObject("Word.Application")
    On Error GoTo 0

    If wdApp Is Nothing Then
        MsgBox "Word is not installed on this computer.", vbExclamation
        Exit Sub
    End If

    ' Make Word visible
    wdApp.Visible = True

    ' Open the Word document
    Set wdDoc = wdApp.Documents.Open(filePath)

    ' Check if the document opened successfully
    If wdDoc Is Nothing Then
        MsgBox "Failed to open the document.", vbExclamation
    End If

    ' Clean up
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub


Sub FindSlideWithCollect()
    Dim slide As slide
    Dim slideNum As Integer
    
    For Each slide In ActivePresentation.Slides
        For Each shp In slide.Shapes
            If shp.HasTextFrame Then
                If InStr(1, shp.TextFrame.TextRange.Text, "Collect of the day") > 0 Then
                    slideNum = slide.slideIndex
                    Exit For ' Exit loop if text is found on the slide
                End If
            End If
        Next shp
    Next slide
    ActivePresentation.Slides(slideNum).Select
End Sub


Function MyRange(ByVal StartIndex As Long, ByVal StopIndex As Long) As Variant
    Dim A() As Long
    Dim i As Long


    ReDim A(StartIndex To StopIndex)
    For i = StartIndex To StopIndex: A(i) = i: Next
    MyRange = A
End Function


Sub FindSlideWithFLCoffering()
    Dim slide As slide
    Dim slideNum As Integer
    
    For Each slide In ActivePresentation.Slides
        For Each shp In slide.Shapes
            If shp.HasTextFrame Then
                If InStr(1, shp.TextFrame.TextRange.Text, "OFFERING:") > 0 Then
                    slideNum = slide.slideIndex
                    Exit For ' Exit loop if text is found on the slide
                End If
            End If
        Next shp
    Next slide
    
    slideNum = slideNum - 1
    
    ActivePresentation.Slides.InsertFromFile _
    "C:\Users\first\Documents\Presentation_Local\FLC_offering.pptx", slideNum
    
    ActivePresentation.Slides(slideNum).Delete
    

    
End Sub


Sub ReplaceLivestreamFiles()
    Dim folderPath As String
    Dim fileName As String
    Dim baseName As String
    Dim lastRow As Long
        
    ' Change this to the path of your folder
    folderPath = "C:\Users\first\Documents\Presentation_Local\Livestream_Replacements\"
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    fileName = Dir(folderPath & "*.*")
    lastRow = 1
    
    Do While fileName <> ""
        baseName = Left(fileName, InStrRev(fileName, ".") - 1)
        
        Dim slide As slide
        Dim notesText As String
        Dim slideFound As Boolean
        
        slideFound = False
        
        For Each slide In ActivePresentation.Slides
            notesText = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
            If InStr(1, notesText, baseName, vbTextCompare) > 0 Then
                slide.SlideShowTransition.AdvanceOnTime = msoTrue
                slide.SlideShowTransition.AdvanceTime = 0
                slideFound = True
                Exit For
            End If
        Next slide
        
        If slideFound < 0 Then
            
            Dim slideNum As Integer
            Dim slideNum2 As Integer
            
            slideNum = slide.slideIndex
            
            Dim CountSlidesNum As Integer
            Dim oPresentation As Presentation
            Set oPresentation = Presentations.Open(folderPath & fileName)
            
            slideNum2 = slideNum - oPresentation.Slides.Count + 1
            oPresentation.Close
            
            ActivePresentation.Slides.InsertFromFile _
            folderPath & fileName, slideNum
            
            ActivePresentation.Slides.Range(MyRange(slideNum2, slideNum)).Delete
            
        End If
        
        lastRow = lastRow + 1
        fileName = Dir
    Loop
    
End Sub



Sub FindSlideWithMTOoffering()
    Dim slide As slide
    Dim slideNum As Integer
    
    For Each slide In ActivePresentation.Slides
        For Each shp In slide.Shapes
            If shp.HasTextFrame Then
                If InStr(1, shp.TextFrame.TextRange.Text, "OFFERING:") > 0 Then
                    slideNum = slide.slideIndex
                    Exit For ' Exit loop if text is found on the slide
                End If
            End If
        Next shp
    Next slide
    
    slideNum = slideNum - 1
    
    ActivePresentation.Slides.InsertFromFile _
    "C:\Users\first\Documents\Presentation_Local\MTOoffering.pptx", slideNum
    
    ActivePresentation.Slides(slideNum).Delete
    

    
End Sub


Sub FindSlideWithMTObells()
    Dim slide As slide
    Dim slideNum As Integer
    
    For Each slide In ActivePresentation.Slides
        For Each shp In slide.Shapes
            If shp.HasTextFrame Then
                If InStr(1, shp.TextFrame.TextRange.Text, "MUSIC WITH CHOIR AND ORGAN FROM") > 0 Then
                    slideNum = slide.slideIndex
                    Exit For ' Exit loop if text is found on the slide
                End If
            End If
        Next shp
    Next slide
    
    slideNum = slideNum - 3
    
    ActivePresentation.Slides.InsertFromFile _
    "C:\Users\first\Documents\Presentation_Local\MTObellTower.pptx", slideNum
    

    
End Sub






Sub MTOmod_for_spptx()
    Dim slide As slide
    Dim slideNum As Integer
    
    'insert Giving slide
    
    For Each slide In ActivePresentation.Slides
        For Each shp In slide.Shapes
            If shp.HasTextFrame Then
                If InStr(1, shp.TextFrame.TextRange.Text, "OFFERING:") > 0 Then
                    slideNum = slide.slideIndex
                    Exit For ' Exit loop if text is found on the slide
                End If
            End If
        Next shp
    Next slide
    
    slideNum = slideNum - 1
    
    ActivePresentation.Slides.InsertFromFile _
    "C:\Users\first\Documents\Presentation_Local\MTOoffering.pptx", slideNum
    
    ActivePresentation.Slides(slideNum).Delete
    

    'ReplaceLetTheVineyardsMTO()

    Dim notesText As String
    Dim slideFound As Boolean
    
    slideFound = False
    
    For Each slide In ActivePresentation.Slides
        notesText = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
        If InStr(1, notesText, "Let_the_Vineyards_be_Fruitful-955", vbTextCompare) > 0 Then
            slide.SlideShowTransition.AdvanceOnTime = msoTrue
            slide.SlideShowTransition.AdvanceTime = 0
            slideFound = True
            Exit For
        End If
    Next slide
    
    If slideFound < 0 Then
        
        Dim slideNum2 As Integer
        
        slideNum = slide.slideIndex
        
        Dim CountSlidesNum As Integer
        Dim oPresentation As Presentation
        Set oPresentation = Presentations.Open("C:\Users\first\Documents\Presentation_Local\Let_the_Vineyards_be_Fruitful-955.pptx")
        
        slideNum2 = slideNum - oPresentation.Slides.Count + 1
        oPresentation.Close
        
        ActivePresentation.Slides.InsertFromFile _
        "C:\Users\first\Documents\Presentation_Local\Let_the_Vineyards_be_Fruitful-955.pptx", slideNum
        
        ActivePresentation.Slides.Range(MyRange(slideNum2, slideNum)).Delete
        
    End If
    
    
    'insert Bells
    
    For Each slide In ActivePresentation.Slides
        For Each shp In slide.Shapes
            If shp.HasTextFrame Then
                If InStr(1, shp.TextFrame.TextRange.Text, "MUSIC WITH CHOIR AND ORGAN FROM") > 0 Then
                    slideNum = slide.slideIndex
                    Exit For ' Exit loop if text is found on the slide
                End If
            End If
        Next shp
    Next slide
    
    slideNum = slideNum - 3
    
    ActivePresentation.Slides.InsertFromFile _
    "C:\Users\first\Documents\Presentation_Local\MTObellTower.pptx", slideNum
    
    
'ExportforRemote()

    Dim j As Integer
    Dim baseName2 As String
    Dim charCode As Integer

    Randomize
    For j = 1 To 3
        ' Generate a random number between 97 ("a") and 122 ("z")
        charCode = Int((26 * Rnd) + 97)
        baseName2 = baseName2 & Chr(charCode)
    Next j



'SaveSlidesAsJPG()
'TransitionTimes()
    Dim oSlide As slide
    Dim slideNumber As Integer
    Dim sImagePath As String
    Dim sImageName As String
    Dim transitionTime As Double

    ' Set the desired image path (e.g., "C:\Images\")
    sImagePath = "C:\Users\first\source\repos\PowerPoint-Remote 3006\PowerPoint Remote\bin\Debug\net5.0-windows\win-x64\wwwroot\powerpoint_parts\"
    
    Kill sImagePath & "*.*"

    ' Loop through each slide in the presentation
    For Each oSlide In ActivePresentation.Slides
        'get slide number
        slideNumber = oSlide.slideIndex
        'get transitionTime
        transitionTime = oSlide.SlideShowTransition.AdvanceTime
        'create txt file for transition time
        strFileName = sImagePath & baseName2 & oSlide.slideIndex & ".txt"
        intFileNum = FreeFile()
        Open strFileName For Output As intFileNum
        Print #intFileNum, oSlide.SlideShowTransition.AdvanceTime
        Close #intFileNum
        ' Construct the image file name (e.g., "1.jpg")
        sImageName = sImagePath & baseName2 & oSlide.slideIndex & ".jpg"

        ' Export the slide as a JPG image
        oSlide.Export sImageName, "JPG"
    Next oSlide
    
    
'//create slide count number txt file
    sImageName = Replace(sImageName, sImagePath & baseName2, "")
    sImageName = Replace(sImageName, ".jpg", "")
    
    strFileName = sImagePath & "0.txt"
    intFileNum = FreeFile()
    Open strFileName For Output As intFileNum
    Print #intFileNum, sImageName
    Print #intFileNum, baseName2
    Close #intFileNum
    
    
'ExportNotes()
    Dim oSl As slide
    Dim oSh As Shape
    Dim strNotesText As String
    
    ' Get the notes text
    For Each oSl In ActivePresentation.Slides
        For Each oSh In oSl.NotesPage.Shapes
            If oSh.PlaceholderFormat.Type = ppPlaceholderBody Then
                If oSh.HasTextFrame Then
                    If oSh.TextFrame.HasText Then
                        ' now write the text to file
                        strFileName = sImagePath & baseName2 & oSl.slideIndex & ".txt"
                        intFileNum = FreeFile()
                        Open strFileName For Append As intFileNum
                        Print #intFileNum, EncodeToHTML2(oSh.TextFrame.TextRange.Text)
                        Close #intFileNum
                    End If
                End If
            End If
        Next oSh
    Next oSl
    

    
End Sub


Sub MTOmod_for_lpptx()
    Dim slide As slide
    Dim slideNum As Integer
    
    'insert Giving slide
    
    For Each slide In ActivePresentation.Slides
        For Each shp In slide.Shapes
            If shp.HasTextFrame Then
                If InStr(1, shp.TextFrame.TextRange.Text, "OFFERING:") > 0 Then
                    slideNum = slide.slideIndex
                    Exit For ' Exit loop if text is found on the slide
                End If
            End If
        Next shp
    Next slide
    
    slideNum = slideNum - 1
    
    ActivePresentation.Slides.InsertFromFile _
    "C:\Users\first\Documents\Presentation_Local\MTOoffering.pptx", slideNum
    
    ActivePresentation.Slides(slideNum).Delete
    

    'ReplaceLetTheVineyardsMTO()

    Dim notesText As String
    Dim slideFound As Boolean
    
    slideFound = False
    
    For Each slide In ActivePresentation.Slides
        notesText = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
        If InStr(1, notesText, "Let_the_Vineyards_be_Fruitful-955", vbTextCompare) > 0 Then
            slide.SlideShowTransition.AdvanceOnTime = msoTrue
            slide.SlideShowTransition.AdvanceTime = 0
            slideFound = True
            Exit For
        End If
    Next slide
    
    If slideFound < 0 Then
        
        Dim slideNum2 As Integer
        
        slideNum = slide.slideIndex
        
        Dim CountSlidesNum As Integer
        Dim oPresentation As Presentation
        Set oPresentation = Presentations.Open("C:\Users\first\Documents\Presentation_Local\Let_the_Vineyards_be_Fruitful-955.pptx")
        
        slideNum2 = slideNum - oPresentation.Slides.Count + 1
        oPresentation.Close
        
        ActivePresentation.Slides.InsertFromFile _
        "C:\Users\first\Documents\Presentation_Local\Let_the_Vineyards_be_Fruitful-955.pptx", slideNum
        
        ActivePresentation.Slides.Range(MyRange(slideNum2, slideNum)).Delete
        
    End If
    

'insert Bells
    
    For Each slide In ActivePresentation.Slides
        For Each shp In slide.Shapes
            If shp.HasTextFrame Then
                If InStr(1, shp.TextFrame.TextRange.Text, "MUSIC WITH CHOIR AND ORGAN FROM") > 0 Then
                    slideNum = slide.slideIndex
                    Exit For ' Exit loop if text is found on the slide
                End If
            End If
        Next shp
    Next slide
    
    slideNum = slideNum - 3
    
    ActivePresentation.Slides.InsertFromFile _
    "C:\Users\first\Documents\Presentation_Local\MTObellTower.pptx", slideNum

    
    
'RemoveAudioFromAllSlides

   Dim oSl As slide
    Dim oSh As Shape
    Dim x As Long
    
    For Each oSl In ActivePresentation.Slides
        For x = oSl.Shapes.Count To 1 Step -1
            With oSl.Shapes(x)
                If .Type = msoMedia Then
                    If .MediaType = ppMediaTypeSound Then
                        .Delete
                    End If
                End If
            End With
        Next     ' x
    Next    ' Slide

End Sub


Sub ExportNotesTXT()

    Dim oSl As slide
    Dim oSh As Shape
    Dim strNotesText As String
    
    ' Get the notes text
    For Each oSl In ActivePresentation.Slides
        For Each oSh In oSl.NotesPage.Shapes
            If oSh.PlaceholderFormat.Type = ppPlaceholderBody Then
                If oSh.HasTextFrame Then
                    If oSh.TextFrame.HasText Then
                        ' now write the text to file
                        strFileName = sImagePath & baseName2 & oSl.slideIndex & ".txt"
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
    
Sub ExportNotesAsRTF()

    Dim oSlide As slide
    Dim oShape As Shape
    Dim strRTF As String
    Dim strFileName As String
    Dim intFileNum As Integer
    Dim slideIndex As Integer

    ' Loop through each slide
    For Each oSlide In ActivePresentation.Slides
        slideIndex = oSlide.slideIndex
        For Each oShape In oSlide.NotesPage.Shapes
            If oShape.PlaceholderFormat.Type = ppPlaceholderBody Then
                If oShape.HasTextFrame Then
                    If oShape.TextFrame.HasText Then
                        ' Get RTF text
                        strRTF = oShape.TextFrame.TextRange.RTF

                        ' Define file name for each slide's notes
                        strFileName = ActivePresentation.Path & "\" & _
                                      "Slide_" & slideIndex & "_Notes.rtf"

                        ' Write RTF to file
                        intFileNum = FreeFile
                        Open strFileName For Binary Access Write As #intFileNum
                        Put #intFileNum, , strRTF
                        Close #intFileNum
                    End If
                End If
            End If
        Next oShape
    Next oSlide

    MsgBox "Slide notes exported as RTF files."
    
End Sub

Function EncodeToHTML(sInput As String) As String
    Dim oHTML As Object
    Set oHTML = CreateObject("htmlfile")
    
    oHTML.Body.innerText = sInput
    EncodeToHTML = oHTML.Body.innerHTML
    
    Set oHTML = Nothing
End Function



