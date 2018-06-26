Attribute VB_Name = "ssml"
Sub GetSsmlFromAllSlides()
    Dim sld As Slide
    Dim ssml As String, baseName As String
    Dim json As Object

    baseName = ActivePresentation.name
    
    For Each sld In ActivePresentation.Slides
    
        json = GetJsonFromNotesInSlide(sld.SlideNumber)
        ssml = GetSsmlFromSlide(sld.SlideNumber)
        
        
        SaveSsml baseName, sld.SlideNumber, ssml
    Next sld

End Sub
Function GetSsmlFromSlide(slideNum, json)
    Dim ssml As String, txt As String
    Dim sld As Object
    
    Set sld = ActivePresentation.Slides(slideNum)
    ssml = "'<speak>"
    
    For Each s In sld.Shapes
        If s.HasTextFrame Then
            With s.TextFrame
                If .HasText Then
                    txt = .TextRange.Text
                    ssml = ssml & "<p>" & txt & "</p>"
                End If
            End With
        End If
    Next
    
    ssml = ssml & "</speak>'"
    Debug.Print ssml
    
    GetSsmlFromSlide = ssml
    
End Function
Sub SaveSsml(baseName, slideNum, ssmlTxt, Optional ByVal voiceId As String = "Joanna")
    SsmlToVoice baseName, slideNum, ssmlTxt, voiceId
End Sub
Sub SsmlToVoice(baseName, slideNum, txt, voiceId)

    Dim fileName As String, arg As String
    fileName = "'/Users/vgvcode/Library/Application Scripts/com.microsoft.Powerpoint/'" & baseName & "." & slideNum & ".mp3"
    args = voiceId & "#" & txt & "#" & fileName
    Debug.Print args

    RunMyScript = AppleScriptTask("Text2Speech.scpt", "Speak", args)
End Sub
Function GetNotesInSlideAsJson(slideNum)

    Dim json As Object
    
    On Error GoTo EH
    Set sld = ActivePresentation.Slides(slideNum)

    Set myNotes = ActivePresentation.Slides(slideNum).NotesPage.Shapes.Placeholders(2)
    
    If myNotes.HasTextFrame Then
        With myNotes.TextFrame
            If .HasText Then
                'MsgBox .TextRange.Text
                Set json = JsonConverter.ParseJson(.TextRange.Text)
                'MsgBox json("Type")
                Set GetNotesInSlideAsJson = json
            End If
        End With
    End If
    
EH:
    errStr = "{""Error"": Err.Description}"
    GetNotesInSlideAsJson = JsonConverter.ParseJson(errStr)
    
End Function
Function GetJsonFromNotesInSlide(Optional ByVal slideNum As Integer = 1)

    Dim json As Object
    
    Set sld = ActivePresentation.Slides(slideNum)

    Set myNotes = ActivePresentation.Slides(slideNum).NotesPage.Shapes.Placeholders(2)
    
    If myNotes.HasTextFrame Then
        With myNotes.TextFrame
            If .HasText Then
                Set json = JsonConverter.ParseJson(.TextRange.Text)
                GetJsonFromNotesInSlide = json
            End If
        End With
    End If
    
EH:
    Set json = JsonConverter.ParseJson("{}")
    GetJsonFromNotesInSlide = json
    
End Function



