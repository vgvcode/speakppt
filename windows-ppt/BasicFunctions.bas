Attribute VB_Name = "BasicFunctions"
Sub OnSlideShowPageChange()
    Dim i As Integer
    i = ActivePresentation.SlideShowWindow.View.CurrentShowPosition
    
    Set myDocument = ActivePresentation.Slides(i)
    
    For Each s In myDocument.Shapes
    
        If s.HasTextFrame Then
    
            With s.TextFrame
    
                If .HasText Then MsgBox .TextRange.Text
    
            End With
    
        End If
    
    Next
End Sub

Sub ShowTextInSlideNum(slideNum)
    sld = Application.ActiveWindow.View.Slide
'    Set myDocument = ActivePresentation.Slides(slideNum)

    For Each s In sld.Shapes
    
        If s.HasTextFrame Then
    
            With s.TextFrame
    
                If .HasText Then MsgBox .TextRange.Text
    
            End With
    
        End If
    
    Next
End Sub
Sub TestFn(i)
    MsgBox i
End Sub

Sub Save_PowerPoint_Slide_as_Images()
    Dim sImagePath As String
    Dim sImageName As String
    Dim sld As Slide '* Slide Object
    Dim lScaleWidth As Long '* Scale Width
    Dim lScaleHeight As Long '* Scale Height
    On Error GoTo Err_ImageSave
    sImagePath = "/Users/vgvcode/Documents/python/polly/output/"
    For Each sld In ActivePresentation.Slides
        sImageName = sld.Name & ".jpg"
        sld.Export sImagePath & sImageName, "JPG"
    Next sld
Err_ImageSave:
    If Err <> 0 Then
        MsgBox Err.Description
    End If
    
End Sub
Sub PlaySlideShow()
    ActivePresentation.SlideShowSettings.Run
End Sub
Sub Save_PowerPoint_Slide_as_Images2()
    ActivePresentation.Export Path:="/Users/vgvcode/Documents/python/polly/output/", FilterName:="JPG", ScaleWidth:=800, ScaleHeight:=600
End Sub

Sub WriteTextInSlideToFile()
    Set sld = ActivePresentation.Slides(1)
'    Set myDocument = ActivePresentation.Slides(slideNum)

    ActivePresentation.Slides(1).NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.InsertAfter "Added Text"
    
    filename = "1.txt"
    
    For Each s In sld.Shapes
    
        If s.HasTextFrame Then
    
            With s.TextFrame
    
                If .HasText Then
                    MsgBox .TextRange.Text
                    txt = .TextRange.Text
                    TextFile_Create filename, txt
                End If
            End With
    
        End If
    
    Next

End Sub
Sub TextFile_Create(filename, txt)
    'PURPOSE: Add More Text To The End Of A Text File
    'SOURCE: www.TheSpreadsheetGuru.com
    
    Dim TextFile As Integer
    Dim FilePath As String
    
    'What is the file path and name for the new text file?
      FilePath = "/Users/vgvcode/Documents/python/polly/output/" + filename
    
    'Determine the next file number available for use by the FileOpen function
      TextFile = FreeFile
    
    'Open the text file
      Open FilePath For Append As TextFile
    
    'Write some lines of text
      Print #TextFile, txt
      
    'Save & Close Text File
      Close TextFile

End Sub

Sub ReadTextFromNotesInSlide(Optional ByVal slideNum As Integer = 1)

    Set sld = ActivePresentation.Slides(slideNum)

    Set myNotes = ActivePresentation.Slides(slideNum).NotesPage.Shapes.Placeholders(2)
    
    If myNotes.HasTextFrame Then
        With myNotes.TextFrame
            If .HasText Then
                MsgBox .TextRange.Text
            End If
        End With
    End If
    
End Sub
Sub ReadTextFromNotesInPresentation()

    Dim sld As Slide

    For Each sld In ActivePresentation.Slides
        ReadTextFromNotesInSlide sld.SlideNumber
    Next sld

End Sub
Sub JsonExample()
    'needs JsonConverter.bas and Dictionary.cls
    Dim Json As Object
    Set Json = JsonConverter.ParseJson("{""a"":123,""b"":[1,2,3,4],""c"":{""d"":456}}")
    
    ' Json("a") -> 123
    ' Json("b")(2) -> 2
    ' Json("c")("d") -> 456
    Json("c")("e") = 789
    
    Debug.Print JsonConverter.ConvertToJson(Json)
    ' -> "{"a":123,"b":[1,2,3,4],"c":{"d":456,"e":789}}"
    
    Debug.Print JsonConverter.ConvertToJson(Json, Whitespace:=2)
    ' -> "{
    '       "a": 123,
    '       "b": [
    '         1,
    '         2,
    '         3,
    '         4
    '       ],
    '       "c": {
    '         "d": 456,
    '         "e": 789
    '       }
    '     }"
End Sub
