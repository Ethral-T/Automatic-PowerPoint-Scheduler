Set objArgs = WScript.Arguments

If objArgs.Count >= 1 Then
    WScript.Echo "Usage: ConvertSlides.vbs <OutputFolder>"
    WScript.Quit 1
End If

OutputFolder = objArgs(0)

' Create a PowerPoint application object
Set objPPT = CreateObject("PowerPoint.Application")
objPPT.Visible = True

' Process all PowerPoint files in the current directory
Set fso = CreateObject("Scripting.FileSystemObject")
For Each file In fso.getFolder(".").Files
    ext = LCase(fso.GetExtensionName(file))
    If ext = "ppt" Or ext = "pptx" Then
        Set objPresentation = objPPT.Presentations.Open(file.Path)
        For Each oSl In objPresentation.ConvertSlides
            num = Right("000" & oSl.SlideIndex, 4)
            oSl.Export OutputFolder & "\" & "_Slide" & oSl.SlideIndex & ".jpg", "JPG", 3840, 2160
        Next
        objPresentation.Close
    End If
Next

' Close PowerPoint application
objPPT.Quit
Set objPPT = Nothing