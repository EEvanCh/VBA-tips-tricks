Function ExtractChars(rng As Range) As String
    Dim strInput As String
    Dim regex As Object
    Dim strOutput As String
    
    If rng Is Nothing Or IsEmpty(rng) Then
        ExtractChars = ""
        Exit Function
    End If
    
    On Error Resume Next
    strInput = CStr(rng.value)
    If Err.Number <> 0 Then
        ExtractChars = ""
        Exit Function
    End If
    On Error GoTo 0
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = "[^A-Za-zА-Яа-я0-9]"
        .Global = True
    End With
    
    strOutput = regex.Replace(strInput, "")
    
    ExtractChars = LCase(strOutput)
End Function
