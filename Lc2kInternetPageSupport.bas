Attribute VB_Name = "Lc2kInternetPageSupport"
Public Declare Function shellexecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
String, ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public cmlstop As Boolean
Public exitflag As Boolean
Public TempInternetFilesPath
Public llhwnd

'
' taks care of saving file(s) to disk found on web.
'
'               ** note below **
'   bug: needs to be string (PUT VALUE!!) fix!
'
Public Sub StoreInternetPage(varPageData, varPageAddress)
On Local Error Resume Next
' encode address to filename standerd
s = encode(varPageAddress, "W", "WW")
s = encode(s, "/", "WS")
s = encode(s, ":", "WC")
s = encode(s, "?", "WQ")
s = encode(s, "%", "WP")
s = encode(s, "$", "WD")
MkDir "c:\tempinet\"
Open "c:\tempinet\" & s For Binary As #1
Put #1, 1, varPageData
Close #1
End Sub
'
' Reverses the encode effect.
'
Public Function RestoreFilenameToAddress(varFilename)
s = encode(varFilename, "WD", "$")
s = encode(s, "WP", "%")
s = encode(s, "WQ", "?")
s = encode(s, "WC", ":")
s = encode(s, "WS", "/")
RestoreFilenameToAddress = encode(s, "WW", "W")
End Function
'
' Use: To encode web address' to the filename standerd
'
Public Function encode(varString, varBlock, varNewBlock)
For i = 1 To Len(varString)
    If Mid(varString, i, Len(varBlock)) = varBlock Then
    encode = encode & varNewBlock
    i = i + Len(varBlock) - 1
    Else
    encode = encode & Mid(varString, i, 1)
    End If
Next i
End Function
