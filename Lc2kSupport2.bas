Attribute VB_Name = "Lc2kSupport2"

Dim EmailMessageRecp, EmailMessage

'
'
'   Get(s) Complete Domain Name And Type
'
Public _
Function GetUrlDomainName(varUrl)

        If cmlstop = True Then
        OODMODE = True
        Exit Function
        Else
        End If

Dim ch1 As Long
                    ' lower case url
                    varUrl = LCase( _
                                    varUrl)
                    
ch1 = InStr(1, varUrl, "http://")
If ch1 > 0 Then ch1 = ch1 + 7
If ch1 < 1 Then ch1 = 1
ch2 = InStr(ch1, varUrl, "www.")
If ch2 < 1 Then
    If ch1 = 1 Then
    ch2 = 1
    Else
    ch2 = ch1
    End If
End If

ch3 = InStr(ch2, varUrl, "/")
' if no / at end then keep going
If ch3 = 0 Then ch3 = Len(varUrl) + 1

' ch2 - start of domain name
' ch3 - end of domain name+typename(com,net,org)
GetUrlDomainName = Mid(varUrl, ch2, (ch3 - ch2))

'If InStr(GetUrlDomainName, "r.hotbot") Then
'exitflag = True
'End If


End Function

'
' Removes Document Arguments and Extension
'
Public _
Function RemoveUrlArguments(varUrl)
i = bbInstr(varUrl, "/.")
RemoveUrlArguments = Mid(varUrl, 1, i - 1)
End Function

'
' BackWord ByteInstr (multi-char find)
'
Public _
Function bbInstr(varString, varTofind)
i = Len(varString)
Do
    For Y = 1 To Len(varTofind)
    If Mid(varString, i, 1) = Mid(varTofind, Y, 1) Then
        bbInstr = i
        Exit Function
        End If
    Next Y
    i = i - 1
Loop Until i = 1
    bbInstr = -1
End Function

'
'
'
Sub AddMajorDomain(Url, view As TreeView)

        If cmlstop = True Then
        Exit Sub
        Else
        End If


' for all major domains ******
        MajorDomain = GetUrlDomainName(Url)


If exitflag = True Then Exit Sub

        If GetKeyIndex(MajorDomain, view) = -1 Then
        view.Nodes.Add "mdomains", 4, MajorDomain, MajorDomain, 1
        view.Nodes.Add MajorDomain, 4, MajorDomain & "Count", "1", 14
        view.Nodes.Add MajorDomain, 4, MajorDomain & "Email", "WebMaster@" & MajorDomain, 2
        EmailMessageRecp = EmailMessageRecp & "WebMaster@" & MajorDomain & ","
        EmailMessage = EmailMessage & MajorDomain & Chr$(13)
        Clipboard.SetText EmailMessageRecp & Chr$(13) & EmailMessage
        Else
        ' *** below code also includes icon settings for servers ****
        
        ' reg server
        i = GetKeyIndex(MajorDomain & "Count", view)
        ii = GetKeyIndex(MajorDomain, view)
        view.Nodes.Item(i).Text = view.Nodes.Item(i).Text + 1
       
        ' meduim server
        If view.Nodes.Item(i).Text > 50 Then
        view.Nodes.Item(ii).Image = 15
        End If
        ' Big server
        If view.Nodes.Item(i).Text > 150 Then
        view.Nodes.Item(ii).Image = 16
        End If
        
        If view.Nodes.Item(i).Text > 250 Then
        view.Nodes.Item(ii).Image = 17
        End If
        
        If view.Nodes.Item(i).Text > 350 Then
        view.Nodes.Item(ii).Image = 18
        End If
        End If
      
End Sub
Sub RunBrowserWindow(BrowserObject As WebBrowser, SetUrl, SetFrameName)
BrowserObject.Navigate2 SetUrl, SetFrameName
End Sub
