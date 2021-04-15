Attribute VB_Name = "Lc2kRunModule"
'
Dim Searched(900) As Variant, SearchedIndex As Long
Dim OODMODE As Boolean, OOD_PURL As String
Dim INFMODE As Boolean


'
' Settings
'
Sub Lc2kSetMode_OOD(bValue As Boolean, purl)
If bValue = True Then
OODMODE = True
OOD_PURL = purl
Else
OODMODE = False
End If
End Sub

'
' Run Module
'
'
Sub Lc2kEngine(Url As String _
                            , view As TreeView _
                            , inet As inet _
                            , PMaxNodes As Long _
                            , ProcessorTime As Long _
                            , Frm As Form _
                            )
                            
                            
                            
On Error GoTo errh
        
        If cmlstop = True Then
         GoTo CleanUp
        Else
        End If
        
        
   Dim Stack As New W32LIB_Stack
   Dim Strip As New HtmlStrip
   Dim CurUrl As String
   Dim SDevice As New FSD

INFMODE = True


'If DLF = True Then ' clear directory
'RmDir "c:\lc2k\dlf\"
'End If



    Stack.Reset 1
    Stack.Reset 0
   
    If INFMODE = True Then
    view.Nodes.Clear 'cmljr
    
    view.Nodes.Add , , "info", "WebDirectoryInfoNode", 14
    view.Nodes.Add "info", 4, "rd", , 14
    view.Nodes.Add "info", 4, "gd", , 14
    view.Nodes.Add "info", 4, "tx", , 14
    view.Nodes.Add "info", 4, "cg", , 14
    view.Nodes.Add "info", 4, "hd", , 14
    view.Nodes.Add "info", 4, "ex", , 14
    view.Nodes.Add "info", 4, "hl", , 14
    view.Nodes.Add "info", 4, "em", , 14
    view.Nodes.Add "info", 4, "ot", , 14
    End If
    
    ' **** Major Domains *****
    view.Nodes.Add , , "mdomains", "Domain Listing", 14
    
    
    ' OOD MODE
    If OODMODE = False Then
   view.Nodes.Add , , Url, Url, 1
    End If
   
   Stack.Push 0, Url
   
   Do
   
   Do
repop:
        WaitProcessorTime ProcessorTime, True
        CurUrl = Stack.pop(0)
        
        If InStr(1, CurUrl, Wnd.excludetxt) > 0 Then GoTo repop  'cml
        
        If cmlstop = True Then
         GoTo CleanUp
        Else
        End If
        
        
        Frm.Show
        Frm.Caption = "Init (" & CurUrl & ")"
        
        
        If HasUrlBeenSearchedBefore(CurUrl) = True Then
            If Stack.lne(0) = True Then Exit Do
            GoTo repop
            End If
        SetUrlSearched CurUrl
        
        Frm.Caption = "Getting (" & CurUrl & ")"
        '
        '
        '
        '
        '
       
        oldimage = view.Nodes.Item(GetKeyIndex(CurUrl, view)).Image
        view.Nodes.Item(GetKeyIndex(CurUrl, view)).Image = 11
        
        Data = GetDocument(CurUrl, inet)
        view.Nodes.Item(GetKeyIndex(CurUrl, view)).Image = oldimage
        
    
        
        
        
        
        ' ***** StoreInternetPage ****
        StoreInternetPage Data, CurUrl
        ' ***** Add speacil node section
        
        
        'Dim xurl
        AddMajorDomain CurUrl, view
        
        
        Frm.Caption = "Extracting(" & CurUrl & ")"
        
        Strip.SetClassToFindLinks
        Strip.StripDocument Data, False
        Strip.SetClassToFindScr
        Strip.StripDocument Data, True
        i = 0
        
        Do
        WaitProcessorTime ProcessorTime, True
            Gt = Strip.aGet(i)
            If Gt = "" Then Exit Do
            

If exitflag = True Then GoTo CleanUp
            
            
            Gt = FixUrl(Gt, CurUrl)
                If HasUrlBeenSearchedBefore(Gt) = False Then
                Stack.Push 1, Gt
                Else
                Gt = Gt & ".dl"
                End If
            
            '
            '
            '   Check MaxNodes
            '
            If PMaxNodes < view.Nodes.Count Then GoTo CleanUp
            
            ' OODMODE_PURL
            If OODMODE = True Then CurUrl = OOD_PURL
           
           
            ' ***** StoreInternetPage ****
            StoreInternetPage "", Gt
            ' ***** Add Major Domain ****
            AddMajorDomain Gt, view
            
            AddNode view, CurUrl, Gt
            i = i + 1
        Loop While Strip.aGet(i) <> ""
            ' OODMODE EXIT
            If OODMODE = True Then GoTo CleanUp
            
    Loop Until Stack.lne(0) = True
        Stack.Reset 0
        Stack.CopyStack 1, 0
        Stack.Reset 1
        WaitProcessorTime ProcessorTime, True
        
    Loop Until Stack.Peek(0, 0) = ""
CleanUp:
'Stack.LoadDebugWindow
'Stack.WaitForDebugWindow
Stack.Reset 0: Stack.Reset 1
Frm.Caption = "Done. LastDoc(" & CurUrl & ") Nodes:" & view.Nodes.Count
Exit Sub


errh:
MsgBox Err.Description
Sleep 500
Err.Clear
GoTo CleanUp
Stack.Reset 0: Stack.Reset 1
SearchedIndex = 0
GoTo repop

Exit Sub
                
End Sub
                            
                            
Function HasUrlBeenSearchedBefore(Url) As Boolean
For i = 0 To SearchedIndex
    If Searched(i) = Url Then
        HasUrlBeenSearchedBefore = True
        Exit Function
        End If
        
Next i
        HasUrlBeenSearchedBefore = False
End Function
Sub SetUrlSearched(Url)
Searched(SearchedIndex + 1) = Url
SearchedIndex = SearchedIndex + 1
End Sub

Sub WaitProcessorTime(ProcessorTime, UseL As Boolean)
If UseL = True Then
For X = 0 To ProcessorTime * 2
    DoEvents
Next X
Exit Sub
Else
s = Timer
Do While Timer - s < ProcessorTime
DoEvents
Loop
End If
End Sub

Function GetDocument(Url, inet As inet)
On Local Error GoTo gde
    GetDocument = LCase(inet.OpenURL(Url))
    Exit Function
gde:
    GetDocument = ""
    
    
End Function

