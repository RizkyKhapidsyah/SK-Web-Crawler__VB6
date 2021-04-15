Attribute VB_Name = "Lc2kSupport1"
Dim fdomains, fgdomains, fhtml, fcgiasp, fexe, fdupl _
, ftxt, fhlp, fem, fot

' -------------------------------------------------
' if key has been used before then get its index
' (view.nodes.item(=index).text = "whatever")
' -------------------------------------------------
Function GetKeyIndex(Key, view As TreeView)
For X = 1 To view.Nodes.Count
 If view.Nodes.Item(X).Key = Key Then
 GetKeyIndex = X
 Exit Function
 End If
 Next X
 GetKeyIndex = -1
 Exit Function
End Function
' -------------------------------------------------
' Find out if the key has been used before.
' -------------------------------------------------
Function HasKeyBeenUsed(view As Object, Key)
For X = 1 To view.Nodes.Count
 If view.Nodes.Item(X).Key = Key Then
 HasKeyBeenUsed = True
 Exit Function
 End If
 Next X
 HasKeyBeenUsed = False
 Exit Function
End Function
' -------------------------------------------------
' Takes care of setting the icons
' find the parents and everthing else.
Sub AddNode(view As Object, ParentDocument, ThisDocument)
On Local Error Resume Next
' -------------------------------------------------
' Check for parent of ThisDocument
If VFindParent(view, ParentDocument) = False Then
    ' -------------------------------------------------
    ' Parent not found alert user.. for debuging..
    'MsgBox "Parent:" & ParentDocument & Chr$(13) & "ThisDoc:" & ThisDocument, vbExclamation, "Parent Key Missing"
    view.Nodes.Add , , ParentDocument, ParentDocument, 8
    SMParentMissing = True
    SMVParentDocument = ParentDocument
    End If
   
' -------------------------------------------------
' Set Normal Icon
aicon = 13

' -------------------------------------------------
' Get Document Type
doctype = GetUrlDocumentType(ThisDocument)

' -------------------------------------------------
' Registerd Normal Domains
' -------------------------------------------------
If doctype = "com" Or doctype = "net" Or doctype = "org" Then
aicon = 1
fdomains = fdomains + 1
view.Nodes.Item(GetKeyIndex("rd", view)).Text = "Registered Domains:" & fdomains
End If
' -------------------------------------------------
' Goverment Domains
' -------------------------------------------------
If doctype = "gov" Then
fgdomains = fgdomains + 1
view.Nodes.Item(GetKeyIndex("gd", view)).Text = "Goverment Domains:" & fgdomains
aicon = 12
End If

' -------------------------------------------------
' HTML or HTM Documents
' -------------------------------------------------
If doctype = "htm" Or doctype = "html" Then
aicon = 3
fhtml = fhtml + 1
view.Nodes.Item(GetKeyIndex("hd", view)).Text = "HyperTextTransferProtocal:" & fhtml
End If
' -------------------------------------------------
' ASP
' -------------------------------------------------
If doctype = "asp" Then
fcgiasp = fcgiasp + 1
view.Nodes.Item(GetKeyIndex("cg", view)).Text = "CGI & ASP:" & fcgiasp
aicon = 4
End If
' -------------------------------------------------
' CGI
' -------------------------------------------------
If doctype = "cgi" Then
fcgiasp = fcgiasp + 1
view.Nodes.Item(GetKeyIndex("cg", view)).Text = "CGI & ASP:" & fcgiasp
aicon = 6
End If
' -------------------------------------------------
' Executables! Setups! Compressed Files! Encrypted!
' -------------------------------------------------
If doctype = "exe" Or doctype = "zip" Or doctype = "cab" Or doctype = "ace" Or doctype = "dll" Then
fexe = fexe + 1
view.Nodes.Item(GetKeyIndex("ex", view)).Text = "Executable/Compressed: " & fexe
aicon = 7
End If
' -------------------------------------------------
' Link, mostly used by the internal working of
' this application
' -------------------------------------------------
If doctype = "dl" Then
fdupl = fdupl + 1
view.Nodes.Item(GetKeyIndex("dl", view)).Text = "DuplicateLinks:" & fdupl
aicon = 8
End If
' -------------------------------------------------
' Readable Documents (TextEditor,WordProcessor)
' -------------------------------------------------
If doctype = "txt " Or doctype = "doc" Or doctype = "wrd" Then
ftxt = ftxt + 1
view.Nodes.Item(GetKeyIndex("tx", view)).Text = "Text/Documents:" & ftxt
aicon = 9
End If
' -------------------------------------------------
' DocumentType = HtmlHelp or WindowsHelp
' -------------------------------------------------
If doctype = "chp" Or doctype = "hlp" Then
fhlp = fhlp + 1
view.Nodes.Item(GetKeyIndex("hl", view)).Text = "HelpFiles(hlp):" & fhlp
aicon = 10
End If

' -------------------------------------------------
' DocumentType = Email!
' -------------------------------------------------
If InStr(1, ThisDocument, "mailto:") = 1 Then
aicon = 2
view.Nodes.Item(GetKeyIndex("em", view)).Text = "EmailAddress:" & fem
End If

view.Nodes.Add ParentDocument, 4, ThisDocument, ThisDocument, aicon

' temp!!!!
WriteRawTreeData ParentDocument, 4, ThisDocument, ThisDocument, aicon

    If aicon = 5 Then
    fot = fot + 1
    view.Nodes.Item(GetKeyIndex("ot", view)).Text = "Other:" & fot
    End If
End Sub
'
'
' write rawtree data
'
Sub WriteRawTreeData(Parent, aType, Key, Child, Icon)
Open "c:\urls.rtd" For Append As #1
Print #1, Parent
Print #1, aType
Print #1, Key
Print #1, Child
Print #1, Icon
Close #1
End Sub

' -------------------------------------------------
' Gets the documents path is there is one.
' http://www.mysite.com/mypage.htm
' (=http://www.mysite.com/) i think the "/" is on the end
' -------------------------------------------------
Function GetUrlPath(document)
If Len(document) = 0 Then Exit Function
Dim X As Long: X = Len(document)
Do
    If Mid(document, X, 1) = "/" Then
    GetUrlPath = Mid(document, 1, X - 1) & "/"
    Exit Function
    End If
X = X - 1
Loop Until X = 1
End Function
' -------------------------------------------------
' find parent in view list, is used to see if document is
' already in the list also..
' -------------------------------------------------
Function VFindParent(view As Object, Parent) As Boolean
For X = 1 To view.Nodes.Count
 If view.Nodes.Item(X).Key = Parent Then
 VFindParent = True: Exit Function
 End If
 Next X
 VFindParent = False
End Function
' -------------------------------------------------
' Gets Documents Domain Type www.mysite.com (=COM)
' -------------------------------------------------
Function GetUrlDocumentDomainType(document)
On Local Error GoTo le
Dim X As Long: X = 1
' if www is present or http://
If InStr(1, document, "www.") > 0 Then
Y = InStr(1, document, "www.") + 5
Else
If InStr(1, document, "http://") > 0 Then
Y = InStr(1, document, "http://") + 8
Else
Y = 1
End If
End If

If InStr(Y, document, ".") > 0 Then
yyy = InStr(Y, document, ".")
Else
Exit Function
End If
' http://www.aol.com(/)hello/mypage.htm
' if / is presetn seperating domain from path
If InStr(Y, document, "/") > 0 Then
yy = InStr(Y, document, "/") - 1
Else
yy = Len(document)
End If
GetUrlDocumentDomainType = Mid(document, yyy + 1, (yy - yyy))
Exit Function
le: GetUrlDocumentDomainType = ""
End Function
' -------------------------------------------------
' Gets url extension page.htm (EXT=htm)
' -------------------------------------------------
Function GetUrlDocumentType(document)
Dim X As Long: X = Len(document)
' if path contains only a folder then this means no file
' attached to extract! hehe
If bwithfolder = True Then Exit Function ' no file.
Do
    If Mid(document, X, 1) = "." Then
    GetUrlDocumentType = Mid(document, X + 1, Len(document))
    Exit Function
    End If
X = X - 1
Loop Until X = 1
End Function
' -------------------------------------------------
' Fixes Url. ParentDocument is the document
' the link was found on. The link is refranced by Document
' -------------------------------------------------
Function FixUrl(document As Variant, ParentDocument As Variant)
bwithopar = False
bwithdomain = False
bwithfolder = False
' is document without parent /~
If Mid(document, 1, 1) = "/" Then
bwithopar = True
End If

' is document domain (com,net,org)
aext = GetUrlDocumentDomainType(document)
If aext = "org" Or aext = "com" Or aext = "net" Then bwithdomain = True
' is document folder directory or file
' document is folder ~/
If InStr(1, document, "http://") > 0 Or InStr(1, document, "www.") > 0 Then bwithdomain = True

If Mid(document, Len(document), 1) = "/" Then
bwithfolder = True
Else
' document is file ~?
bwithfolder = False
End If

If bwithdomain = False And bwithopar = False Then GoTo FixPath
' fix document path if it has no parent
If bwithopar = True Then
FixPath:
    'fix document if url path is curropted by getpath.
    urlpath = GetUrlPath(ParentDocument)
    If LCase(urlpath) = "http://" Then
        
        If Mid(urlpath, Len(urlpath), 1) = "/" Then
        urlpath = ParentDocument
        Else
        urlpath = ParentDocument & "/"
        End If
        
    End If
        If Mid(document, 1, 1) <> "/" Then document = "/" & document
        If urlpath <> "" Then
        If Mid(urlpath, Len(urlpath), 1) <> "/" Then urlpath = urlpath & "/"
        End If
        
    document = urlpath & Mid(document, 2, Len(document))
    End If
FixUrl = document
End Function
