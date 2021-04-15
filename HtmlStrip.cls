VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HtmlStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim Stripped(500) As Variant, CurrentIndex As Long

Public idMaxLenth As Long
Public szBlockStart As Variant
Public szBlockEnd As Variant
Function aGet(index) As Variant
aGet = Stripped(index)
End Function
Sub aPut(index, value As Variant)
Stripped(index) = value
End Sub
'
' <A HREF="MAILTO: FINDS ALL EMAIL LINKS"></A>
'
Sub SetClassToFindEmails()
' set header and footer block varibles
' these determine what is to be found!
Me.szBlockStart = "<a href=" & Chr$(34) & "mailto:"
Me.szBlockEnd = Chr$(34) ' you could add [& ">"] ..
' ---- this has never been tested ---
Me.idMaxLenth = 30 '// no longer than 30 characters.. never tested.
End Sub
'
' "THIS FIND EVERYTHING IN QUOTATIONS"
'
Sub SetClassToFindAllBlocks()
' set header and footer block varibles
' these determine what is to be found!
Me.szBlockStart = Chr$(34)
Me.szBlockEnd = Chr$(34) ' you could add [& ">"] ..
' ---- this has never been tested ---
Me.idMaxLenth = 30 '// no longer than 30 characters.. never tested.
End Sub
'
' <A HREF="FINDS ALL LINKS"></A>
'
Sub SetClassToFindLinks()
'set header and footer block varibles
' This determine what is to be found
Me.szBlockStart = "<a href=" & Chr$(34)
Me.szBlockEnd = Chr$(34) ' you could add [& ">"]
' --- this has never been tested ---
Me.idMaxLenth = 30 ' no longer than 30 character
End Sub
Sub SetClassToFindScr()
'set header and footer block varibles
' This determine what is to be found
Me.szBlockStart = "scr=" & Chr$(34)
Me.szBlockEnd = Chr$(34) ' you could add [& ">"]
' --- this has never been tested ---
Me.idMaxLenth = 30 ' no longer than 30 character
End Sub
'
'
'
' strips document of all links or emails or
' whatever is set to find..
Sub StripDocument(DocumentData As Variant, Append As Boolean)
Dim CurrentOffset As Long
Dim ChStart As Long
Dim ChEnd As Long

' new append functions for links and images
If Append = False Then
    CurrentIndex = 0
    Else
    CurrentIndex = CurrentIndex + 1
    End If
    
CurrentOffset = 1
Do
    ' find header of block (get byte position)
    ChStart = InStr(CurrentOffset, DocumentData, szBlockStart)
    If ChStart > 0 Then '(**) if not header found exit do
    ' get footer of block (get byte position)
    ChEnd = InStr((ChStart + Len(szBlockStart)), DocumentData, szBlockEnd)
    If ChEnd > 0 Then ' (**)if no end found exit do
    ' check maxlenth
    If ChStart - ChEnd > idMaxLenth Then ' void
    CurrentOffset = ChEnd ' move offset up.
        Else
        ' extract block from document                      ^ You have to subtrace the ChStart From ChEnd Since it says **lenth** no staopping byte!
    Stripped(CurrentIndex) = Mid(DocumentData, (ChStart), (ChEnd - ChStart))
    
    ' ******** custome configuration for correcting the link return *******
    Stripped(CurrentIndex) = Mid(Stripped(CurrentIndex), InStr(1, Stripped(CurrentIndex), Chr$(34)) + 1, Len(Stripped(CurrentIndex)))
    ' *********************************************************************
    
    ' move the index up for the next array position
    CurrentIndex = CurrentIndex + 1
    ' put offset above found email - or it will find it agin!
    CurrentOffset = ChEnd
    '
    '
    ' Easy Debug Way! You can remove this it is not needed.
    ' Its just for demonstration and testing. Use g and p functions
    ' for accessing the array outside the function.
    'MsgBox Stripped(CurrentIndex - 1)
    '
    '
    '
    End If
    Else
    Exit Do ' can not find another block
    End If
    
    Else
    Exit Do ' cant find block footer (found header)
    End If
DoEvents ' no dead locks (just regular locks)!
' failsafe  guess.
Loop Until CurrentOffset > Len(DocumentData)
End Sub

