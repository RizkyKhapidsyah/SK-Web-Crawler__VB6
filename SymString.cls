VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FSD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub Store(Data, id)
Open "data.dat" For Binary As #1: Close #1
Open "data.dat" For Input As #1
Open "temp.dat" For Output As #2
Do
If EOF(1) Then Exit Do
Input #1, dat: DoEvents
    If Mid(dat, 1, InStr(1, dat, "#") - 1) = id Then
    b = "*"
    Print #2, id & "#" & CSV(Data)
    Else
    Print #2, dat
    End If
Loop Until EOF(1)
    If b <> "*" Then
    Print #2, id & "#" & CSV(Data)
    End If
    Close #1, #2
    FileCopy "temp.dat", "data.dat"
    'Kill "temp.dat"
End Sub
Function Recv(id)
Open "data.dat" For Binary As #1: Close #1
Open "data.dat" For Input As #1
Do
If EOF(1) Then Exit Do
Input #1, dat: DoEvents
    If Mid(dat, 1, InStr(1, dat, "#") - 1) = id Then
    Recv = CVS(Mid(dat, InStr(1, dat, "#") + 1, Len(dat)))
    End If
    Loop Until EOF(1)
    Close #1
End Function

Function CSV(xstring)
For X = 1 To Len(xstring)
CSV = CSV & AA(Asc(Mid(xstring, X, 1))): DoEvents
Next X
End Function
Function CVS(xstring)
For X = 1 To Len(xstring)
esv = Mid(xstring, X, 3)
CVS = CVS & Chr$(esv)
X = X + 2
Next X
End Function
Function AA(xs)
If Len(xs) = 3 Then AA = xs
If Len(xs) = 2 Then AA = "0" & xs
If Len(xs) = 1 Then AA = "00" & xs
If Len(xs) = 0 Then AA = "000"
End Function
Sub ShowIds(view As TreeView)
Open "data.dat" For Input As #1
kid = view.Nodes.Count
view.Nodes.Add , , "K" & kid, "DeviceStorageFile Listing"
Do
Input #1, Data: DoEvents
view.Nodes.Add "K" & kid, 4, view.Nodes.Count & Mid(Data, 1, InStr(1, Data, "#") - 1), Mid(Data, 1, InStr(1, Data, "#") - 1)
Loop Until EOF(1)
Close #1
End Sub

