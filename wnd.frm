VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Wnd 
   Caption         =   "Web Crawler 2.3"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   10950
   Icon            =   "wnd.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   10950
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tbar 
      Align           =   1  'Align Top
      Height          =   960
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   1693
      ButtonWidth     =   609
      ButtonHeight    =   1535
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton Command5 
         Caption         =   "Get  i.explorer url"
         Height          =   375
         Left            =   15
         TabIndex        =   11
         Top             =   495
         Width           =   2280
      End
      Begin VB.CommandButton Favorites 
         Caption         =   "Favorites"
         Height          =   375
         Left            =   4455
         TabIndex        =   10
         Top             =   30
         Width           =   1095
      End
      Begin VB.CommandButton History 
         Caption         =   "History"
         Height          =   375
         Left            =   3375
         TabIndex        =   9
         Top             =   30
         Width           =   1095
      End
      Begin VB.TextBox excludetxt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7665
         TabIndex        =   8
         Text            =   "exclude"
         Top             =   495
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Restart"
         Height          =   375
         Left            =   2295
         TabIndex        =   7
         Top             =   30
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Stop"
         Height          =   375
         Left            =   1215
         TabIndex        =   6
         Top             =   30
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Start"
         Height          =   375
         Left            =   15
         TabIndex        =   5
         Top             =   30
         Width           =   1215
      End
      Begin VB.TextBox aurl 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2295
         TabIndex        =   3
         Text            =   "http://www.hotbot.com/?MT=vb+source+code+free"
         Top             =   495
         Width           =   5310
      End
      Begin VB.TextBox amn 
         Height          =   300
         Left            =   9345
         TabIndex        =   2
         Text            =   "1000"
         Top             =   135
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Url to Search."
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   3375
      End
   End
   Begin MSComctlLib.TreeView view 
      Height          =   5790
      Left            =   0
      TabIndex        =   0
      Top             =   885
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   10213
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin InetCtlsObjects.Inet inet 
      Left            =   5280
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5280
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wnd.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wnd.frx":0460
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wnd.frx":2C14
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wnd.frx":53C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wnd.frx":7B7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wnd.frx":A330
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wnd.frx":A784
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wnd.frx":BA08
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wnd.frx":E1BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wnd.frx":EC8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wnd.frx":11440
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wnd.frx":13BF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wnd.frx":14048
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wnd.frx":167FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wnd.frx":1A018
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wnd.frx":1A46C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wnd.frx":1A8C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wnd.frx":1AD14
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Wnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
cmlstop = True
OODMODE = True
End Sub

Public Sub Command2_Click()
On Error Resume Next
Kill "c:\tempinet\*.*"
RmDir "c:\tempinet\"

cmlstop = False
OODMODE = False

Dim Stack As New W32LIB_Stack
Stack.Reset 0: Stack.Reset 1
Me.Caption = "Restart."  ' LastDoc(" & CurUrl & ") Nodes:" '& view.Nodes.Count

Lc2kEngine aurl, Wnd.view, Wnd.inet, amn, 1, Wnd
Me.Show
Me.Refresh

'Lc2kEngine MainFrm.aurl, Wnd.view, Wnd.inet, MainFrm.amn, 1, Wnd
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim urlx
urlx = GetAddressText
MsgBox urlx
aurl = urlx
Dim ret&
Dim costr As String
    'If MsgBox("Would you like to display this page?", vbYesNo, "Display Page Url") = vbYes Then
        'RunBrowserWindow BrowserWnd.Web, view.SelectedItem.Text, ""
        'BrowserWnd.Show
        'MsgBox view.SelectedItem.Text
    
        'costr = view.SelectedItem.Text
        'ret& = shellexecute(Me.hwnd, "Open", costr, "", "", 3)
    'End If
End Sub

Private Sub Command4_Click()
Lc2kEngine aurl, Wnd.view, Wnd.inet, amn, 1, Wnd
End Sub

Private Sub Command5_Click()
On Error GoTo CallErrorA
    'Dim usText() As Byte ' That's right, a byte array
    Dim iPos As Integer
    Dim sClassName As String
    Dim GetAddressText As String
    'Dim xc As Long
    Dim lhwnd As Long
    Dim WindowHandle As Long
    
    lhwnd = 0
   sClassName = ("IEFrame")
   lhwnd = FindWindowEx(lhwnd, 0, sClassName, vbNullString)
   sClassName = ("WorkerA")
   lhwnd = FindWindowEx(lhwnd, 0, sClassName, vbNullString)
   sClassName = ("ReBarWindow32")
   lhwnd = FindWindowEx(lhwnd, 0, sClassName, vbNullString)
   sClassName = ("ComboBoxEx32")
   lhwnd = FindWindowEx(lhwnd, 0, sClassName, vbNullString)
   sClassName = ("ComboBox")
   lhwnd = FindWindowEx(lhwnd, 0, sClassName, vbNullString)
   sClassName = ("Edit")
   lhwnd = FindWindowEx(lhwnd, 0, sClassName, vbNullString)
        
        WindowHandle& = lhwnd
        Dim buffer As String, TextLength As Long
        TextLength& = SendMessage(WindowHandle&, WM_GETTEXTLENGTH, 0&, 0&)
        buffer$ = String(TextLength&, 0&)
        Call SendMessageByString(WindowHandle&, WM_GETTEXT, TextLength& + 1, buffer$)
        'MsgBox buffer$
        aurl = buffer$
        


Exit_Command193_Click:
    Exit Sub

Err_Command193_Click:
    MsgBox Err.Description
    Resume Exit_Command193_Click
   Exit Sub
CallErrorA:
    MsgBox Err.Description
    Err.Clear

End Sub

Private Sub Favorites_Click()
On Error Resume Next

Dim costr As String
       
        costr = "c:\windows\favorites"
        
        ret& = shellexecute(Me.hwnd, "Open", costr, "", "", 3)
   
End Sub

Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Button = 2 Then PopupMenu mnuFile

End Sub
Private Sub Form_Resize()
view.Width = Wnd.ScaleWidth
view.Top = tbar.Height
view.Height = Wnd.ScaleHeight - tbar.Height

End Sub

Private Sub History_Click()
On Error Resume Next

Dim costr As String
       
        costr = "c:\windows\history"
        
        ret& = shellexecute(Me.hwnd, "Open", costr, "", "", 3)
   
End Sub

Private Sub view_DblClick()
    Dim ret&
    Dim costr As String
    'If MsgBox("Would you like to display this page?", vbYesNo, "Display Page Url") = vbYes Then
        'RunBrowserWindow BrowserWnd.Web, view.SelectedItem.Text, ""
        'BrowserWnd.Show
        'MsgBox view.SelectedItem.Text
    
        costr = view.SelectedItem.Text
        ret& = shellexecute(Me.hwnd, "Open", costr, "", "", 3)
    'End If
End Sub
Sub view_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'   If Button = 2 Then PopupMenu mnuFile

End Sub
Private Sub FindIt(ByVal sClassName As String)
On Error GoTo CallErrorA
If llhwnd = 0 Then llhwnd = lhwnd
    lhwnd = FindWindowEx(llhwnd, 0, sClassName, vbNullString)
    llhwnd = lhwnd
CallErrorA:
End Sub

Private Function GetAddressText() As String
On Error GoTo CallErrorA
    Dim usText() As Byte ' That's right, a byte array
    Dim iPos As Integer
    lhwnd = 0
    llhwnd = 0
    Call FindIt("IEFrame")
    'Call FindIt("WorkerA")
    'Call FindIt("ReBarWindow32")
    'Call FindIt("ComboBoxEx32")
    'Call FindIt("ComboBox")
    llhwnd = 960
    Call FindIt("Edit")
    ReDim usText(0 To SendMessage(lhwnd, WM_GETTEXTLENGTH, 0, ByVal 0&) + 1) ' +1 for Null Char
    If UBound(usText) = 1 Then
        GetAddressText = ""
    Else
        ' Length in first word:
        usText(0) = UBound(usText) And 255
        usText(1) = UBound(usText) \ 256
        Call SendMessage(lhwnd, WM_GETTEXT, UBound(usText), usText(0))
        ' Convert to string:
        GetAddressText = StrConv(usText, vbUnicode)
        ' Get rid of Null Char:
        iPos = InStr(GetAddressText, vbNullChar)
        If iPos > 0 Then GetAddressText = Left(GetAddressText, iPos - 1)
    End If
    
CallErrorA:
  
End Function

'To Use: use GetAddressText
