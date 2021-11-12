VERSION 5.00
Begin VB.Form frmNanny 
   Caption         =   "IENanny"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "ienanny.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   3720
      TabIndex        =   7
      Top             =   3240
      Width           =   2172
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete From List"
      Height          =   492
      Left            =   3720
      TabIndex        =   6
      Top             =   3720
      Width           =   2172
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   0
      TabIndex        =   4
      Top             =   3240
      Width           =   2052
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Site To Block"
      Height          =   492
      Left            =   0
      TabIndex        =   3
      Top             =   3720
      Width           =   2172
   End
   Begin VB.ListBox lstwindows4 
      Height          =   2205
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   5892
   End
   Begin VB.ListBox lstWindows2 
      Height          =   450
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.ListBox lstWindows 
      Height          =   645
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   4200
      Top             =   2160
   End
   Begin VB.Label Label4 
      Caption         =   "Press F10 to unhide this window"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   9
      Top             =   4800
      Width           =   5172
   End
   Begin VB.Label Label3 
      Caption         =   "Press F9 to hide this window"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   360
      TabIndex        =   8
      Top             =   4440
      Width           =   5172
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Blocked Sites"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1440
      TabIndex        =   5
      Top             =   480
      Width           =   2892
   End
End
Attribute VB_Name = "frmNanny"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim bet As Long
 Dim met As Long

Private Sub Command1_Click()
lstwindows4.AddItem Text1.Text
Text1.Text = ""
End Sub

Private Sub Command2_Click()

ty = lstwindows4.ListCount
For ber = 0 To ty - 1
lstwindows4.ListIndex = ber
If lstwindows4.Text = Text2.Text Then
lstwindows4.RemoveItem ber
Text2.Text = ""
Exit Sub
End If
Next ber
Text2.Text = ""
End Sub

Private Sub Timer1_Timer()
If GetAsyncKeyState(VK_F9) Then
pp = ShowWindow(Me.hwnd, SW_HIDE)
End If
If GetAsyncKeyState(VK_F10) Then
pp = ShowWindow(Me.hwnd, SW_NORMAL)
End If
On Error Resume Next
If GetAsyncKeyState(VK_F9) Then
pp = ShowWindow(Me.hwnd, SW_HIDE)
End If
If GetAsyncKeyState(VK_F10) Then
pp = ShowWindow(Me.hwnd, SW_NORMAL)
End If
Dim hDesktop        As Long
Dim mtr As Long
Dim stext As String * 100
        Dim b As Long
        Dim x As Long
        Dim t As Long
        Dim w As Long
        Dim der As Long
 Dim slength2 As Long

         hDesktop = GetDesktopWindow()
        Call ListChildWindows(lstWindows, hDesktop)
        b = lstWindows.ListCount
        For t = 0 To b - 1
        lstWindows.ListIndex = t
        x = InStr(lstWindows.Text, "Microsoft Internet Explorer")
  
  If x <> 0 Then
  w = InStr(lstWindows.Text, "-")
  
  der = Left(lstWindows.Text, w - 1)
     Call ListChildWindows(lstWindows2, der)
   mp = lstWindows2.ListCount
   For er = 1 To mp
   lstWindows2.ListIndex = er
w = InStr(lstWindows2.Text, "-")
  der = Left(lstWindows2.Text, w - 1)
  dear = GetClassName(der, stext, 100)
  cutdear = Left(stext, dear)
  xx = InStr(cutdear, "Edit")
  If xx <> 0 Then
  Dim wintext As String
Dim slength As Long
Dim retval As Long
Dim tryval As Long
  Dim max As Boolean
  slength = SendMessage(der, WM_GETTEXTLENGTH, ByVal CLng(0), ByVal CLng(0)) + 1
wintext = Space(slength)
Dim ert As Long
ert = slength
retval = SendMessage(der, WM_GETTEXT, ByVal slength, ByVal wintext)
wer = Len(wintext)
wintext = Left(wintext, wer - 1)

rare = lstwindows4.ListCount
If rare = 0 Then Exit Sub
For ytr = 0 To rare
lstwindows4.ListIndex = ytr
check = lstwindows4.Text
If check = "" Then Exit Sub
swq = InStr(wintext, check)
If swq <> 0 Then
slength2 = 35
Dim mess As String
mess = App.Path & "\" & "nope.htm"
qwert = SendMessage(der, WM_SETTEXT, ByVal slength2, ByVal mess)
Call keybd_event(VK_RETURN, 1, 0, 0)
End If
Next ytr
   End If
    Next er
    End If
    Next t
End Sub
