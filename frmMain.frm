VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4635
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":08CA
   ScaleHeight     =   4260
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   390
      Top             =   4005
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   150
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer timMain 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   3885
   End
   Begin VB.Label lblAbout 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed By:- Keral.C.Patel.      Email:-keral82@keral.com"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   150
      Left            =   4545
      TabIndex        =   13
      Top             =   4080
      Width           =   3165
   End
   Begin VB.Label cmdS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Start Now"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   1710
      TabIndex        =   12
      Top             =   3735
      Width           =   1020
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   1425
      Top             =   3540
      Width           =   1530
   End
   Begin VB.Label txtMain 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   11
      Top             =   3255
      Width           =   45
   End
   Begin VB.Label lblMain 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Window DC        :"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   4
      Left            =   195
      TabIndex        =   10
      Top             =   3000
      Width           =   1125
   End
   Begin VB.Label txtMain 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   165
      TabIndex        =   9
      Top             =   2655
      Width           =   45
   End
   Begin VB.Label txtMain 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   165
      TabIndex        =   8
      Top             =   2055
      Width           =   45
   End
   Begin VB.Label txtMain 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   165
      TabIndex        =   7
      Top             =   1455
      Width           =   45
   End
   Begin VB.Label txtMain 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   6
      Top             =   855
      Width           =   45
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   4365
      TabIndex        =   5
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Window SPY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   1620
      TabIndex        =   4
      Top             =   75
      Width           =   1365
   End
   Begin VB.Label lblMain 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Window Handle  :"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   3
      Top             =   600
      Width           =   1275
   End
   Begin VB.Label lblMain 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Window Caption :"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   1
      Left            =   165
      TabIndex        =   2
      Top             =   1200
      Width           =   1260
   End
   Begin VB.Label lblMain 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Window Parent   :"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   2
      Left            =   165
      TabIndex        =   1
      Top             =   1800
      Width           =   1275
   End
   Begin VB.Label lblMain 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Window Class     :"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   3
      Left            =   165
      TabIndex        =   0
      Top             =   2400
      Width           =   1275
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
       Dim s As Integer
       Dim dta As String
Const HTCAPTION = 2
Const WM_NCLBUTTONDOWN = &HA1
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long


Private Declare Function ShellExecute Lib "shell32.dll" Alias _
       "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As _
       String, ByVal lpFile As String, ByVal lpParameters As String, _
       ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public LastState As Integer
Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const SC_RESTORE = &HF120&
Private Const SC_SIZE = &HF000&
Public Sub TrayIconCallback(Msg As Long)
   If Msg = WM_LBUTTONDBLCLK Then
      Me.Visible = True
      Me.WindowState = vbNormal
   End If
End Sub
Private Sub mnuTray_Click()
If frmMain.Visible = "false" Then
frmMain.Visible = "true"
Else
frmMain.Visible = "false"
frmMain.Visible = "true"
End If
End Sub
Private Sub cmdS_Click()
'Just for enabling and disabling the timer
If cmdS.Caption = "Start Now" Then
    timMain.Enabled = True
    cmdS.Caption = "Stop Now"
Else
    cmdS.Caption = "Start Now"
    Screen.MousePointer = vbDefault
    timMain.Enabled = False
    For i = 0 To 4
    txtMain(i).Caption = ""
    Next
End If
End Sub

Private Sub cmdS_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
cmdS.ForeColor = vbRed
End Sub

Private Sub Form_DblClick()
    On Error GoTo fileOpenErrr
       CDialog.CancelError = True
       CDialog.Flags = &H4& Or &H100&
       CDialog.DefaultExt = ".jpg"
       CDialog.DialogTitle = "Select File To Open"
       CDialog.Filter = "JPEG (*.jpg)|*.jpg|GIF (*.gif)|*.gif|BITMAP (*.bmp)|*.bmp"
       CDialog.ShowOpen
Set frmMain.Picture = LoadPicture(CDialog.FileName)
fileOpenErrr:
       Exit Sub
End Sub

Private Sub Form_Load()
PutWindowOnTop frmMain
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'for movaable form
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub


Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
cmdS.ForeColor = &H404040
End Sub

Private Sub Label3_Click()
Unload Me
End Sub

Private Sub Label4_Click()
frmMain.Hide
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'for movaable form
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub



Private Sub Timer1_Timer()
If lblAbout.Left >= -3550 Then
lblAbout.Left = lblAbout.Left - 100
Else
lblAbout.Left = 4545
End If
End Sub

Private Sub timMain_Timer()

Dim P As POINTAPI

Dim hWn As Long

Dim WinCap As String * 255
Dim ClName As String * 255

Dim OldParent As Long, Parent As Long

'First, get the cursor position of mouse
GetCursorPos P


'WindowFromPoint returns the handle of the window under the mouse

hWn = WindowFromPoint(P.x, P.Y)
txtMain(0).Caption = hWn


'Determine the caption, using the handle we obtained above

GetWindowText hWn, WinCap, 254
txtMain(1).Caption = WinCap
If Trim(txtMain(1).Caption) = "" Then txtMain(1).Caption = "-------"


'Find the parent using the GetParent function. The loop is for
'detecting the Zero-th level parent of our window


Parent = GetParent(hWn)
Do While Parent
OldParent = Parent
Parent = GetParent(OldParent)
Loop
If Parent Then OldParent = Parent
GetWindowText OldParent, WinCap, 254
txtMain(2).Caption = WinCap
If Trim(txtMain(2).Caption) = "" Then txtMain(2).Caption = "-------"


'Get the class name of our window

GetClassName hWn, ClName, 254
txtMain(3).Caption = ClName
If Trim(txtMain(3).Caption) = "" Then txtMain(3).Caption = "-------"
   
   If SendMessage(hWn, EM_GETPASSWORDCHAR, 0, 1&) <> 0 Then
   SendMessage hWn, EM_SETPASSWORDCHAR, 0, 1&
   SendMessage hWn, EM_SETMODIFY, True, 1&
   End If


'get DC
txtMain(4).Caption = GetDC(hWn)
If Trim(txtMain(4).Caption) = "" Then txtMain(4).Caption = "-------"
End Sub

