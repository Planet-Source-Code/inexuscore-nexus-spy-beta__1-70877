VERSION 5.00
Begin VB.Form frm_about 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nexus Spy (BETA)"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4995
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_about.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "inexuscore@gmail.com"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   1560
      TabIndex        =   2
      Top             =   840
      Width           =   1860
   End
   Begin VB.Image imgHand 
      Height          =   480
      Left            =   240
      Picture         =   "frm_about.frx":7D42
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Programmed by INexusCore"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Index           =   2
      Left            =   1380
      TabIndex        =   1
      Top             =   600
      Width           =   2235
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nexus Spy (BETA)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   360
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   180
      Width           =   2355
   End
   Begin VB.Image Image1 
      Height          =   2685
      Left            =   1080
      Picture         =   "frm_about.frx":7E94
      Top             =   1200
      Width           =   2850
   End
End
Attribute VB_Name = "frm_about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'// private win32 apis declaration
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'// private constants declaration
Private Const SW_SHOWNORMAL = 1

Private Sub Form_KeyPress(KeyAscii As Integer)
    '// if Escape key or Enter key is pressed
    If KeyAscii = 13 Or KeyAscii = 27 Or KeyAscii = 32 Then
        Unload Me   '// unload about form
    End If
End Sub

Private Sub Form_Load()
    '// set a custom mouse icon for lblEmail control
    lblEmail.MouseIcon = imgHand.Picture
    lblEmail.MousePointer = 99
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '// update lblEmail forecolor
    lblEmail.ForeColor = &HC0&
End Sub

Private Sub lblEmail_Click()
    '// Call ShellExecute to run windows's OutlookExpress
    ShellExecute Me.hwnd, "open", "mailto:inexuscore@gmail.com", vbNullString, "C:\", SW_SHOWNORMAL
End Sub

Private Sub lblEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '// update lblEmail forecolor
    lblEmail.ForeColor = vbWhite
End Sub

Private Sub lblEmail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '// update lblEmail forecolor
    lblEmail.ForeColor = &HC0&
End Sub
