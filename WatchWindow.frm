VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STOP WATCH"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "WatchWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "PRINT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   18
      ToolTipText     =   "Print Timings"
      Top             =   1200
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   3600
      Picture         =   "WatchWindow.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   510
      TabIndex        =   17
      ToolTipText     =   "Click to send in background (system tray)."
      Top             =   2880
      Width           =   570
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   3600
      Picture         =   "WatchWindow.frx":0884
      ScaleHeight     =   495
      ScaleWidth      =   510
      TabIndex        =   15
      ToolTipText     =   "Click to send in background (system tray)."
      Top             =   2880
      Width           =   570
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      ToolTipText     =   "Help"
      Top             =   1680
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "WatchWindow.frx":0CC6
      Left            =   1800
      List            =   "WatchWindow.frx":0CC8
      Style           =   2  'Dropdown List
      TabIndex        =   10
      ToolTipText     =   "Select a Lap to see its timings.No Lap timing available now."
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SPLIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Split Laps. Lap 1 in progress."
      Top             =   1200
      Width           =   975
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3240
      Top             =   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      ToolTipText     =   "About"
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      ToolTipText     =   "EXIT/RESET"
      Top             =   1200
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3600
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      ToolTipText     =   "START/STOP"
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Previous Lap's Times -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   19
      ToolTipText     =   "Previous Lap's Time are shown below."
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   0
      Picture         =   "WatchWindow.frx":0CCA
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   3840
      Picture         =   "WatchWindow.frx":110C
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   3840
      Picture         =   "WatchWindow.frx":154E
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "CHRONO"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   16
      ToolTipText     =   "STOP WATCH "
      Top             =   0
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "WatchWindow.frx":1990
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   14
      ToolTipText     =   "Current Lap Number"
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Lap End Time"
      Top             =   2520
      Width           =   4095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "LAST 10 LAPS :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Lap Time"
      Top             =   2160
      Width           =   4095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   1440
      TabIndex        =   7
      ToolTipText     =   "Minutes Elapsed"
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   3720
      TabIndex        =   3
      ToolTipText     =   "100 th of a Second Elapsed"
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      ToolTipText     =   "Seconds Elapsed"
      Top             =   600
      Width           =   495
   End
   Begin VB.Menu mnushell 
      Caption         =   "Menu"
      Enabled         =   0   'False
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu mnustart 
         Caption         =   "Start"
      End
      Begin VB.Menu mnustop 
         Caption         =   "Stop"
         Visible         =   0   'False
      End
      Begin VB.Menu mnureset 
         Caption         =   "Reset"
         Visible         =   0   'False
      End
      Begin VB.Menu mnusplit 
         Caption         =   "Split"
         Visible         =   0   'False
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnurestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnutop 
         Caption         =   "Always on top"
      End
      Begin VB.Menu mnutray 
         Caption         =   "System Tray"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim minn, secn, msecn, cx, cy, i, j, k, m, p, tp, flag, ttmin(10), ttsec(10), ttmsec(10), tlmin(10), tlsec(10), tlmsec(10) As Integer
Dim min, sec, msec, tmin(10), tsec(10), tmsec(10), lmin(10), lsec(10), lmsec(10), prt1(1000), prt2(1000) As String

Private Sub Combo1_Click()
k = Combo1.Text
l = k \ 10
l = k - (l * 10)
Select Case l
Case "1"
Label6.Caption = "Lap " & k & " Time        = " & lmin(1) & " : " & lsec(1) & " : " & lmsec(1)
Label8.Caption = "Lap " & k & " Ended at = " & tmin(1) & " : " & tsec(1) & " : " & tmsec(1)
Case "2"
Label6.Caption = "Lap " & k & " Time        = " & lmin(2) & " : " & lsec(2) & " : " & lmsec(2)
Label8.Caption = "Lap " & k & " Ended at = " & tmin(2) & " : " & tsec(2) & " : " & tmsec(2)
Case "3"
Label6.Caption = "Lap " & k & " Time        = " & lmin(3) & " : " & lsec(3) & " : " & lmsec(3)
Label8.Caption = "Lap " & k & " Ended at = " & tmin(3) & " : " & tsec(3) & " : " & tmsec(3)
Case "4"
Label6.Caption = "Lap " & k & " Time        = " & lmin(4) & " : " & lsec(4) & " : " & lmsec(4)
Label8.Caption = "Lap " & k & " Ended at = " & tmin(4) & " : " & tsec(4) & " : " & tmsec(4)
Case "5"
Label6.Caption = "Lap " & k & " Time        = " & lmin(5) & " : " & lsec(5) & " : " & lmsec(5)
Label8.Caption = "Lap " & k & " Ended at = " & tmin(5) & " : " & tsec(5) & " : " & tmsec(5)
Case "6"
Label6.Caption = "Lap " & k & " Time        = " & lmin(6) & " : " & lsec(6) & " : " & lmsec(6)
Label8.Caption = "Lap " & k & " Ended at = " & tmin(6) & " : " & tsec(6) & " : " & tmsec(6)
Case "7"
Label6.Caption = "Lap " & k & " Time        = " & lmin(7) & " : " & lsec(7) & " : " & lmsec(7)
Label8.Caption = "Lap " & k & " Ended at = " & tmin(7) & " : " & tsec(7) & " : " & tmsec(7)
Case "8"
Label6.Caption = "Lap " & k & " Time        = " & lmin(8) & " : " & lsec(8) & " : " & lmsec(8)
Label8.Caption = "Lap " & k & " Ended at = " & tmin(8) & " : " & tsec(8) & " : " & tmsec(8)
Case "9"
Label6.Caption = "Lap " & k & " Time        = " & lmin(9) & " : " & lsec(9) & " : " & lmsec(9)
Label8.Caption = "Lap " & k & " Ended at = " & tmin(9) & " : " & tsec(9) & " : " & tmsec(9)
Case "0"
Label6.Caption = "Lap " & k & " Time        = " & lmin(10) & " : " & lsec(10) & " : " & lmsec(10)
Label8.Caption = "Lap " & k & " Ended at = " & tmin(10) & " : " & tsec(10) & " : " & tmsec(10)
End Select
End Sub

Private Sub Command1_Click()
If Command1.Caption = "STOP" Then

Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Command3.Enabled = True
Command3.Visible = True
Command4.Enabled = False
Label9.Enabled = False
Label9.Visible = False
Command5.Enabled = True
Command5.Visible = True
Command1.Caption = "START"
Command2.Caption = "RESET"
mnustart.Visible = True
mnustop.Visible = False
mnuexit.Visible = False
mnureset.Visible = True
mnusplit.Visible = False

If (j >= 1000) Then
MsgBox "You can not run this stop watch for more than 1000 laps." & vbCrLf & "Reset to run again.", vbOKOnly + vbCritical, "STOP WATCH - Error"
Else
Call Command4_Click
End If

Else

If (j >= 1000) Then
MsgBox "You can not run this stop watch for more than 1000 laps." & vbCrLf & "Reset to run again.", vbOKOnly + vbCritical, "STOP WATCH - Error"
Else
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
Command3.Enabled = False
Command3.Visible = False
Command4.Enabled = True
Label9.Enabled = True
Label9.Visible = True
Command5.Enabled = False
Command5.Visible = False
Command1.Caption = "STOP"
Command2.Caption = "EXIT"
mnustop.Visible = True
mnustart.Visible = False
mnuexit.Visible = True
mnureset.Visible = False
mnusplit.Visible = True
End If
End If

Call systemtray

End Sub

Private Sub Command2_Click()
If Command2.Caption = "EXIT" Then
Shell_NotifyIcon NIM_DELETE, NotifyIcon
Unload Form1
MsgBox "Thanks for using STOP WATCH Ver 1.0" & vbCrLf & "Made by PARMENDER DAHIYA." & vbCrLf & "For any Query or bug reporting mail at ps_dahiya@yahoo.com.", vbOKOnly + vbInformation, "STOP WATCH - Thanks"
End
End If
If Command2.Caption = "RESET" Then
Label1.Caption = "00"
Label3.Caption = "00"
Label5.Caption = "00"
Label6.Caption = ""
Label6.Visible = False
Label8.Caption = ""
Label8.Visible = False
Label9.Caption = "LAP 1"
Combo1.ToolTipText = "Select a Lap to see its timings.No Lap timing available now."
mnureset.Visible = False
mnuexit.Visible = True
secn = 0
msecn = 0
minn = 0
For i = 1 To 10
ttmin(i) = 0
ttsec(i) = 0
ttmsec(i) = 0
tlmin(i) = 0
tlsec(i) = 0
tlmsec(i) = 0
Next i
i = 0

For p = 1 To j
prt1(p) = ""
prt2(p) = ""
Next p

j = 0
Command2.Caption = "EXIT"
Combo1.Clear
End If
flag = 1
Call systemtray

End Sub

Private Sub Command3_Click()
MsgBox "Stop Watch - Beta Version 1.0" & vbCrLf & "Copyright - Parmender Dahiya" & vbCrLf & "TATA CONSULTANCY SERVICES" & vbCrLf & "www.tcs.com" & vbCrLf & "ps_dahiya@yahoo.com", vbOKOnly + vbInformation, "STOP WATCH - About"
End Sub

Private Sub Command4_Click()

If (j >= 1000) Then
MsgBox "You can not run this stop watch for more than 1000 laps." & vbCrLf & "Reset to run again.", vbOKOnly + vbCritical, "STOP WATCH - Error"
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Command3.Enabled = True
Command3.Visible = True
Command4.Enabled = False
Label9.Enabled = False
Label9.Visible = False
Command5.Enabled = True
Command5.Visible = True
Command1.Caption = "START"
Command2.Caption = "RESET"
mnustart.Visible = True
mnustop.Visible = False
mnuexit.Visible = False
mnureset.Visible = True
mnusplit.Visible = False

Else

i = i + 1
j = j + 1
If i > 10 Then
i = 1
End If

tmin(i) = min
tsec(i) = sec
tmsec(i) = msec

ttmin(i) = minn
ttsec(i) = secn
ttmsec(i) = msecn

If i = 1 Then
tlmin(i) = ttmin(i) - ttmin(10)
tlsec(i) = ttsec(i) - ttsec(10)
tlmsec(i) = ttmsec(i) - ttmsec(10)
Else
tlmin(i) = ttmin(i) - ttmin(i - 1)
tlsec(i) = ttsec(i) - ttsec(i - 1)
tlmsec(i) = ttmsec(i) - ttmsec(i - 1)
End If

If tlmsec(i) < 0 Then
tlsec(i) = tlsec(i) - 1
tlmsec(i) = 100 + tlmsec(i)
End If
If tlsec(i) < 0 Then
If tlmin(i) < 0 Then
tlmin(i) = 0
Else
tlmin(i) = tlmin(i) - 1
End If
tlsec(i) = 60 + tlsec(i)
End If
lmin(i) = tlmin(i)
lsec(i) = tlsec(i)
lmsec(i) = tlmsec(i)

Select Case tlmin(i)
Case 0
lmin(i) = "00"
Case 1
lmin(i) = "01"
Case 2
lmin(i) = "02"
Case 3
lmin(i) = "03"
Case 4
lmin(i) = "04"
Case 5
lmin(i) = "05"
Case 6
lmin(i) = "06"
Case 7
lmin(i) = "07"
Case 8
lmin(i) = "08"
Case 9
lmin(i) = "09"
End Select

Select Case tlsec(i)
Case 0
lsec(i) = "00"
Case 1
lsec(i) = "01"
Case 2
lsec(i) = "02"
Case 3
lsec(i) = "03"
Case 4
lsec(i) = "04"
Case 5
lsec(i) = "05"
Case 6
lsec(i) = "06"
Case 7
lsec(i) = "07"
Case 8
lsec(i) = "08"
Case 9
lsec(i) = "09"
End Select

Select Case tlmsec(i)
Case 0
lmsec(i) = "00"
Case 1
lmsec(i) = "01"
Case 2
lmsec(i) = "02"
Case 3
lmsec(i) = "03"
Case 4
lmsec(i) = "04"
Case 5
lmsec(i) = "05"
Case 6
lmsec(i) = "06"
Case 7
lmsec(i) = "07"
Case 8
lmsec(i) = "08"
Case 9
lmsec(i) = "09"
End Select

Label6.Visible = True
Label8.Visible = True
Label6.Caption = "Last Lap Time        = " & lmin(i) & " : " & lsec(i) & " : " & lmsec(i)
Label8.Caption = "Last Lap Ended at = " & tmin(i) & " : " & tsec(i) & " : " & tmsec(i)
If Combo1.ListCount = 10 Then
Combo1.RemoveItem 0
End If
Combo1.AddItem j
Combo1.ToolTipText = "Select a Lap out of last 10. " & j & " lap(s) elapsed."
Command4.ToolTipText = "Split Laps. Lap " & (j + 1) & " in progress."
Label9.Caption = "LAP " & (j + 1)
prt1(j) = "Lap " & j & " Time = " & lmin(i) & " : " & lsec(i) & " : " & lmsec(i)
prt2(j) = "Lap " & j & " Ended at = " & tmin(i) & " : " & tsec(i) & " : " & tmsec(i)
Call systemtray
End If

End Sub

Private Sub Command5_Click()
MsgBox "Operational Help :" & vbCrLf & "1. To Start or Stop the STOP WATCH click on START/STOP Button (Whichever appears)." & vbCrLf & "2. To RESET or EXIT click on RESET/EXIT button (Whichever appears)." & vbCrLf & "3. In the DropDown list below you can see upto last 10 Laps." & vbCrLf & "4. To SPLIT Laps, click on SPLIT button.This button will be enabled only when watch is running." & vbCrLf & "5. To print the timings of all the laps elapsed click on the PRINT button." & vbCrLf & "6. To put the STOP WATCH in system tray click on the picture on bottom right." & vbCrLf & "7. To see the description of a Button, Text or List rest the mouse there." & vbCrLf & "Limitations : " & vbCrLf & "1. You can run this STOP WATCH for 1000 laps only." & vbCrLf & "2. When minutes goes above the four digits then STOP WATCH will be working but you may not be able to see them pro" & vbCrLf & "For further help or bug reporting you can mail to PARMENDER DAHIYA at ps_dahiya@yahoo.com", vbOKOnly + vbInformation, "STOP WATCH - Help"
End Sub

Private Sub Command6_Click()
If Combo1.ListCount = 0 Then
MsgBox "No data available for printing.", vbOKOnly + vbCritical, "STOP WATCH - Error"
Exit Sub
End If
Printer.ScaleMode = 3
Printer.FontSize = 18
Printer.FontBold = True
Printer.FontUnderline = True
Printer.FontName = "Arial"
Printer.Print
Printer.Print Tab(18); "STOP WATCH LAP TIMINGS";
Printer.Print
Printer.Print
Printer.FontSize = 10
Printer.FontBold = False
Printer.FontUnderline = False
For p = 1 To j
tp = p \ 30
tp = (p - tp * 30)
Printer.Print
Printer.Print Tab(15); prt1(p) & "             " & prt2(p)

If tp = 0 Then
Printer.FontSize = 8
Printer.CurrentX = 2000
Printer.CurrentY = 6500
Printer.Print "Page " & Printer.Page
Printer.CurrentX = 3200
Printer.CurrentY = 6500
Printer.Print "STOP WATCH made by :Parmender Dahiya"
Printer.FontSize = 10
Printer.NewPage
Printer.Print
End If

Next p

Printer.Print
cx = Printer.CurrentX
cy = Printer.CurrentY
Printer.Line -(4000, cy)
Printer.Print Tab(15); "TOTAL LAPS :    " & j
Printer.Print

cx = Printer.CurrentX
cy = Printer.CurrentY
Printer.Line -(4000, cy)
Printer.Print Tab(50); "END OF REPORT"
Printer.Print
cx = Printer.CurrentX
cy = Printer.CurrentY
Printer.Line -(4000, cy)

Printer.FontSize = 8
Printer.CurrentX = 2000
Printer.CurrentY = 6500
Printer.Print "Page " & Printer.Page
Printer.CurrentX = 3200
Printer.CurrentY = 6500
Printer.Print "STOP WATCH made by :Parmender Dahiya"
Printer.EndDoc
MsgBox "Report has been sent to printer.", vbOKOnly + vbInformation, "STOP WATCH - Printing"
End Sub

Private Sub Form_Load()
Label1.Caption = "00"
Label3.Caption = "00"
Label5.Caption = "00"
Label6.Caption = ""
Label8.Caption = ""
Label6.Visible = False
Label8.Visible = False
Command4.Enabled = False
Label9.Enabled = False
Label9.Visible = False
Label9.Caption = "LAP 1"
secn = 0
msecn = 0
minn = 0
msec = "00"
sec = "00"
min = "00"
i = 0
j = 0
Combo1.Clear
End Sub
Private Sub mnuexit_Click()
Call Command2_Click
End Sub

Private Sub mnureset_Click()
Call Command2_Click
End Sub

Private Sub mnurestore_Click()
mnushell.Visible = False
mnushell.Enabled = False
Form1.Show
End Sub
Private Sub mnusplit_Click()
Call Command4_Click
End Sub

Private Sub mnustart_Click()
Call Command1_Click
End Sub

Private Sub mnustop_Click()
Call Command1_Click
End Sub

Private Sub OLE1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub

Private Sub OLE1_Updated(Code As Integer)

End Sub

Private Sub mnutop_Click()
If mnutop.Checked = False Then
SetWindowPos Form1.hWnd, -1, 0, 0, 0, 0, &H1 Or &H2
mnutop.Checked = True
Call mnurestore_Click
Else
SetWindowPos Form1.hWnd, -2, 0, 0, 0, 0, &H1 Or &H2
mnutop.Checked = False
End If

End Sub

Private Sub mnutray_Click()
Call Picture1_Click
End Sub

Private Sub Picture1_Click()
Call systemtray
mnushell.Visible = True
mnushell.Enabled = True
Form1.Hide
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Hex(x) = "1E3C" Then
Form1.PopupMenu Form1.mnushell
End If

If Hex(x) = "1E0F" Then
mnushell.Visible = False
mnushell.Enabled = False

Form1.Show
End If
End Sub

Private Sub Timer1_Timer()
msecn = msecn + 1
msec = msecn
Select Case msecn
Case 0
msec = "00"
Case 1
msec = "01"
Case 2
msec = "02"
Case 3
msec = "03"
Case 4
msec = "04"
Case 5
msec = "05"
Case 6
msec = "06"
Case 7
msec = "07"
Case 8
msec = "08"
Case 9
msec = "09"
Case 100
msecn = 0
msec = "00"
Timer1.Enabled = False
Timer2.Enabled = False
Call Timer2_Timer
Timer1.Enabled = True
Timer2.Enabled = True
End Select

Label3.Caption = msec
End Sub

Private Sub Timer2_Timer()
msecn = 0
secn = secn + 1
sec = secn
Select Case secn
Case 0
sec = "00"
Case 1
sec = "01"
Case 2
sec = "02"
Case 3
sec = "03"
Case 4
sec = "04"
Case 5
sec = "05"
Case 6
sec = "06"
Case 7
sec = "07"
Case 8
sec = "08"
Case 9
sec = "09"
Case 60
secn = 0
sec = "00"
End Select
Label1.Caption = sec
m = secn \ 2
m = (secn - m * 2)
If m <> 1 Then
Image1.Visible = False
Image3.Visible = True
Image2.Visible = False
Image4.Visible = True
Else
Image3.Visible = False
Image1.Visible = True
Image4.Visible = False
Image2.Visible = True
End If
End Sub

Private Sub Timer3_Timer()
msecn = 0
secn = 0
minn = minn + 1
min = minn
Select Case minn
Case 0
min = "00"
Case 1
min = "01"
Case 2
min = "02"
Case 3
min = "03"
Case 4
min = "04"
Case 5
min = "05"
Case 6
min = "06"
Case 7
min = "07"
Case 8
min = "08"
Case 9
min = "09"
End Select
Label5.Caption = min
End Sub

Private Sub systemtray()
'We set up Picture1 to accept callback data
'in it's MouseMove procedure.

NotifyIcon.cbSize = Len(NotifyIcon)
NotifyIcon.hWnd = Picture1.hWnd
NotifyIcon.uID = 1&
NotifyIcon.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
NotifyIcon.uCallbackMessage = WM_MOUSEMOVE

'Now, we set up the icon and tool tip message
NotifyIcon.hIcon = Picture1.Picture
If Command1.Caption = "STOP" Then
NotifyIcon.szTip = "Stop Watch." & vbCrLf & "Status: RUNNING." & vbCrLf & "Lap(s) Elapsed: " & j & Chr$(0)
Else
NotifyIcon.szTip = "Stop Watch." & vbCrLf & "Status: STOPPED." & vbCrLf & "Lap(s) Elapsed: " & j & Chr$(0)
End If

'Lastly, we add the icon
Shell_NotifyIcon NIM_DELETE, NotifyIcon
Shell_NotifyIcon NIM_ADD, NotifyIcon

End Sub



