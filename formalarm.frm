VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmalarm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                         ALARM"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "formalarm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1920
      Picture         =   "formalarm.frx":0442
      ScaleHeight     =   330
      ScaleWidth      =   1425
      TabIndex        =   5
      Top             =   1320
      Width           =   1425
      Begin VB.Label Command1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   50
         Width           =   1455
      End
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5400
      Top             =   1200
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   4800
      Top             =   1320
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4200
      Top             =   1200
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3720
      Top             =   1200
   End
   Begin MCI.MMControl MM 
      Height          =   330
      Left            =   840
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4320
      Picture         =   "formalarm.frx":0AD0
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   4320
      Picture         =   "formalarm.frx":0F12
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lbldate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2280
      TabIndex        =   4
      Top             =   480
      Width           =   135
   End
   Begin VB.Label lbltime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "And On Time: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   225
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alarm Set On Date: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   225
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   1650
   End
End
Attribute VB_Name = "frmalarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Unload Me

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim VarSound, VarTime As String, VarDate As String, varbeep

VarDate = GetFromINI("Alarm", "Infodate", App.Path & "\Hoursdata.dat")
VarTime = GetFromINI("Alarm", "Time", App.Path & "\Hoursdata.dat")
VarSound = GetFromINI("Alarm", "Soundfile", App.Path & "\Hoursdata.dat")
varbeep = GetFromINI("Alarm", "beep", App.Path & "\Hoursdata.dat")

lbldate.Caption = VarDate
lbltime.Caption = VarTime
   
   If varbeep = 0 Then
    MM.Command = "close"
    MM.FileName = VarSound
    MM.Command = "open"
    MM.Command = "play"
    Else
    Timer4.Enabled = True
    End If
    
End Sub

Private Sub Timer1_Timer()
Image1.Visible = False
Image2.Visible = True

Timer2.Enabled = True
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Image1.Visible = True
Image2.Visible = False

Timer2.Enabled = False
Timer1.Enabled = True
End Sub

Private Sub Timer3_Timer()
Dim nReturnValue As Long
    nReturnValue = FlashWindow(frmalarm.hWnd, True)
    
End Sub

Private Sub Timer4_Timer()
Beep
End Sub
