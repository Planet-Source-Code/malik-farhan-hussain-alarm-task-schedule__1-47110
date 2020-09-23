VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   " FANEE  ` s   HOURS"
   ClientHeight    =   3870
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   8145
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3824.472
   ScaleMode       =   0  'User
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Options"
      Height          =   3615
      Left            =   0
      TabIndex        =   50
      Top             =   120
      Visible         =   0   'False
      Width           =   7935
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3240
         Picture         =   "frmmain.frx":08FF
         ScaleHeight     =   330
         ScaleWidth      =   1425
         TabIndex        =   55
         Top             =   2520
         Width           =   1425
         Begin VB.Label Label31 
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
            TabIndex        =   56
            Top             =   50
            Width           =   1455
         End
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Start With Windows"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   52
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Minimize To Tray  On Startup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1320
         TabIndex        =   51
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Line Line24 
         X1              =   3480
         X2              =   4920
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OPTIONS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3600
         TabIndex        =   54
         Top             =   480
         Width           =   1230
      End
      Begin VB.Image Image16 
         Height          =   480
         Left            =   5760
         Picture         =   "frmmain.frx":0F8D
         Top             =   720
         Width           =   480
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   2895
         Left            =   720
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   480
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   48
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3480
      Top             =   5520
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   120
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   7440
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Left            =   7440
      Top             =   1080
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6840
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2295
      Begin VB.Line Line1 
         BorderColor     =   &H80000006&
         BorderWidth     =   3
         X1              =   720
         X2              =   1680
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000012&
         X1              =   600
         X2              =   1680
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000006&
         X1              =   720
         X2              =   1680
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   2310
         Left            =   0
         Picture         =   "frmmain.frx":1957
         Top             =   0
         Width           =   2280
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Task Schedule"
      Height          =   3015
      Left            =   120
      TabIndex        =   27
      Top             =   3840
      Visible         =   0   'False
      Width           =   7935
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   600
         TabIndex        =   33
         Top             =   615
         Width           =   255
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   28
         Text            =   "Choose Sound To Play"
         Top             =   1418
         Width           =   2655
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   840
         Top             =   1800
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   4560
         TabIndex        =   29
         ToolTipText     =   "Choose Time"
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   24444930
         CurrentDate     =   37816
         MaxDate         =   42369
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   4560
         TabIndex        =   30
         ToolTipText     =   "Chosse Gate"
         Top             =   600
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   24444928
         CurrentDate     =   37375
         MaxDate         =   401768
         MinDate         =   37257
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Stop Beep "
         Height          =   255
         Left            =   1800
         TabIndex        =   47
         Top             =   630
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lblbeep 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Test Beep "
         Height          =   255
         Left            =   1800
         TabIndex        =   46
         Top             =   630
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Line Line20 
         BorderWidth     =   2
         X1              =   7800
         X2              =   7800
         Y1              =   120
         Y2              =   2760
      End
      Begin VB.Line Line19 
         BorderWidth     =   2
         X1              =   120
         X2              =   7800
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Image Image11 
         Height          =   480
         Left            =   7200
         Picture         =   "frmmain.frx":5DA1
         Top             =   2280
         Width           =   480
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel Alarm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4440
         TabIndex        =   42
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set Alarm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2280
         TabIndex        =   41
         Top             =   2400
         Width           =   810
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Beep :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   960
         TabIndex        =   32
         Top             =   637
         Width           =   585
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Choose Sound To Play :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   40
         Top             =   1200
         Width           =   1725
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Choose Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4560
         TabIndex        =   39
         Top             =   1080
         Width           =   930
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Choose Day / Month / Date / Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4560
         TabIndex        =   38
         Top             =   360
         Width           =   2415
      End
      Begin VB.Image Image7 
         Height          =   480
         Left            =   3360
         Picture         =   "frmmain.frx":61E3
         Stretch         =   -1  'True
         ToolTipText     =   "Choose Sound To Play "
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image6 
         Height          =   480
         Left            =   3360
         Picture         =   "frmmain.frx":6EAD
         Stretch         =   -1  'True
         ToolTipText     =   "Task Running"
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Alarm Settings"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2280
         TabIndex        =   31
         Top             =   0
         Width           =   1395
      End
      Begin VB.Line Line23 
         BorderWidth     =   2
         X1              =   120
         X2              =   120
         Y1              =   120
         Y2              =   2760
      End
      Begin VB.Line Line22 
         BorderWidth     =   2
         X1              =   120
         X2              =   2160
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line21 
         BorderWidth     =   2
         X1              =   3840
         X2              =   7800
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   2400
         Picture         =   "frmmain.frx":7B77
         ToolTipText     =   "Set Alarm"
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   4680
         Picture         =   "frmmain.frx":8476
         ToolTipText     =   "Cancel Alarm"
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image Image10 
         Height          =   480
         Left            =   240
         Picture         =   "frmmain.frx":8D75
         Top             =   2280
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Task Schedule"
      Height          =   2895
      Left            =   120
      TabIndex        =   19
      Top             =   3840
      Visible         =   0   'False
      Width           =   7935
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   840
         Top             =   1800
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   24
         Text            =   "Choose Sound To Play"
         Top             =   1418
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   23
         Text            =   "Choose Program To perform Task"
         Top             =   578
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker picktasktime 
         Height          =   375
         Left            =   4560
         TabIndex        =   20
         ToolTipText     =   "Choose Time"
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   24444930
         CurrentDate     =   37816
         MaxDate         =   42260
      End
      Begin MSComCtl2.DTPicker picktaskdate 
         Height          =   285
         Left            =   4560
         TabIndex        =   21
         ToolTipText     =   "Choose Date"
         Top             =   600
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   24444928
         CurrentDate     =   37375
         MaxDate         =   401768
         MinDate         =   37257
      End
      Begin VB.Image Image13 
         Height          =   480
         Left            =   240
         Picture         =   "frmmain.frx":91B7
         Top             =   2160
         Width           =   480
      End
      Begin VB.Image Image12 
         Height          =   480
         Left            =   7200
         Picture         =   "frmmain.frx":95F9
         Top             =   2160
         Width           =   480
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel Task"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4680
         TabIndex        =   44
         Top             =   2400
         Width           =   1005
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set Task"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2280
         TabIndex        =   43
         Top             =   2400
         Width           =   720
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Choose Sound To Play :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   37
         Top             =   1200
         Width           =   1725
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Choose Task File :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   36
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Choose Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4560
         TabIndex        =   35
         Top             =   1080
         Width           =   930
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Choose Day / Month / Date / Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4560
         TabIndex        =   34
         Top             =   360
         Width           =   2415
      End
      Begin VB.Image canceltask 
         Height          =   480
         Left            =   4800
         Picture         =   "frmmain.frx":9A3B
         ToolTipText     =   "Cancel Task"
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image settask 
         Height          =   480
         Left            =   2400
         Picture         =   "frmmain.frx":A33A
         ToolTipText     =   "Set Task"
         Top             =   1920
         Width           =   480
      End
      Begin VB.Line Line18 
         BorderWidth     =   2
         X1              =   120
         X2              =   7800
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line17 
         BorderWidth     =   2
         X1              =   7800
         X2              =   7800
         Y1              =   120
         Y2              =   2760
      End
      Begin VB.Line Line16 
         BorderWidth     =   2
         X1              =   3840
         X2              =   7800
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line15 
         BorderWidth     =   2
         X1              =   120
         X2              =   2160
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line14 
         BorderWidth     =   2
         X1              =   120
         X2              =   120
         Y1              =   120
         Y2              =   2760
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Task  Schedule"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2280
         TabIndex        =   22
         Top             =   0
         Width           =   1455
      End
      Begin VB.Image wave2 
         Height          =   480
         Left            =   3360
         Picture         =   "frmmain.frx":AC39
         ToolTipText     =   "Task Running"
         Top             =   1320
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image wave1 
         Height          =   480
         Left            =   3360
         Picture         =   "frmmain.frx":B903
         ToolTipText     =   "Choose Sound To Play "
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image folder2 
         Height          =   480
         Left            =   3360
         Picture         =   "frmmain.frx":C5CD
         ToolTipText     =   "Task Running"
         Top             =   390
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image folder1 
         Height          =   480
         Left            =   3360
         Picture         =   "frmmain.frx":D297
         ToolTipText     =   "Choose File"
         Top             =   390
         Width           =   480
      End
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5400
      TabIndex        =   57
      ToolTipText     =   "About"
      Top             =   2640
      Width           =   495
   End
   Begin VB.Image Image18 
      Height          =   480
      Left            =   4920
      Picture         =   "frmmain.frx":DF61
      ToolTipText     =   "about"
      Top             =   2520
      Width           =   480
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3650
      TabIndex        =   53
      ToolTipText     =   "Options"
      Top             =   2640
      Width           =   645
   End
   Begin VB.Image Image17 
      Height          =   480
      Left            =   3120
      Picture         =   "frmmain.frx":EC2B
      ToolTipText     =   "Options"
      Top             =   2520
      Width           =   480
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6960
      TabIndex        =   49
      ToolTipText     =   "Exit"
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image Image15 
      Height          =   480
      Left            =   6600
      Picture         =   "frmmain.frx":F5F5
      Stretch         =   -1  'True
      ToolTipText     =   "Exit"
      Top             =   2520
      Width           =   480
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Minimize"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Left            =   6720
      TabIndex        =   45
      ToolTipText     =   "Minimize to Tary"
      Top             =   3360
      Width           =   870
   End
   Begin VB.Image Image14 
      Height          =   480
      Left            =   6120
      Picture         =   "frmmain.frx":F73F
      ToolTipText     =   "Minimize to Tary"
      Top             =   3240
      Width           =   480
   End
   Begin VB.Image tick2 
      Height          =   480
      Left            =   4920
      Picture         =   "frmmain.frx":FB81
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Perform Task"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Left            =   1080
      TabIndex        =   25
      Top             =   3360
      Width           =   1260
   End
   Begin VB.Image Image9 
      Height          =   480
      Left            =   3480
      Picture         =   "frmmain.frx":FFC3
      ToolTipText     =   "   ?    "
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image8 
      Height          =   480
      Left            =   480
      Picture         =   "frmmain.frx":10405
      ToolTipText     =   "   ?    "
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Set Alarm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Left            =   3960
      TabIndex        =   26
      Top             =   3360
      Width           =   945
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   3480
      Picture         =   "frmmain.frx":10847
      Top             =   3240
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   480
      Picture         =   "frmmain.frx":10C89
      Top             =   3240
      Width           =   480
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   240
      X2              =   7920
      Y1              =   3083.295
      Y2              =   3083.295
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Info"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3240
      TabIndex        =   18
      Top             =   120
      Width           =   1005
   End
   Begin VB.Line Line26 
      BorderWidth     =   2
      X1              =   240
      X2              =   240
      Y1              =   237.177
      Y2              =   3083.295
   End
   Begin VB.Line Line30 
      BorderWidth     =   2
      X1              =   4320
      X2              =   7920
      Y1              =   237.177
      Y2              =   237.177
   End
   Begin VB.Line Line29 
      BorderWidth     =   2
      X1              =   240
      X2              =   3120
      Y1              =   237.177
      Y2              =   237.177
   End
   Begin VB.Line Line28 
      BorderWidth     =   2
      X1              =   7920
      X2              =   7920
      Y1              =   237.177
      Y2              =   3083.295
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Left            =   4680
      TabIndex        =   17
      Top             =   1440
      Width           =   345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      Height          =   195
      Left            =   4680
      TabIndex        =   16
      Top             =   480
      Width           =   345
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Day of The Year"
      Height          =   195
      Left            =   3855
      TabIndex        =   15
      Top             =   1680
      Width           =   1170
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      Height          =   195
      Left            =   4695
      TabIndex        =   14
      Top             =   1200
      Width           =   330
   End
   Begin VB.Label lbltime 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5400
      TabIndex        =   13
      Top             =   480
      Width           =   135
   End
   Begin VB.Label lblday 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5400
      TabIndex        =   12
      Top             =   720
      Width           =   135
   End
   Begin VB.Label lbldate 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5400
      TabIndex        =   11
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label lbldayofyear 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5520
      TabIndex        =   10
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Day"
      Height          =   195
      Left            =   4740
      TabIndex        =   9
      Top             =   720
      Width           =   285
   End
   Begin VB.Label lblmonth 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5400
      TabIndex        =   8
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      Height          =   195
      Left            =   4575
      TabIndex        =   7
      Top             =   960
      Width           =   450
   End
   Begin VB.Label lblyear 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5400
      TabIndex        =   6
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Days untilll"
      Height          =   195
      Left            =   3840
      TabIndex        =   5
      Top             =   1920
      Width           =   750
   End
   Begin VB.Label lbltill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   4680
      TabIndex        =   4
      Top             =   1920
      Width           =   45
   End
   Begin VB.Line Line4 
      X1              =   3720
      X2              =   7560
      Y1              =   711.53
      Y2              =   711.53
   End
   Begin VB.Line Line5 
      X1              =   3720
      X2              =   7560
      Y1              =   948.706
      Y2              =   948.706
   End
   Begin VB.Line Line6 
      X1              =   3720
      X2              =   7560
      Y1              =   1185.883
      Y2              =   1185.883
   End
   Begin VB.Line Line7 
      X1              =   7560
      X2              =   3720
      Y1              =   1423.059
      Y2              =   1423.059
   End
   Begin VB.Line Line8 
      X1              =   3720
      X2              =   7560
      Y1              =   1660.236
      Y2              =   1660.236
   End
   Begin VB.Line Line9 
      X1              =   3720
      X2              =   7560
      Y1              =   1897.412
      Y2              =   1897.412
   End
   Begin VB.Line Line10 
      X1              =   3720
      X2              =   7560
      Y1              =   2134.589
      Y2              =   2134.589
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IT `s Now"
      Height          =   195
      Left            =   4335
      TabIndex        =   3
      Top             =   2160
      Width           =   690
   End
   Begin VB.Line Line11 
      X1              =   5160
      X2              =   5160
      Y1              =   474.353
      Y2              =   1660.236
   End
   Begin VB.Line Line12 
      X1              =   5160
      X2              =   5160
      Y1              =   1660.236
      Y2              =   2371.765
   End
   Begin VB.Shape Shape1 
      Height          =   1935
      Left            =   3720
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label lblleft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   195
      Left            =   5400
      TabIndex        =   2
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label lblnow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   195
      Left            =   5400
      TabIndex        =   1
      Top             =   2160
      Width           =   135
   End
   Begin VB.Image tick1 
      Height          =   480
      Left            =   2400
      Picture         =   "frmmain.frx":12983
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu menuBar 
      Caption         =   "systray"
      Visible         =   0   'False
      Begin VB.Menu mnuMain 
         Caption         =   "Restore"
         Index           =   0
      End
      Begin VB.Menu mnuMain 
         Caption         =   "About"
         Index           =   1
      End
      Begin VB.Menu mnuMain 
         Caption         =   "Exit"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim stsk As Boolean, BeepCheck As Boolean
Dim HourLength As Integer, MinuteLength As Integer, _
SecondLength As Integer
Dim MidX As Integer, MidY As Integer
Dim dateVar As Date, timevar As Date, timenow As String, datenow As String, TaskFileVar As String, SoundFileVar As String, Beepvar As String
Dim checkdate As String, checktime As String
Dim AlarmCheckdate As String, AlarmChecktime As String
Dim iday As String, iyear As String, imonth As String, idayname As String
Dim TaskStarted As Boolean, AlarmStarted As Boolean
Dim x, y, r, xm, b As Boolean
Dim at As Date, st As Date
Const PI = 3.14159

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim t As NOTIFYICONDATA

Sub LengthAndCentre()
On Error Resume Next
    Dim d As Integer
    If Frame1.Width < Frame1.Height Then
        HourLength = Frame1.Width * 30 / 200 ' 50%
        MinuteLength = Frame1.Width * 45 / 200 ' 80%
        SecondLength = Frame1.Width * 57 / 200 ' 90%
    Else
        HourLength = Frame1.Height * 30 / 200 ' 50%
        MinuteLength = Frame1.Height * 45 / 200 ' 80%
        SecondLength = Frame1.Height * 57 / 200 ' 90%
    End If
    MidX = Frame1.Width \ 2
    MidY = Frame1.Height \ 2
    Line1.X1 = MidX
    Line2.X1 = MidX
    Line3.X1 = MidX
   
    Line1.Y1 = MidY
    Line2.Y1 = MidY
    Line3.Y1 = MidY
    Call Timer1_Timer 'just To avoid flicker
End Sub
Private Sub canceltask_Click()
On Error Resume Next
Call WriteToINI("Task", "Started", "0", App.Path & "\Hoursdata.dat")
Call WriteToINI("Task", "Date", "00/00/00", App.Path & "\Hoursdata.dat")
Call WriteToINI("Task", "time", "00:00:00", App.Path & "\Hoursdata.dat")

Text1.Enabled = True
Text2.Enabled = True
folder1.Visible = True
wave1.Visible = True
folder2.Visible = False
wave2.Visible = False
picktaskdate.Enabled = True
picktasktime.Enabled = True
TaskStarted = False
TaskStarted = False
tick1.Visible = False
Timer3.Enabled = False
End Sub

Private Sub Check1_Click()
On Error Resume Next
If Check1.Value = 1 Then
 Text3.Enabled = False
 Image7.Visible = False
 Image6.Visible = True
 BeepCheck = True
 lblbeep.Visible = True
 Label12.Visible = False
 Else
 Text3.Enabled = True
 Image7.Visible = True
 Image6.Visible = False
 BeepCheck = False
  lblbeep.Visible = False
  Label12.Visible = False
  Timer5.Enabled = False
 End If

End Sub

Private Sub Check2_Click()
On Error Resume Next
If Check2.Value = 1 Then
Call WriteToINI("Startup", "Minimize", "1", App.Path & "\Hoursdata.dat")
Else
Call WriteToINI("Startup", "Minimize", "0", App.Path & "\Hoursdata.dat")
End If
End Sub

Private Sub Check3_Click()
On Error Resume Next
Dim StrName As String
StrName = "Hours.exe/background" ' name That will appear in startup

If Check3.Value = 1 Then
Call savestring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", StrName, App.Path & "\" & App.EXEName & ".exe")
Call WriteToINI("Startup", "AutoStart", "1", App.Path & "\Hoursdata.dat")
Else
StrName = "Hours.exe/background"
Call DeleteValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", StrName)
Call WriteToINI("Startup", "AutoStart", "0", App.Path & "\Hoursdata.dat")
End If

End Sub



Private Sub DTPicker1_Change()
On Error Resume Next
Dim hour As String, min As String, sec As String, itask As String, mytime As String

mytime = Format(DTPicker1.Value, "h:m:s")
hour = DTPicker1.hour
min = DTPicker1.Minute
sec = DTPicker1.Second
Call WriteToINI("Alarm", "Time", mytime, App.Path & "\Hoursdata.dat")
Call WriteToINI("Alarm", "Hour", hour, App.Path & "\Hoursdata.dat")
Call WriteToINI("Alarm", "Minutes", min, App.Path & "\Hoursdata.dat")
Call WriteToINI("Alarm", "seconds", sec, App.Path & "\Hoursdata.dat")

End Sub

Private Sub DTPicker2_Change()
On Error Resume Next
Dim hour As String, min As String, sec As String, itask As String, mytime As String
itask = Format(DTPicker2.Value, "dddd, mmm d yyyy")
iday = DTPicker2.Day
imonth = DTPicker2.Month
iyear = DTPicker2.Year
idayname = DTPicker2.DayOfWeek
Call WriteToINI("Alarm", "date", DTPicker2.Value, App.Path & "\Hoursdata.dat")
Call WriteToINI("Alarm", "Day", iday, App.Path & "\Hoursdata.dat")
Call WriteToINI("Alarm", "Month", imonth, App.Path & "\Hoursdata.dat")
Call WriteToINI("Alarm", "Year", iyear, App.Path & "\Hoursdata.dat")
Call WriteToINI("Alarm", "InfoDate", itask, App.Path & "\Hoursdata.dat")

End Sub

Private Sub folder1_Click()
On Error Resume Next
CommonDialog1.Filter = "All Files|*.*"
CommonDialog1.ShowOpen
Text1.Text = CommonDialog1.FileName
End Sub

Private Sub Form_Load()
On Error Resume Next

   t.cbSize = Len(t)
    t.hWnd = Picture1.hWnd
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    'this is where the Form's Icon gets called
          t.hIcon = Me.Icon
    'this is where the tool tip goes
          t.szTip = "Fanee`s Hours" & Chr$(0)
    
    Shell_NotifyIcon NIM_ADD, t
    
    App.TaskVisible = False


Dim datevar1, datevar2, timevar1, timevar2
datevar1 = GetFromINI("Task", "date", App.Path & "\Hoursdata.dat")
timevar1 = GetFromINI("Task", "Time", App.Path & "\Hoursdata.dat")
datevar2 = GetFromINI("Alarm", "date", App.Path & "\Hoursdata.dat")
timevar2 = GetFromINI("Alarm", "Time", App.Path & "\Hoursdata.dat")
DTPicker1.Value = timevar2
DTPicker2.Value = datevar2
picktasktime.Value = timevar1
picktaskdate.Value = datevar1


Dim ChecktasVar, CheckalaVar, checkstart, checkAutostart
CheckalaVar = GetFromINI("Alarm", "Started", App.Path & "\Hoursdata.dat")
ChecktasVar = GetFromINI("Task", "Started", App.Path & "\Hoursdata.dat")
checkstart = GetFromINI("Startup", "Minimize", App.Path & "\Hoursdata.dat")
checkAutostart = GetFromINI("Startup", "Autostart", App.Path & "\Hoursdata.dat")
If CheckalaVar = 1 Then
Call Image5_Click
End If

If ChecktasVar = 1 Then
Call settask_Click
End If

If checkstart = 1 Then
Me.Hide
Check2.Value = 1
End If
If checkstart = 0 Then
Check2.Value = 0
End If
If checkAutostart = 1 Then
Check3.Value = 1
Else
Check3.Value = 0
End If

 Me.AutoRedraw = True
    Line1.BorderWidth = 7
    Line2.BorderWidth = 3
    Line3.BorderWidth = 1
    Timer1.Interval = 1000
    Call LengthAndCentre
    Call Timer1_Timer
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Image10_Click()
On Error Resume Next
Call Image9_Click
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    t.cbSize = Len(t)
    t.hWnd = Picture1.hWnd
    t.uId = 1&
    Shell_NotifyIcon NIM_DELETE, t
End Sub

Private Sub Image15_Click()
End
End Sub

Private Sub Image17_Click()
Frame4.Visible = True
End Sub

Private Sub Image18_Click()
frmabout.Show
frmabout.SetFocus
End Sub

Private Sub Label24_Click()
Call Image5_Click
End Sub

Private Sub Label25_Click()
Call Image4_Click
End Sub

Private Sub Label28_Click()
End
End Sub

Private Sub Label29_Click()
Frame4.Visible = True
End Sub

Private Sub Label31_Click()
Frame4.Visible = False
End Sub

Private Sub Label32_Click()
frmabout.Show
frmabout.SetFocus
End Sub

Private Sub mnuMain_Click(Index As Integer)

Select Case Index
    Case 0
    Me.Show
    Me.WindowState = 0
    Case 1
    frmabout.Show
    Case 2
    End
    Case Else
End Select
End Sub

Private Sub Image11_Click()
On Error Resume Next
Call Image9_Click
End Sub

Private Sub Image12_Click()
On Error Resume Next
Call Image8_Click
End Sub

Private Sub Image13_Click()
On Error Resume Next
Call Image8_Click

End Sub

Private Sub Image14_Click()
Me.Hide
End Sub

Private Sub Image2_Click()
Call Label2_Click
End Sub

Private Sub Image3_Click()
Call Label3_Click
End Sub

Private Sub Image4_Click()
On Error Resume Next
Call WriteToINI("Alarm", "Started", "0", App.Path & "\Hoursdata.dat")
Call WriteToINI("Alarm", "Date", "00/00/00", App.Path & "\Hoursdata.dat")
Call WriteToINI("Alarm", "time", "00:00:00", App.Path & "\Hoursdata.dat")

Image5.Enabled = True
Text3.Enabled = True
Check1.Enabled = True
Image7.Visible = True
Image6.Visible = False
DTPicker1.Enabled = True
DTPicker2.Enabled = True
AlarmStarted = False
Timer4.Enabled = False
tick2.Visible = False
End Sub

Private Sub Image5_Click()
On Error Resume Next
AlarmChecktime = GetFromINI("Alarm", "Time", App.Path & "\Hoursdata.dat")
AlarmCheckdate = GetFromINI("Alarm", "InfoDate", App.Path & "\Hoursdata.dat")

datenow = Format(Date, "dddd, mmm d yyyy")
timenow = Format(Time, "h:m:s")

Image5.Enabled = False
Text3.Enabled = False
Check1.Enabled = False
Image7.Visible = False
Image6.Visible = True
DTPicker1.Enabled = False
DTPicker2.Enabled = False
AlarmStarted = True

If BeepCheck = True Then
Dim ibeep As String
Call WriteToINI("Alarm", "Beep", "1", App.Path & "\Hoursdata.dat")
Call WriteToINI("Alarm", "Soundfile", "Choose Sound To Play", App.Path & "\Hoursdata.dat")
Else
Dim Soundfile As String
Soundfile = Text3.Text
Call WriteToINI("Alarm", "Soundfile", Soundfile, App.Path & "\Hoursdata.dat")
Call WriteToINI("Alarm", "Beep", "0", App.Path & "\Hoursdata.dat")
End If

Call DTPicker2_Change
Call DTPicker1_Change
Call WriteToINI("Alarm", "Started", "1", App.Path & "\Hoursdata.dat")


Timer4.Enabled = True

Me.Height = 4300
Frame3.Visible = False
Label3.Enabled = True
Image9.Visible = False
Image3.Visible = True
Label2.Enabled = True
Image2.Visible = True
End Sub

Private Sub Image7_Click()
On Error Resume Next
CommonDialog1.Filter = "WaveAudio(*.wav)|*.wav| MP3(*.mp3)|*.mp3|"
CommonDialog1.ShowOpen
Text3.Text = CommonDialog1.FileName
End Sub

Private Sub Image8_Click()
On Error Resume Next
Me.Height = 4300
Frame2.Visible = False
Image2.Visible = True
Image8.Visible = False
Image3.Visible = True
Label3.Enabled = True
Label2.Enabled = True
End Sub

Private Sub Image9_Click()
On Error Resume Next
Me.Height = 4300
Frame3.Visible = False
Label3.Enabled = True
Image9.Visible = False
Image3.Visible = True
Label2.Enabled = True
Image2.Visible = True
If AlarmStarted = True Then
tick2.Visible = True
End If
End Sub

Private Sub Label12_Click()
Timer5.Enabled = False
Label12.Visible = False
lblbeep.Visible = True
End Sub

Private Sub Label2_Click()
On Error Resume Next
Me.Height = 7400
Frame2.Visible = True

dateVar = GetFromINI("Task", "date", App.Path & "\Hoursdata.dat")
picktaskdate.Value = dateVar
timevar = GetFromINI("Task", "Time", App.Path & "\Hoursdata.dat")
picktasktime.Value = timevar
SoundFileVar = GetFromINI("Task", "soundfile", App.Path & "\Hoursdata.dat")
Text2.Text = SoundFileVar
TaskFileVar = GetFromINI("Task", "taskfile", App.Path & "\Hoursdata.dat")
Text1.Text = TaskFileVar


Image2.Visible = False
Image8.Visible = True
Image3.Visible = False
Label3.Enabled = False
Label2.Enabled = False


End Sub

Private Sub Label27_Click()
Me.Hide
End Sub

Private Sub Label3_Click()
On Error Resume Next
Me.Height = 7400
Frame3.Visible = True

dateVar = GetFromINI("Alarm", "date", App.Path & "\Hoursdata.dat")
DTPicker2.Value = dateVar
timevar = GetFromINI("Alarm", "Time", App.Path & "\Hoursdata.dat")
DTPicker1.Value = timevar


SoundFileVar = GetFromINI("Alarm", "soundfile", App.Path & "\Hoursdata.dat")
Beepvar = GetFromINI("Alarm", "Beep", App.Path & "\Hoursdata.dat")
If Beepvar = 0 Then
Text3.Text = SoundFileVar
Else
 Text3.Enabled = False
 Image7.Visible = False
 Image6.Visible = True
 BeepCheck = True
 Check1.Value = 1
End If

Label3.Enabled = False
Image9.Visible = True
Image3.Visible = False
Label2.Enabled = False
Image2.Visible = False

End Sub

Private Sub lblbeep_Click()
Timer5.Enabled = True
Label12.Visible = True
lblbeep.Visible = False
End Sub

Private Sub settask_Click()
On Error Resume Next
Text1.Enabled = False
Text2.Enabled = False
folder1.Visible = False
wave1.Visible = False
folder2.Visible = True
wave2.Visible = True
picktaskdate.Enabled = False
picktasktime.Enabled = False
TaskStarted = True

Dim Taskfile As String, Soundfile As String
Taskfile = Text1.Text
Soundfile = Text2.Text

Call WriteToINI("Task", "Taskfile", Taskfile, App.Path & "\Hoursdata.dat")
Call WriteToINI("Task", "Soundfile", Soundfile, App.Path & "\Hoursdata.dat")
Call picktaskdate_Change
Call picktasktime_Change
Call WriteToINI("Task", "Started", "1", App.Path & "\Hoursdata.dat")

Timer3.Enabled = True

Me.Height = 4300
Frame2.Visible = False
Image2.Visible = True
Image8.Visible = False
Image3.Visible = True
Label3.Enabled = True
Label2.Enabled = True
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
Dim Hours As Single, Minutes As Single, Seconds As Single
    Dim TrueHours As Single
    'Beep
    Hours = hour(Time)
    Minutes = Minute(Time)
    Seconds = Second(Time)
    TrueHours = Hours + Minutes / 60
    'HourHand
    Line1.X2 = HourLength * Cos(PI / 180 * (30 * TrueHours - 90)) + MidX
    Line1.Y2 = HourLength * Sin(PI / 180 * (30 * TrueHours - 90)) + MidY
    'MinuteHand
    Line2.X2 = MinuteLength * Cos(PI / 180 * (6 * Minutes - 90)) + MidX
    Line2.Y2 = MinuteLength * Sin(PI / 180 * (6 * Minutes - 90)) + MidY
    'SecondsHand
    Line3.X2 = SecondLength * Cos(PI / 180 * (6 * Seconds - 90)) + MidX
    Line3.Y2 = SecondLength * Sin(PI / 180 * (6 * Seconds - 90)) + MidY
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
lbldate.Caption = Format(Date, "dddd, mmm d yyyy")
Dim a
a = Format(Time, "hh:mm:ss")
lbltime.Caption = a
lbldayofyear.Caption = DatePart("y", Date)
lblyear.Caption = DatePart("yyyy", Date)
lbltill.Caption = Val(DatePart("yyyy", Date)) + 1
lblleft.Caption = 365 - Val(DatePart("y", Date))

Select Case DatePart("w", Date)
        Case "1"
    lblday.Caption = "Sunday"
        Case "2"
    lblday.Caption = "Monday"
        Case "3"
    lblday.Caption = "Tuesday"
        Case "4"
    lblday.Caption = "Wednesday"
        Case "5"
    lblday.Caption = "Thursday"
        Case "6"
    lblday.Caption = "Friday"
        Case "7"
    lblday.Caption = "Saturday"
End Select

Select Case DatePart("m", Date)
        Case "1"
lblmonth.Caption = "January"
        Case "2"
lblmonth.Caption = "February"
        Case "3"
lblmonth.Caption = "March"
        Case "4"
lblmonth.Caption = "April"
        Case "5"
lblmonth.Caption = "May"
        Case "6"
lblmonth.Caption = "June"
        Case "7"
lblmonth.Caption = "July"
        Case "8"
lblmonth.Caption = "August"
        Case "9"
lblmonth.Caption = "September"
        Case "10"
lblmonth.Caption = "October"
        Case "11"
lblmonth.Caption = "November"
        Case "12"
lblmonth.Caption = "December"
End Select
    
Select Case Format(Time, "hh:mm:ss")
    Case "23:00:00" To "23:59:59"
    lblnow.Caption = "Approaching Midnight"
    Case "21:00:00" To "22:59:59"
    lblnow.Caption = "Late Night"
    Case "00:00:00" To "00:59:59"
    lblnow.Caption = "Midnight"
    Case "01:00:00" To "05:59:59"
    lblnow.Caption = "Early Morning"
    Case "06:00:00" To "11:59:59"
    lblnow.Caption = "Morning"
    Case "12:00:00" To "12:59:59"
    lblnow.Caption = "Midday"
    Case "13:00:00" To "17:59:59"
    lblnow.Caption = "Afternoon"
    Case "18:00:00" To "20:59:59"
    lblnow.Caption = "Evening"
End Select

Dim datevar1, datevar2, timevar1, timevar2
datevar1 = GetFromINI("Task", "date", App.Path & "\Hoursdata.dat")
timevar1 = GetFromINI("Task", "Time", App.Path & "\Hoursdata.dat")
datevar2 = GetFromINI("Alarm", "date", App.Path & "\Hoursdata.dat")
timevar2 = GetFromINI("Alarm", "Time", App.Path & "\Hoursdata.dat")
datevar1 = Format(datevar1, "dddd, mmm d, yyyy")
datevar2 = Format(datevar2, "dddd, mmm d, yyyy")
timevar1 = Format(timevar1, "hh:mm:ss")
timevar2 = Format(timevar2, "hh:mm:ss")

Label2.ToolTipText = "Task Date: " & datevar1 & "   Task Time : " & timevar1
Label3.ToolTipText = "Alarm Date: " & datevar2 & "   Alarm Time : " & timevar2
Image2.ToolTipText = "Task Date: " & datevar1 & "   Task Time : " & timevar1
Image3.ToolTipText = "Alarm Date: " & datevar2 & "   Alarm Time : " & timevar2

End Sub

Private Sub picktaskdate_Change()
On Error Resume Next
Dim hour As String, min As String, sec As String, itask As String, mytime As String
itask = Format(picktaskdate.Value, "dddd, mmm d yyyy")
iday = picktaskdate.Day
imonth = picktaskdate.Month
iyear = picktaskdate.Year
idayname = picktaskdate.DayOfWeek
Call WriteToINI("Task", "date", picktaskdate.Value, App.Path & "\Hoursdata.dat")
Call WriteToINI("Task", "Day", iday, App.Path & "\Hoursdata.dat")
Call WriteToINI("Task", "Month", imonth, App.Path & "\Hoursdata.dat")
Call WriteToINI("Task", "Year", iyear, App.Path & "\Hoursdata.dat")
Call WriteToINI("Task", "InfoDate", itask, App.Path & "\Hoursdata.dat")
End Sub
Private Sub picktasktime_Change()
On Error Resume Next
Dim hour As String, min As String, sec As String, itask As String, mytime As String

mytime = Format(picktasktime.Value, "h:m:s")
hour = picktasktime.hour
min = picktasktime.Minute
sec = picktasktime.Second
Call WriteToINI("Task", "Time", mytime, App.Path & "\Hoursdata.dat")
Call WriteToINI("Task", "Hour", hour, App.Path & "\Hoursdata.dat")
Call WriteToINI("Task", "Minutes", min, App.Path & "\Hoursdata.dat")
Call WriteToINI("Task", "seconds", sec, App.Path & "\Hoursdata.dat")
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
If TaskStarted = True Then
tick1.Visible = True
End If

checktime = GetFromINI("Task", "Time", App.Path & "\Hoursdata.dat")
checkdate = GetFromINI("Task", "InfoDate", App.Path & "\Hoursdata.dat")

datenow = Format(Date, "dddd, mmm d yyyy")
timenow = Format(Time, "h:m:s")

If timenow = checktime And datenow = checkdate Then
frmtask.Show
frmtask.SetFocus
Call WriteToINI("Task", "Date", "00/00/00", App.Path & "\Hoursdata.dat")
Call WriteToINI("Task", "time", "00:00:00", App.Path & "\Hoursdata.dat")

Text1.Enabled = True
Text2.Enabled = True
folder1.Visible = True
wave1.Visible = True
folder2.Visible = False
wave2.Visible = False
picktaskdate.Enabled = True
picktasktime.Enabled = True
TaskStarted = False
TaskStarted = False
Call WriteToINI("Task", "Started", "0", App.Path & "\Hoursdata.dat")
Timer3.Enabled = False
tick1.Visible = False
End If

End Sub

Private Sub Timer4_Timer()
On Error Resume Next
If AlarmStarted = True Then
tick2.Visible = True
End If
AlarmChecktime = GetFromINI("Alarm", "Time", App.Path & "\Hoursdata.dat")
AlarmCheckdate = GetFromINI("Alarm", "InfoDate", App.Path & "\Hoursdata.dat")

datenow = Format(Date, "dddd, mmm d yyyy")
timenow = Format(Time, "h:m:s")

If timenow = AlarmChecktime And datenow = AlarmCheckdate Then
frmalarm.Show
frmalarm.SetFocus
Call WriteToINI("Alarm", "Date", "00/00/00", App.Path & "\Hoursdata.dat")
Call WriteToINI("Alarm", "time", "00:00:00", App.Path & "\Hoursdata.dat")

Image5.Enabled = True
Text3.Enabled = True
Check1.Enabled = True
Image7.Visible = True
Image6.Visible = False
DTPicker1.Enabled = True
DTPicker2.Enabled = True
AlarmStarted = False
Call WriteToINI("Alarm", "Started", "0", App.Path & "\Hoursdata.dat")
Timer4.Enabled = False
tick2.Visible = False
End If
End Sub

Private Sub Timer5_Timer()
Beep
End Sub

Private Sub wave1_Click()
On Error Resume Next
CommonDialog1.Filter = "WaveAudio(*.wav)|*.wav| MP3(*.mp3)|*.mp3|"
CommonDialog1.ShowOpen
Text2.Text = CommonDialog1.FileName
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Static rec As Boolean, msg As Long
    Dim RetVal As String
    Dim returnstring
    Dim retvalue
    msg = x / Screen.TwipsPerPixelX
    If rec = False Then
        rec = True
    Select Case msg
   
    Case WM_RBUTTONDBLCLK: 'not used in this program
    Case WM_RBUTTONDOWN:   'not used in this program
    Case WM_LBUTTONUP:
    Me.WindowState = 1
    Case WM_RBUTTONUP:
    'if Right Mouse Button is down then
    'Bring up the Popup Menu
       Me.PopupMenu menuBar
    End Select
        rec = False
    End If
End Sub
Public Sub savestring(Hkey As Long, strPath As String, strValue As String, strdata As String)
    'EXAMPLE:
    'Call savestring(HKEY_CURRENT_USER, "Software\VBW\Registry", "String", text1.text)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(Hkey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)
End Sub

Public Function DeleteValue(ByVal Hkey As Long, ByVal strPath As String, ByVal strValue As String)
    'EXAMPLE:
    'Call DeleteValue(HKEY_CURRENT_USER, "Software\VBW\Registry", "Dword")
    Dim keyhand As Long
    Dim r As Long
    r = RegOpenKey(Hkey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function
