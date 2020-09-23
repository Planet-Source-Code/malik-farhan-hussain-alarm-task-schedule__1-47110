VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Author"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      Index           =   0
      X1              =   2400
      X2              =   5160
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   4680
      Picture         =   "about.frx":0CCA
      Stretch         =   -1  'True
      ToolTipText     =   "Ok"
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   360
      Picture         =   "about.frx":0E14
      Top             =   120
      Width           =   1530
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Me And Myself"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3240
      TabIndex        =   5
      Top             =   2280
      Width           =   1050
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Special Thanks:"
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
      Left            =   2160
      TabIndex        =   4
      Top             =   1920
      Width           =   1350
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks to www.planet-source-code.com for Guidance"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Width           =   3840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   5400
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "iamfanee@hotmail.com"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3000
      TabIndex        =   2
      Top             =   1200
      Width           =   1650
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact:"
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
      Left            =   2145
      TabIndex        =   1
      Top             =   960
      Width           =   705
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Malik Farhan Hussain"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2505
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image2_Click()
Unload Me
End Sub
