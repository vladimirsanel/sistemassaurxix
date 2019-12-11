VERSION 5.00
Begin VB.Form FormAcercade 
   BackColor       =   &H00FF8080&
   Caption         =   "ACERDA DE..."
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdSalir 
      BackColor       =   &H000000FF&
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   18960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   360
      Width           =   1335
   End
   Begin VB.Frame FrameAcercade 
      BackColor       =   &H00FF8080&
      Height          =   5055
      Left            =   12360
      TabIndex        =   0
      Top             =   3120
      Width           =   5175
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "-  VERA, SANTA FE  -"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   4320
         Width           =   3735
      End
      Begin VB.Image Image1 
         Height          =   420
         Left            =   360
         Picture         =   "FormAcerdade.frx":0000
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   420
      End
      Begin VB.Image Image4 
         Height          =   420
         Left            =   360
         Picture         =   "FormAcerdade.frx":B42A
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   420
      End
      Begin VB.Image Image3 
         Height          =   420
         Left            =   360
         Picture         =   "FormAcerdade.frx":15D36
         Stretch         =   -1  'True
         Top             =   480
         Width           =   420
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "03483 - 154 53 576"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   480
         Width           =   3135
      End
      Begin VB.Image Image5 
         Height          =   420
         Left            =   360
         Picture         =   "FormAcerdade.frx":20258
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   420
      End
      Begin VB.Image Image2 
         Height          =   420
         Left            =   360
         Picture         =   "FormAcerdade.frx":22DCD
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   420
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "vladimirsanel"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "vladimirsanel"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "vladimirsanel"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "vladimir.sanel.vs@gmail.com"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   2880
         Width           =   4335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "Allende Lezama N° 1436"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Top             =   3960
         Width           =   3735
      End
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "PROGRAMACIONES"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   12360
      TabIndex        =   9
      Top             =   2520
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "VLADIMIR SANEL"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   12360
      TabIndex        =   8
      Top             =   2040
      Width           =   5175
   End
   Begin VB.Image Image6 
      Height          =   6105
      Left            =   2400
      Picture         =   "FormAcerdade.frx":23C2B
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   9810
   End
End
Attribute VB_Name = "FormAcercade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdSalir_Click()
Unload Me
End Sub
