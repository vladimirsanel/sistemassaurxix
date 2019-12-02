VERSION 5.00
Begin VB.Form FormArteIncisos 
   BackColor       =   &H00FF8080&
   Caption         =   "ARTÍCULOS E INCISOS"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame FrameDesp 
      BackColor       =   &H00FF8080&
      Height          =   735
      Left            =   3840
      TabIndex        =   14
      Top             =   9600
      Width           =   3495
      Begin VB.CommandButton CmdUltimo 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "BOOTLE"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2520
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdSiguiente 
         Caption         =   ">I"
         BeginProperty Font 
            Name            =   "BOOTLE"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1800
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdAnterior 
         Caption         =   "I<"
         BeginProperty Font 
            Name            =   "BOOTLE"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1080
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdPrimero 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "BOOTLE"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   360
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton CmdMDISSAURXIX 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PANTALLA PRINCIPAL"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9720
      Width           =   2415
   End
   Begin VB.CommandButton CmdActAdmin 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ACTUACIONES ADMINISTRATIVAS"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9720
      Width           =   2415
   End
   Begin VB.CommandButton CmdPersonalPolicial 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PERSONAL POLICIAL"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9720
      Width           =   2415
   End
   Begin VB.CommandButton CmdSalir 
      BackColor       =   &H000000FF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   19680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton CmdGuardar 
      BackColor       =   &H00FFC0C0&
      Caption         =   "GUARDAR"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton CmdBorrar 
      BackColor       =   &H00FFC0C0&
      Caption         =   "BORRAR"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton CmdNuevo 
      BackColor       =   &H00FFC0C0&
      Caption         =   "NUEVO REGISRO"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6000
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5790
      Left            =   3840
      TabIndex        =   5
      Top             =   3720
      Width           =   13815
   End
   Begin VB.TextBox TxtArticulo 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3840
      TabIndex        =   1
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox TxtInciso 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6240
      TabIndex        =   0
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label LblDescripcionArt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCIÓN"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label LblArteIncisos 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ARTÍCULOS E INCISOS"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   4
      Top             =   1200
      Width           =   13575
   End
   Begin VB.Label LblArticulo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ARTÍCULO"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label LblInciso 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "INCISO"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
End
Attribute VB_Name = "FormArteIncisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LblActuaciones_Click()

End Sub

Private Sub CmdActAdmin_Click()
FormActAdmin.Show
End Sub

Private Sub CmdMDISSAURXIX_Click()
MDISSAURXIX.Show
End Sub

Private Sub CmdPersonalPolicial_Click()
FormPersonalPolicial.Show
End Sub

Private Sub CmdSalir_Click()
    If MsgBox("              ¿SALIR?", vbQuestion + vbYesNo, "SALIR") = vbYes Then
    Unload Me
    End If
End Sub
