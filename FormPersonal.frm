VERSION 5.00
Begin VB.Form FormPersonalPolicial 
   BackColor       =   &H00FF8080&
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame FrameConsulta 
      BackColor       =   &H00FF8080&
      Height          =   3255
      Left            =   1680
      TabIndex        =   9
      Top             =   6120
      Width           =   13455
   End
   Begin VB.CommandButton CmdEscyJerarquias 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ESCALAFÓN Y JERARQUÍA"
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
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   9840
      Width           =   2415
   End
   Begin VB.Frame FrameBuscar 
      BackColor       =   &H00FF8080&
      Height          =   1335
      Left            =   16200
      TabIndex        =   27
      Top             =   3120
      Width           =   3255
      Begin VB.TextBox TxtBuscar 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   2775
      End
      Begin VB.CommandButton CmdBuscar 
         BackColor       =   &H00FFC0C0&
         Caption         =   "BUSCAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1305
      End
   End
   Begin VB.Frame FrameCmd 
      BackColor       =   &H00FF8080&
      Height          =   3015
      Left            =   16200
      TabIndex        =   26
      Top             =   4800
      Width           =   3255
      Begin VB.CommandButton CmdModificar 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MODIFICAR"
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CommandButton CmdActualizar 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ACTUALIZAR"
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2280
         Width           =   2295
      End
      Begin VB.CommandButton CmdAlta 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ALTA"
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton CmdBaja 
         BackColor       =   &H00FFC0C0&
         Caption         =   "BAJA"
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   960
         Width           =   2295
      End
   End
   Begin VB.Frame FrameDatos 
      BackColor       =   &H00FF8080&
      Height          =   4455
      Left            =   1680
      TabIndex        =   20
      Top             =   1560
      Width           =   13455
      Begin VB.TextBox TxtDni 
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
         Left            =   10440
         TabIndex        =   5
         Top             =   2280
         Width           =   2175
      End
      Begin VB.CommandButton CmdRegistrar 
         BackColor       =   &H00FFC0C0&
         Caption         =   "REGISTRAR"
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3960
         Width           =   1575
      End
      Begin VB.CommandButton CmdBorrar 
         BackColor       =   &H00FFC0C0&
         Caption         =   "BORRAR"
         BeginProperty Font 
            Name            =   "Bauhaus 93"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11040
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3960
         Width           =   1575
      End
      Begin VB.TextBox TxtApeyNom 
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
         Left            =   2400
         TabIndex        =   3
         Top             =   1440
         Width           =   10215
      End
      Begin VB.TextBox TxtNi 
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
         Left            =   10440
         TabIndex        =   2
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox TxtDestino 
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
         Left            =   2400
         TabIndex        =   4
         Top             =   2280
         Width           =   6135
      End
      Begin VB.TextBox TxtJerarquia 
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
         Left            =   2400
         TabIndex        =   1
         Top             =   720
         Width           =   6135
      End
      Begin VB.TextBox TxtObs 
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
         Left            =   2400
         TabIndex        =   6
         Top             =   3120
         Width           =   10215
      End
      Begin VB.Label LblDni 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "D.N.I."
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
         Left            =   9720
         TabIndex        =   28
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label LblApeyNom 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "APELLIDO Y NOMBRE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         TabIndex        =   25
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label LblNi 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "N.I."
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
         Left            =   9840
         TabIndex        =   24
         Top             =   720
         Width           =   495
      End
      Begin VB.Label LblDestino 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "DESTINO"
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
         Left            =   1080
         TabIndex        =   23
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label LblJerarquia 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "JERARQUIA"
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
         Left            =   720
         TabIndex        =   22
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label LblObs 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACIONES"
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
         Left            =   360
         TabIndex        =   21
         Top             =   3120
         Width           =   1815
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   9840
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   9840
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
      TabIndex        =   19
      Top             =   240
      Width           =   375
   End
   Begin VB.Label LblPersonal 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "PERSONAL POLICIAL"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   13215
   End
End
Attribute VB_Name = "FormPersonalPolicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdActAdmin_Click()
   FormActAdmin.Show
End Sub


'-------------ALTA-------------'
Private Sub CmdAlta_Click()
   CmdBorrar.Enabled = True
   CmdRegistrar.Enabled = True
   CmdModificar.Enabled = False
   CmdActualizar.Enabled = False
   CmdBaja.Enabled = False
   TxtBuscar.Enabled = False
   CmdBuscar.Enabled = True
   FrameDatos.Enabled = True

   TxtJerarquia.Text = ""
   TxtNi.Text = ""
   TxtApeyNom.Text = ""
   TxtDestino.Text = ""
   TxtDni.Text = ""
   TxtObs.Text = ""
   TxtBuscar.Text = ""
End Sub


Private Sub CmdArteIncisos_Click()
FormArteIncisos.Show
End Sub


'-----------BORRAR-----------'
Private Sub CmdBorrar_Click()
If MsgBox("¿DESEA BORRAR Y CANCELAR EL INGRESO DE DATOS?", vbQuestion + vbYesNo, "ATENCION!") = vbYes Then
   CmdBorrar.Enabled = False
   CmdRegistrar.Enabled = False
   CmdAlta.Enabled = True
   CmdModificar.Enabled = False
   CmdActualizar.Enabled = False
   CmdBaja.Enabled = False
   TxtBuscar.Enabled = False
   CmdBuscar.Enabled = True
   FrameDatos.Enabled = False

   TxtJerarquia.Text = ""
   TxtNi.Text = ""
   TxtApeyNom.Text = ""
   TxtDni.Text = ""
   TxtObs.Text = ""
   
   CmdAlta.SetFocus
End If
End Sub

Private Sub CmdBuscar_Click()
   TxtJerarquia.Enabled = False
   TxtNi.Enabled = False
   TxtApeyNom.Enabled = False
   TxtDestino.Enabled = False
   TxtDni.Enabled = False
   TxtObs.Enabled = False
   CmdBaja.Enabled = True
   CmdModificar.Enabled = True
   CmdActualizar.Enabled = True
End Sub

Private Sub CmdEscyJerarquias_Click()
FormEscyJerarquia.Show
End Sub

Private Sub CmdMDISSAURXIX_Click()
MDISSAURXIX.Show
End Sub

Private Sub CmdSalir_Click()
    If MsgBox("               ¿SALIR?", vbQuestion + vbYesNo, "SALIR") = vbYes Then
    Unload Me
    End If
End Sub
