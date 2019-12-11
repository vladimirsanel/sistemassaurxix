VERSION 5.00
Begin VB.MDIForm MDISSAURXIX 
   BackColor       =   &H00C00000&
   Caption         =   "REGISTRO DE ACTUACIONES ADMINISTRATIVAS POLICIALES"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   1170
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   4170
      Left            =   0
      ScaleHeight     =   852.757
      ScaleMode       =   0  'User
      ScaleWidth      =   20250
      TabIndex        =   4
      Top             =   3045
      Width           =   20250
      Begin VB.CommandButton CmdEscyJer 
         BackColor       =   &H00C0C000&
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
         Height          =   615
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CommandButton CmdPersonalPolicial 
         BackColor       =   &H00C0C000&
         Caption         =   "PERSONAL &POLICIAL"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11280
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CommandButton CmdActAdmin 
         BackColor       =   &H00C0C000&
         Caption         =   "&ACTUACIONES ADMINISTIVAS"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Image Image1 
         Height          =   2430
         Left            =   9120
         Picture         =   "MDISSAURXIX.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2355
      End
   End
   Begin VB.PictureBox LblPersonal 
      Align           =   1  'Align Top
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   3045
      Left            =   0
      ScaleHeight     =   3045
      ScaleWidth      =   20250
      TabIndex        =   0
      Top             =   0
      Width           =   20250
      Begin VB.CommandButton CmdCerrar 
         BackColor       =   &H00C0C000&
         Caption         =   "C&ERRAR"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   18840
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "SECCIÓN SUMARIOS ADMINISTRATIVOS"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6480
         TabIndex        =   1
         Top             =   840
         Width           =   7815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "UNIDAD REGIONAL XIX - DEPARTAMENTO VERA"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5760
         TabIndex        =   3
         Top             =   360
         Width           =   9255
      End
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "REGISTRO DE ACTUACIONES ADMINISTRATIVAS"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         TabIndex        =   2
         Top             =   1920
         Width           =   11655
      End
   End
   Begin VB.Menu ARCHIVO 
      Caption         =   "ARCHIVO"
      Begin VB.Menu Salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu FORMULARIOS 
      Caption         =   "FORMULARIOS"
      Begin VB.Menu ActuacionesAdministrativas 
         Caption         =   "Actuaciones Administrativas"
      End
      Begin VB.Menu EscalafónyJerarquías 
         Caption         =   "Escalafón y Jerarquías"
      End
      Begin VB.Menu PersonalPolicial 
         Caption         =   "Personal Policial"
      End
   End
   Begin VB.Menu CONSULTAS 
      Caption         =   "CONSULTAS"
      Begin VB.Menu ActVarias 
         Caption         =   "Actuaciones Varias"
      End
      Begin VB.Menu Reglamentación 
         Caption         =   "Reglamentación"
      End
      Begin VB.Menu ActGenerales 
         Caption         =   "Actuaciones en General"
      End
      Begin VB.Menu InfSumarias 
         Caption         =   "Informaciones Sumarias"
      End
      Begin VB.Menu SumAdministrativos 
         Caption         =   "Sumarios Administrativos"
      End
   End
   Begin VB.Menu AYUDA 
      Caption         =   "AYUDA"
      Begin VB.Menu Acerdade 
         Caption         =   "Acerda de..."
      End
   End
   Begin VB.Menu IMPRIMIR 
      Caption         =   "IMPRIMIR"
      Begin VB.Menu ActenGeneral 
         Caption         =   "Actuaciones en General"
      End
      Begin VB.Menu SumariosAdministrativos 
         Caption         =   "Sumarios Administrativos"
      End
      Begin VB.Menu InformacionesSumarias 
         Caption         =   "Informaciones Sumarias"
      End
   End
End
Attribute VB_Name = "MDISSAURXIX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Acerdade_Click()
FormAcercade.Show
End Sub

Private Sub ActVarias_Click()
FormActVarias.Show
End Sub

Private Sub SumariosAdministrativos_Click()
RptSumAdministrativos.Show
End Sub

Private Sub InformacionesSumarias_Click()
RptInfSumarias.Show
End Sub

Private Sub ActenGeneral_Click()
RptActuaciones.Show
End Sub

Private Sub ActGenerales_Click()
FormListadoGral.Show
End Sub

Private Sub InfSumarias_Click()
FormIS.Show
End Sub

Private Sub Leyes_Click()
FormPDF.Show
End Sub

Private Sub RSA_Click()
AcroPDF1.LoadFile "C:\SistemaSSA\sistemassaurxix\R.S.A.pdf"
End Sub

Private Sub Reglamentación_Click()
FormPDF.Show
End Sub

Private Sub Salir_Click()
If MsgBox("              ¿DESEA CERRAR EL SISTEMA?", vbQuestion + vbYesNo, "SALIR") = vbYes Then
   Unload Me
End If
End Sub

Private Sub SumAdministrativos_Click()
FormSA.Show
End Sub

Private Sub CmdCerrar_Click()
   If MsgBox("              ¿DESEA CERRAR EL SISTEMA?", vbQuestion + vbYesNo, "SALIR") = vbYes Then
    Unload Me
   End If
End Sub

Private Sub ActuacionesAdministrativas_Click()
FormActAdmin.Show
FormEscyJerarquia.Enabled = True
FormListadoGral.Enabled = True
FormPersonalPolicial.Enabled = True
End Sub

Private Sub CmdActAdmin_Click()
FormActAdmin.Show
End Sub

Private Sub CmdEscyJer_Click()
FormEscyJerarquia.Show
End Sub

Private Sub CmdPersonalPolicial_Click()
FormPersonalPolicial.Show
End Sub

Private Sub EscalafónyJerarquías_Click()
FormEscyJerarquia.Show
FormActAdmin.Enabled = True
FormListadoGral.Enabled = True
FormPersonalPolicial.Enabled = True
End Sub

Private Sub PersonalPolicial_Click()
FormPersonalPolicial.Show
FormActAdmin.Enabled = True
FormEscyJerarquia.Enabled = True
FormListadoGral.Enabled = True
End Sub
