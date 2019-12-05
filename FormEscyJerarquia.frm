VERSION 5.00
Begin VB.Form FormEscyJerarquia 
   BackColor       =   &H00FF8080&
   Caption         =   "ESCALAFÓN Y JERARQUÍA"
   ClientHeight    =   8970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   15360
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      TabIndex        =   13
      Top             =   240
      Width           =   375
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
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   12
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
      TabIndex        =   11
      Top             =   9840
      Width           =   2415
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
      TabIndex        =   10
      Top             =   9840
      Width           =   2415
   End
   Begin VB.Frame FrameCmd 
      BackColor       =   &H00FF8080&
      Height          =   3615
      Left            =   16200
      TabIndex        =   18
      Top             =   4800
      Width           =   3255
      Begin VB.CommandButton CmdCancelar 
         BackColor       =   &H000000FF&
         Caption         =   "CANCELAR MODIFICACIÓN"
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
         TabIndex        =   23
         Top             =   3000
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
         TabIndex        =   7
         Top             =   960
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
         TabIndex        =   6
         Top             =   360
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
         TabIndex        =   9
         Top             =   2280
         Width           =   2295
      End
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
         TabIndex        =   8
         Top             =   1680
         Width           =   2295
      End
   End
   Begin VB.Frame FrameBuscar 
      BackColor       =   &H00FF8080&
      Height          =   1335
      Left            =   16200
      TabIndex        =   17
      Top             =   3120
      Width           =   3255
      Begin VB.TextBox TxtBuscar 
         Alignment       =   2  'Center
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
         TabIndex        =   22
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
         TabIndex        =   21
         Top             =   240
         Width           =   1305
      End
   End
   Begin VB.Frame FrameDatos 
      BackColor       =   &H00FF8080&
      Height          =   6255
      Left            =   1680
      TabIndex        =   0
      Top             =   1560
      Width           =   13455
      Begin VB.TextBox Txtregistro 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Left            =   12840
         TabIndex        =   20
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox TxtJerarquia 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   3240
         Width           =   6255
      End
      Begin VB.TextBox TxtEscalafon 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   1
         Top             =   1080
         Width           =   6255
      End
      Begin VB.TextBox TxtSubescalafon 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   2
         Top             =   2160
         Width           =   6255
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
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5400
         Width           =   1575
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
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5400
         Width           =   1575
      End
      Begin VB.Label LblJerarquia 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "JERARQUÍA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         Top             =   3285
         Width           =   2055
      End
      Begin VB.Label LblEscalafon 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "ESCALAFÓN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   1140
         Width           =   2175
      End
      Begin VB.Label LblSubescalafon 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         Caption         =   "SUBESCALAFON"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Top             =   2205
         Width           =   2055
      End
   End
   Begin VB.Label LblEscyJerarquia 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ESCALAFÓN Y JERARQUÍA"
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
      TabIndex        =   19
      Top             =   600
      Width           =   13215
   End
End
Attribute VB_Name = "FormEscyJerarquia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CargarDatos()
TxtEscalafon.Text = db.rsESCALAFON.Fields!Escalafon.Value
TxtSubescalafon.Text = db.rsESCALAFON.Fields!Subescalafon.Value
TxtJerarquia.Text = db.rsESCALAFON.Fields!Jerarquia.Value
End Sub

Function BuscarJerarquia(x As Long) As Boolean
BuscarJerarquia = False
If db.rsESCALAFON.RecordCount > 0 Then
   db.rsESCALAFON.MoveFirst
   Do While Not db.rsESCALAFON.EOF

   If x = db.rsESCALAFON.Fields!Registro.Value Then
      BuscarJerarquia = True
   Exit Do
   Else
   db.rsESCALAFON.MoveNext
   End If
Loop

End If
End Function
Function CargarRegistro()
db.rsESCALAFON.MoveLast
Codigo = db.rsESCALAFON.Fields!Registro.Value
C = Codigo + 1
Txtregistro.Text = C
End Function

Private Sub CmdActAdmin_Click()
FormActAdmin.Show
End Sub

Private Sub CmdActualizar_Click()
If MsgBox("¿DESEA ACTUALIZAR LOS DATOS?", vbQuestion + vbYesNo, "ATENCION!") = vbYes Then
   db.rsESCALAFON.Fields!Escalafon.Value = TxtEscalafon.Text
   db.rsESCALAFON.Fields!Subescalafon.Value = TxtSubescalafon.Text
   db.rsESCALAFON.Fields!Jerarquia.Value = TxtJerarquia.Text

   db.rsESCALAFON.Update
   db.rsESCALAFON.Requery
End If

TxtEscalafon.Enabled = True
TxtSubescalafon.Enabled = True
TxtJerarquia.Enabled = True

TxtEscalafon.Text = ""
TxtSubescalafon.Text = ""
TxtJerarquia.Text = ""

CmdActualizar.Enabled = False
CmdAlta.Enabled = True
CmdBaja.Enabled = True
CmdBuscar.Enabled = True
CmdModificar.Enabled = True
CmdBorrar.Enabled = False
CmdRegistrar.Enabled = False
End Sub

Private Sub CmdAlta_Click()
   TxtEscalafon.Text = ""
   TxtSubescalafon.Text = ""
   TxtJerarquia.Text = ""
   
   TxtEscalafon.Enabled = True
   TxtSubescalafon.Enabled = True
   TxtJerarquia.Enabled = True
   
   TxtEscalafon.BackColor = &H80000005
   TxtSubescalafon.BackColor = &H80000005
   TxtJerarquia.BackColor = &H80000005

   TxtEscalafon.SetFocus

End Sub

Private Sub CmdBaja_Click()
If TxtJerarquia.Text = "" Then
   MsgBox "NO EXISTE EL REGISTRO"
Else
End If
   
   Pregunta = MsgBox("¿ELIMINAR REGISTRO?", vbQuestion + vbYesNo, "ATENCIÓN")

If Pregunta = vbYes Then
   db.rsESCALAFON.Delete
   db.rsESCALAFON.Requery
End If

If db.rsESCALAFON.EOF Then
   db.rsESCALAFON.MoveLast
End If

CargarDatos
Me.Refresh
End Sub

Private Sub CmdBuscar_Click()
If TxtBuscar.Text > 0 Then
   BuscarJerarquia (TxtBuscar.Text)
   If BuscarJerarquia(TxtBuscar.Text) = True Then
   CargarDatos
   End If
   
   If BuscarJerarquia(TxtBuscar.Text) = False Then
   MsgBox "JERARQUIA NO REGISTRADA"
   TxtBuscar.Text = ""
   TxtBuscar.SetFocus
   
   TxtEscalafon.Text = ""
   TxtSubescalafon.Text = ""
   TxtJerarquia.Text = ""
   End If
   
End If
   TxtEscalafon.Enabled = False
   TxtSubescalafon.Enabled = False
   TxtJerarquia.Enabled = False
End Sub

Private Sub CmdCancelar_Click()
   CmdRegistrar.Enabled = True
   CmdBorrar.Enabled = True
   CmdAlta.Enabled = True
   CmdBaja.Enabled = True
   CmdBuscar.Enabled = True
   Txtregistro.Enabled = True

   TxtEscalafon.Enabled = False
   TxtSubescalafon.Enabled = False
   TxtJerarquia.Enabled = False

   TxtEscalafon.BackColor = &HE0E0E0
   TxtSubescalafon.BackColor = &HE0E0E0
   TxtJerarquia.BackColor = &HE0E0E0
   TxtBuscar.BackColor = &HE0E0E0
End Sub

Private Sub CmdMDISSAURXIX_Click()
MDISSAURXIX.Show
End Sub

Private Sub CmdModificar_Click()
   TxtEscalafon.Enabled = True
   TxtSubescalafon.Enabled = True
   TxtJerarquia.Enabled = True
   
   TxtEscalafon.BackColor = &H80000005
   TxtSubescalafon.BackColor = &H80000005
   TxtJerarquia.BackColor = &H80000005
   
   CmdRegistrar.Enabled = False
   CmdBorrar.Enabled = False
   CmdAlta.Enabled = False
   CmdBaja.Enabled = False
   CmdBuscar.Enabled = False
   Txtregistro.Enabled = False

   TxtEscalafon.SetFocus
End Sub

Private Sub CmdPersonalPolicial_Click()
FormPersonalPolicial.Show
End Sub

Private Sub CmdRegistrar_Click()
If Not BuscarJerarquia(Txtregistro.Text) Then
       db.rsESCALAFON.AddNew
End If
If MsgBox("¿DESEA REGISTRAR LA JERARQUÍA?", vbQuestion + vbYesNo, "JERARQUÍA REGISTRADA") = vbYes Then
   db.rsESCALAFON.Fields!Escalafon.Value = TxtEscalafon.Text
   db.rsESCALAFON.Fields!Subescalafon.Value = TxtSubescalafon.Text
   db.rsESCALAFON.Fields!Jerarquia.Value = TxtJerarquia.Text

   db.rsESCALAFON.Update
   db.rsESCALAFON.Requery
End If

   TxtEscalafon.Text = ""
   TxtSubescalafon.Text = ""
   TxtJerarquia.Text = ""
   
   CargarRegistro

End Sub

Private Sub CmdSalir_Click()
    If MsgBox("              ¿SALIR?", vbQuestion + vbYesNo, "SALIR") = vbYes Then
    Unload Me
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
db.rsESCALAFON.Open
CargarRegistro
End Sub

