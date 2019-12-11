VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormPersonalPolicial 
   BackColor       =   &H00FF8080&
   ClientHeight    =   8955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15300
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   15300
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdPrimero 
      BackColor       =   &H00FFC0C0&
      Caption         =   "PRIMERO"
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
      Left            =   16440
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton CmdAnterior 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ANTERIOR"
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
      Left            =   16440
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   8640
      Width           =   1335
   End
   Begin VB.CommandButton CmdSiguiente 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SIGUIENTE"
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
      Left            =   17880
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8640
      Width           =   1335
   End
   Begin VB.CommandButton CmdUltimo 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ÚLTIMO"
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
      Left            =   17880
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Frame FrameConsulta 
      BackColor       =   &H00FF8080&
      Height          =   3615
      Left            =   1680
      TabIndex        =   8
      Top             =   5760
      Width           =   13455
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FormPersonal.frx":0000
         Height          =   2535
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   4471
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BackColor       =   16776960
         BorderStyle     =   0
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   17
         RowDividerStyle =   6
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataMember      =   "PERSONAL"
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "jerarquia"
            Caption         =   "            JERARQUÍA"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "ni"
            Caption         =   "    N.I."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "apeNom"
            Caption         =   "              APELLIDO Y NOMBRE"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "destino"
            Caption         =   "                 DESTINO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "dni"
            Caption         =   "          D.N.I."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "obs"
            Caption         =   "              OBSERVACIONES"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            ScrollBars      =   1
            SizeMode        =   1
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   2280,189
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   705,26
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   3165,166
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   2415,118
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   1379,906
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   2954,835
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton CmdEscyJerarquias 
      BackColor       =   &H8000000D&
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
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   9840
      Width           =   1815
   End
   Begin VB.Frame FrameBuscar 
      BackColor       =   &H00FF8080&
      Height          =   1335
      Left            =   16200
      TabIndex        =   26
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   240
         Width           =   1305
      End
   End
   Begin VB.Frame FrameCmd 
      BackColor       =   &H00FF8080&
      Height          =   3015
      Left            =   16200
      TabIndex        =   25
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
         TabIndex        =   13
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
         TabIndex        =   14
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
         TabIndex        =   11
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
         TabIndex        =   12
         Top             =   960
         Width           =   2295
      End
   End
   Begin VB.Frame FrameDatos 
      BackColor       =   &H00FF8080&
      Height          =   4455
      Left            =   1680
      TabIndex        =   19
      Top             =   1200
      Width           =   13455
      Begin VB.ComboBox CboJerarquia 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Malgun Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2400
         TabIndex        =   28
         Top             =   720
         Width           =   6855
      End
      Begin VB.TextBox TxtDni 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
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
         TabIndex        =   4
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
         TabIndex        =   6
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
         TabIndex        =   7
         Top             =   3960
         Width           =   1575
      End
      Begin VB.TextBox TxtapeNom 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
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
         TabIndex        =   2
         Top             =   1440
         Width           =   10215
      End
      Begin VB.TextBox TxtNi 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
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
         TabIndex        =   1
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox TxtDestino 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
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
         Top             =   2280
         Width           =   6855
      End
      Begin VB.TextBox TxtObs 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
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
         TabIndex        =   5
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
         TabIndex        =   27
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label LblapeNom 
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   3120
         Width           =   1815
      End
   End
   Begin VB.CommandButton CmdMDISSAURXIX 
      BackColor       =   &H8000000D&
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
      Height          =   615
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9840
      Width           =   1815
   End
   Begin VB.CommandButton CmdActAdmin 
      BackColor       =   &H8000000D&
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
      Height          =   615
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   9840
      Width           =   1815
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
      TabIndex        =   18
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
      Top             =   240
      Width           =   13215
   End
End
Attribute VB_Name = "FormPersonalPolicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CargarDatos()
CboJerarquia.Text = db.rsPERSONAL.Fields!Jerarquia.Value
TxtNi.Text = db.rsPERSONAL.Fields!Ni.Value
TxtapeNom.Text = db.rsPERSONAL.Fields!apeNom.Value
TxtDestino.Text = db.rsPERSONAL.Fields!destino.Value
TxtDni.Text = db.rsPERSONAL.Fields!Dni.Value
TxtObs.Text = db.rsPERSONAL.Fields!Obs.Value
End Sub

Function BuscarNI(X As Long) As Boolean
BuscarNI = False
If db.rsPERSONAL.RecordCount > 0 Then
   db.rsPERSONAL.MoveFirst
   Do While Not db.rsPERSONAL.EOF

   If X = db.rsPERSONAL.Fields!Ni.Value Then
      BuscarNI = True
   Exit Do
   Else
   db.rsPERSONAL.MoveNext
   End If
Loop

End If
End Function



Private Sub CmdActAdmin_Click()
   FormActAdmin.Show
End Sub


Private Sub CmdActualizar_Click()
If MsgBox("¿DESEA ACTUALIZAR LOS DATOS?", vbQuestion + vbYesNo, "ATENCION!") = vbYes Then
   db.rsPERSONAL.Fields!Jerarquia.Value = CboJerarquia.Text
   db.rsPERSONAL.Fields!Ni.Value = TxtNi.Text
   db.rsPERSONAL.Fields!apeNom.Value = TxtapeNom.Text
   db.rsPERSONAL.Fields!Dni.Value = TxtDni.Text
   db.rsPERSONAL.Fields!destino.Value = CboDestino.Text
   db.rsPERSONAL.Fields!Obs.Value = TxtObs.Text

   db.rsPERSONAL.Update
   db.rsPERSONAL.Requery
End If

CboJerarquia.Enabled = False
TxtNi.Enabled = False
TxtapeNom.Enabled = False
TxtDestino.Enabled = False
TxtDni.Enabled = False
TxtObs.Enabled = False

CmdAlta.Enabled = True
CmdBaja.Enabled = True
CmdBuscar.Enabled = True
CmdModificar.Enabled = True
CmdActualizar.Enabled = False
CmdBorrar.Enabled = False
CmdRegistrar.Enabled = False
End Sub

'-------------ALTA-------------'
Private Sub CmdAlta_Click()
   CmdBorrar.Enabled = True
   CmdRegistrar.Enabled = True
   CmdModificar.Enabled = False
   CmdActualizar.Enabled = False
   CmdBaja.Enabled = False
   TxtBuscar.Enabled = True
   CmdBuscar.Enabled = True
   FrameDatos.Enabled = True
   CboJerarquia.Enabled = True
   TxtNi.Enabled = True
   TxtapeNom.Enabled = True
   TxtDestino.Enabled = True
   TxtDni.Enabled = True
   TxtObs.Enabled = True

   CboJerarquia.Text = ""
   TxtNi.Text = ""
   TxtapeNom.Text = ""
   TxtDestino.Text = ""
   TxtDni.Text = ""
   TxtObs.Text = ""
   TxtBuscar.Text = ""

   CboJerarquia.BackColor = &H80000005
   TxtNi.BackColor = &H80000005
   TxtapeNom.BackColor = &H80000005
   TxtDestino.BackColor = &H80000005
   TxtDni.BackColor = &H80000005
   TxtObs.BackColor = &H80000005
   TxtBuscar.BackColor = &H80000005

End Sub


Private Sub CmdArteIncisos_Click()
FormArteIncisos.Show
End Sub

Private Sub CmdBaja_Click()
If TxtNro.Text = "" Then
   MsgBox "NO EXISTE EL REGISTRO"
Else

End If
   Pregunta = MsgBox("¿ELIMINAR REGISTRO?", vbQuestion + vbYesNo, "ATENCIÓN")

If Pregunta = vbYes Then
   db.rsPERSONAL.Delete
   db.rsPERSONAL.Requery
End If

If db.rsPERSONAL.EOF Then
   db.rsPERSONAL.MoveLast
End If

CargarDatos
Me.Refresh

End Sub

'-----------BORRAR-----------'
Private Sub CmdBorrar_Click()
If MsgBox("¿DESEA BORRAR Y CANCELAR EL INGRESO DE DATOS?", vbQuestion + vbYesNo, "ATENCION!") = vbYes Then
   CmdBorrar.Enabled = True
   CmdRegistrar.Enabled = True
   CmdAlta.Enabled = True
   CmdModificar.Enabled = True
   CmdActualizar.Enabled = True
   CmdBaja.Enabled = True
   TxtBuscar.Enabled = True
   CmdBuscar.Enabled = True
   FrameDatos.Enabled = True

   CboJerarquia.Text = ""
   TxtNi.Text = ""
   TxtapeNom.Text = ""
   TxtDestino.Text = ""
   TxtDni.Text = ""
   TxtObs.Text = ""
   
   CmdAlta.SetFocus
End If
End Sub

Private Sub CmdBuscar_Click()
If TxtBuscar.Text > 0 Then
   BuscarNI (TxtBuscar.Text)
   If BuscarNI(TxtBuscar.Text) = True Then
   CargarDatos
   End If
   
   If BuscarNI(TxtBuscar.Text) = False Then
   MsgBox "PERSONAL NO REGISTRADO"
   TxtBuscar.Text = ""
   TxtBuscar.SetFocus
   
   CboJerarquia.Text = ""
   TxtNi.Text = ""
   TxtapeNom.Text = ""
   TxtDestino.Text = ""
   TxtDni.Text = ""
   TxtObs.Text = ""
   End If
   

   CboJerarquia.Enabled = False
   TxtNi.Enabled = False
   TxtapeNom.Enabled = False
   TxtDestino.Enabled = False
   TxtDni.Enabled = False
   TxtObs.Enabled = False
   CmdBaja.Enabled = True
   CmdModificar.Enabled = True
   CmdActualizar.Enabled = True
End If

End Sub

Private Sub CmdEscyJerarquias_Click()
FormEscyJerarquia.Show
End Sub

Private Sub CmdMDISSAURXIX_Click()
MDISSAURXIX.Show
End Sub

Private Sub CmdModificar_Click()
   CboJerarquia.Enabled = True
   TxtNi.Enabled = True
   TxtapeNom.Enabled = True
   TxtDestino.Enabled = True
   TxtDni.Enabled = True
   TxtObs.Enabled = True
   CmdAlta.Enabled = False
   CmdBaja.Enabled = False
   CmdModificar.Enabled = False
   CmdActualizar.Enabled = True
   CmdRegistrar.Enabled = False
   CmdBorrar.Enabled = True
   FrameBuscar.Enabled = False
   
   CboJerarquia.BackColor = &H80000005
   TxtNi.BackColor = &H80000005
   TxtapeNom.BackColor = &H80000005
   TxtDestino.BackColor = &H80000005
   TxtDni.BackColor = &H80000005
   TxtObs.BackColor = &H80000005
   TxtBuscar.BackColor = &H80000005
End Sub

Private Sub CmdPrimero_Click()
db.rsPERSONAL.MoveFirst
CargarDatos
End Sub

Private Sub CmdUltimo_Click()
db.rsPERSONAL.MoveLast
CargarDatos
End Sub

Private Sub CmdAnterior_Click()
db.rsPERSONAL.MovePrevious
If db.rsPERSONAL.BOF Then
   db.rsPERSONAL.MoveFirst
End If
CargarDatos
End Sub

Private Sub CmdSiguiente_Click()
db.rsPERSONAL.MoveNext
If db.rsPERSONAL.EOF Then
   db.rsPERSONAL.MoveLast
End If
CargarDatos
End Sub

Private Sub CmdRegistrar_Click()
 If Not BuscarNI(TxtNi.Text) Then
        db.rsPERSONAL.AddNew
 End If
If MsgBox("¿DESEA REGISTRAR NUEVO PERSONAL?", vbQuestion + vbYesNo, "PERSONAL REGISTRADO") = vbYes Then
   db.rsPERSONAL.Fields!Jerarquia.Value = CboJerarquia.Text
   db.rsPERSONAL.Fields!Ni.Value = TxtNi.Text
   db.rsPERSONAL.Fields!apeNom.Value = TxtapeNom.Text
   db.rsPERSONAL.Fields!destino.Value = TxtDestino.Text
   db.rsPERSONAL.Fields!Dni.Value = TxtDni.Text
   db.rsPERSONAL.Fields!Obs.Value = TxtObs.Text

   db.rsPERSONAL.Update
   db.rsPERSONAL.Requery
End If

End Sub

Private Sub CmdSalir_Click()
    If MsgBox("               ¿SALIR?", vbQuestion + vbYesNo, "SALIR") = vbYes Then
    Unload Me
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
db.rsPERSONAL.Open
db.rsESCALAFON.Open

Do While Not db.rsESCALAFON.EOF
   CboJerarquia.AddItem (db.rsESCALAFON.Fields!Jerarquia.Value)
   db.rsESCALAFON.MoveNext
Loop
End Sub

