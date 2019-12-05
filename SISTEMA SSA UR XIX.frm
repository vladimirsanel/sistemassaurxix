VERSION 5.00
Begin VB.Form FormActAdmin 
   BackColor       =   &H00FF8080&
   Caption         =   "ACTUACIONES ADMINISTRATIVAS"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   6840
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
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   6840
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   6840
      Width           =   1335
   End
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
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "LISTADO DE SUMARIOS ADMINISTRATIVOS"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   17160
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   9360
      Width           =   2295
   End
   Begin VB.CommandButton CmdInfSumarias 
      BackColor       =   &H000000FF&
      Caption         =   "LISTADO DE INFORMACIONES SUMARIAS"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   17160
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8280
      Width           =   2295
   End
   Begin VB.OptionButton OptInfSumaria 
      BackColor       =   &H00FF8080&
      Caption         =   "INFORMACIÓN SUMARIA"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   11
      Top             =   7080
      Width           =   3135
   End
   Begin VB.OptionButton OptSumAdministrativo 
      BackColor       =   &H00FF8080&
      Caption         =   "SUMARIO ADMINISTRATIVO"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   10
      Top             =   6720
      Width           =   3615
   End
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
      Left            =   11160
      TabIndex        =   4
      Top             =   2520
      Width           =   5295
   End
   Begin VB.CommandButton CmdListadoAct 
      BackColor       =   &H000000FF&
      Caption         =   "LISTADO GENERAL DE ACTUACIONES"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   17160
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7200
      Width           =   2295
   End
   Begin VB.ComboBox CboAnio 
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
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   9600
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox TxtFechaIngreso 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3600
      TabIndex        =   2
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox TxtFechaHecho 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3600
      TabIndex        =   3
      Top             =   1920
      Width           =   2535
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
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6840
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
      Left            =   14880
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6840
      Width           =   1575
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
      TabIndex        =   24
      Top             =   9840
      Width           =   1815
   End
   Begin VB.Frame FrameBuscar 
      BackColor       =   &H00FF8080&
      Height          =   1335
      Left            =   16680
      TabIndex        =   41
      Top             =   1200
      Width           =   3255
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
         TabIndex        =   15
         Top             =   240
         Width           =   1305
      End
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
         TabIndex        =   14
         Top             =   600
         Width           =   2775
      End
   End
   Begin VB.Frame FrameCmd 
      BackColor       =   &H00FF8080&
      Height          =   3015
      Left            =   16680
      TabIndex        =   40
      Top             =   2640
      Width           =   3255
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   1680
         Width           =   2295
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   9840
      Width           =   1815
   End
   Begin VB.CommandButton CmdPersonalPolicial 
      BackColor       =   &H8000000D&
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
      Height          =   615
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   25
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
      TabIndex        =   29
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox TxtCarAdmin 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3600
      TabIndex        =   8
      Top             =   4320
      Width           =   12855
   End
   Begin VB.TextBox TxtCausante 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3600
      TabIndex        =   5
      Top             =   2520
      Width           =   6135
   End
   Begin VB.TextBox TxtNro 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   8400
      TabIndex        =   30
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox TxtCaratula 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3600
      TabIndex        =   7
      Top             =   3720
      Width           =   12855
   End
   Begin VB.TextBox TxtDestino 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3600
      TabIndex        =   6
      Top             =   3120
      Width           =   12855
   End
   Begin VB.TextBox TxtDescripcion 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   4920
      Width           =   12855
   End
   Begin VB.Label LblJerarquia 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "JERARQUÍA"
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
      Left            =   9720
      TabIndex        =   42
      Top             =   2520
      Width           =   1335
   End
   Begin VB.OLE OLE3 
      BackColor       =   &H00FFC0C0&
      DisplayType     =   1  'Icon
      Height          =   1095
      Left            =   840
      OleObjectBlob   =   "SISTEMA SSA UR XIX.frx":0000
      SourceDoc       =   "C:\SISTEMA SSA UR XIX 2017\R.S.A.pdf"
      TabIndex        =   28
      Top             =   9240
      Width           =   1455
   End
   Begin VB.OLE OLE1 
      BackColor       =   &H00FFC0C0&
      DisplayType     =   1  'Icon
      Height          =   1095
      Left            =   840
      OleObjectBlob   =   "SISTEMA SSA UR XIX.frx":1C18
      SourceDoc       =   "C:\SISTEMA SSA UR XIX 2017\DECRETO 0461-15.pdf"
      TabIndex        =   27
      Top             =   8040
      Width           =   1455
   End
   Begin VB.OLE OLE2 
      BackColor       =   &H00FFC0C0&
      DisplayType     =   1  'Icon
      Height          =   1095
      Left            =   840
      OleObjectBlob   =   "SISTEMA SSA UR XIX.frx":3830
      SourceDoc       =   "C:\SISTEMA SSA UR XIX 2017\Ley Nº 12.521-06.pdf"
      TabIndex        =   26
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label LblCarAdmin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "CARÁTULA ADMINISTRATIVA"
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
      Left            =   120
      TabIndex        =   39
      Top             =   4320
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   38
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label LblCausante 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "CAUSANTE"
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
      Left            =   1920
      TabIndex        =   37
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label LblNro 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "NÚMERO"
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
      Left            =   6960
      TabIndex        =   36
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label LblCaratula 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "CARÁTULA"
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
      Left            =   2160
      TabIndex        =   35
      Top             =   3720
      Width           =   1215
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
      Height          =   375
      Left            =   1680
      TabIndex        =   34
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label LblActuaciones 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ACTUACIONES ADMINISTRATIVAS"
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
      Left            =   4920
      TabIndex        =   33
      Top             =   240
      Width           =   9735
   End
   Begin VB.Label LblDescripcion 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION"
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
      Left            =   1200
      TabIndex        =   32
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label LblFechaIngreso 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA DE INGRESO"
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
      Left            =   840
      TabIndex        =   31
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label LblFechaHecho 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA DEL HECHO"
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
      TabIndex        =   0
      Top             =   1920
      Width           =   2655
   End
End
Attribute VB_Name = "FormActAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CargarDatos()
TxtNro.Text = db.rsACTUACIONES.Fields!Nro.Value
TxtFechaIngreso.Text = db.rsACTUACIONES.Fields!FechaIngreso.Value
TxtFechaHecho.Text = db.rsACTUACIONES.Fields!FechaHecho.Value
TxtCausante.Text = db.rsACTUACIONES.Fields!Causante.Value
TxtDestino.Text = db.rsACTUACIONES.Fields!Destino.Value
TxtCaratula.Text = db.rsACTUACIONES.Fields!Caratula.Value
TxtCarAdmin.Text = db.rsACTUACIONES.Fields!CarAdmin.Value
TxtDescripcion.Text = db.rsACTUACIONES.Fields!Descripcion.Value
End Sub

Function CargarNro()
db.rsACTUACIONES.MoveLast
Codigo = db.rsACTUACIONES.Fields!Nro.Value
C = Codigo + 1
TxtNro.Text = C
End Function

Function BuscarActuaciones(x As Long) As Boolean
BuscarActuaciones = False
If db.rsACTUACIONES.RecordCount > 0 Then
   db.rsACTUACIONES.MoveFirst
   Do While Not db.rsACTUACIONES.EOF

   If x = db.rsACTUACIONES.Fields!Nro.Value Then
      BuscarActuaciones = True
   Exit Do
   Else
   db.rsACTUACIONES.MoveNext
   End If
Loop

End If
End Function

Private Sub CmdArteIncisos_Click()
FormArteIncisos.Show
End Sub

Private Sub CmdActualizar_Click()
If MsgBox("¿DESEA ACTUALIZAR LOS DATOS?", vbQuestion + vbYesNo, "ATENCION!") = vbYes Then
   db.rsACTUACIONES.Fields!FechaIngreso.Value = TxtFechaIngreso.Text
   db.rsACTUACIONES.Fields!FechaHecho.Value = TxtFechaHecho.Text
   db.rsACTUACIONES.Fields!Causante.Value = TxtCausante.Text
   db.rsACTUACIONES.Fields!Destino.Value = TxtDestino.Text
   db.rsACTUACIONES.Fields!Caratula.Value = TxtCaratula.Text
   db.rsACTUACIONES.Fields!CarAdmin.Value = TxtCarAdmin.Text
   db.rsACTUACIONES.Fields!Descripcion.Value = TxtDescripcion.Text

   db.rsACTUACIONES.Requery
End If

TxtNro.Enabled = False
TxtFechaIngreso.Enabled = True
TxtFechaHecho.Enabled = True
TxtCausante.Enabled = True
TxtDestino.Enabled = True
TxtCaratula.Enabled = True
TxtCarAdmin.Enabled = True
TxtDescripcion.Enabled = True

CmdActualizar.Enabled = False
CmdAlta.Enabled = True
CmdBaja.Enabled = True
CmdBuscar.Enabled = True
CmdEliminar.Enabled = True
CmdModificar.Enabled = True
CmdBorrar.Enabled = False
CmdRegistrar.Enabled = False
End Sub

Private Sub CmdAlta_Click()
   CboJerarquia.Text = ""
   TxtFechaIngreso.Text = ""
   TxtFechaHecho.Text = ""
   TxtCausante.Text = ""
   TxtDestino.Text = ""
   TxtCaratula.Text = ""
   TxtCarAdmin.Text = ""
   TxtDescripcion.Text = ""
   TxtBuscar.Text = ""
   
   TxtNro.Enabled = False
   CboAnio.Enabled = True
   CboJerarquia.Enabled = True
   TxtFechaIngreso.Enabled = True
   TxtFechaHecho.Enabled = True
   TxtCausante.Enabled = True
   TxtDestino.Enabled = True
   TxtCaratula.Enabled = True
   TxtCarAdmin.Enabled = True
   TxtDescripcion.Enabled = True
   
   CboAnio.BackColor = &H80000005
   CboJerarquia.BackColor = &H80000005
   TxtFechaIngreso.BackColor = &H80000005
   TxtFechaHecho.BackColor = &H80000005
   TxtCausante.BackColor = &H80000005
   TxtDestino.BackColor = &H80000005
   TxtCaratula.BackColor = &H80000005
   TxtCarAdmin.BackColor = &H80000005
   TxtDescripcion.BackColor = &H80000005
   
   CargarNro
End Sub

Private Sub CmdAnterior_Click()
db.rsACTUACIONES.MovePrevious
If db.rsACTUACIONES.BOF Then
   db.rsACTUACIONES.MoveFirst
End If
CargarDatos
End Sub

Private Sub CmdBaja_Click()
If TxtNro.Text = "" Then
   MsgBox "NO EXISTE EL REGISTRO"
Else

End If
   Pregunta = MsgBox("¿ELIMINAR REGISTRO?", vbQuestion + vbYesNo, "ATENCIÓN")

If Pregunta = vbYes Then
   db.rsACTUACIONES.Delete
   db.rsACTUACIONES.Requery
End If

If db.rsACTUACIONES.EOF Then
   db.rsACTUACIONES.MoveLast
End If

CargarDatos
Me.Refresh
End Sub


Private Sub CmdBorrar_Click()
   TxtNro.Text = ""
   CboAnio.Text = ""
   CboJerarquia.Text = ""
   TxtFechaIngreso.Text = ""
   TxtFechaHecho.Text = ""
   TxtCausante.Text = ""
   TxtDestino.Text = ""
   TxtCaratula.Text = ""
   TxtCarAdmin.Text = ""
   TxtDescripcion.Text = ""
   TxtBuscar.Text = ""
   
   TxtFechaIngreso.SetFocus
End Sub

Private Sub CmdBuscar_Click()
If TxtBuscar.Text > 0 Then
   BuscarActuaciones (TxtBuscar.Text)
   If BuscarActuaciones(TxtBuscar.Text) = True Then
   CargarDatos
   End If
   
   If BuscarActuaciones(TxtBuscar.Text) = False Then
   MsgBox "ACTUACIÓN NO REGISTRADA"
   TxtBuscar.Text = ""
   TxtBuscar.SetFocus
   
   TxtNro.Text = ""
   CboAnio.Text = ""
   CboJerarquia.Text = ""
   TxtFechaIngreso.Text = ""
   TxtFechaHecho.Text = ""
   TxtCausante.Text = ""
   TxtDestino.Text = ""
   TxtCaratula.Text = ""
   TxtCarAdmin.Text = ""
   TxtDescripcion.Text = ""
   TxtBuscar.Text = ""
   End If
   
End If
   TxtNro.Enabled = False
   CboAnio.Enabled = False
   CboJerarquia.Enabled = False
   TxtFechaIngreso.Enabled = False
   TxtFechaHecho.Enabled = False
   TxtCausante.Enabled = False
   TxtDestino.Enabled = False
   TxtCaratula.Enabled = False
   TxtCarAdmin.Enabled = False
   TxtDescripcion.Enabled = False
End Sub

Private Sub CmdEscyJerarquias_Click()
FormEscyJerarquia.Show
End Sub

Private Sub CmdListadoAct_Click()
RptActuaciones.Show
End Sub

Private Sub CmdMDISSAURXIX_Click()
MDISSAURXIX.Show
End Sub

Private Sub CmdModificar_Click()
   CboJerarquia.Enabled = True
   TxtFechaIngreso.Enabled = True
   TxtFechaHecho.Enabled = True
   TxtCausante.Enabled = True
   TxtDestino.Enabled = True
   TxtCaratula.Enabled = True
   TxtCarAdmin.Enabled = True
   TxtDescripcion.Enabled = True
   
   CboJerarquia.BackColor = &H80000005
   TxtFechaIngreso.BackColor = &H80000005
   TxtFechaHecho.BackColor = &H80000005
   TxtCausante.BackColor = &H80000005
   TxtDestino.BackColor = &H80000005
   TxtCaratula.BackColor = &H80000005
   TxtCarAdmin.BackColor = &H80000005
   TxtDescripcion.BackColor = &H80000005
End Sub

Private Sub CmdPersonalPolicial_Click()
FormPersonalPolicial.Show
End Sub

Private Sub CmdPrimero_Click()
db.rsACTUACIONES.MoveFirst
CargarDatos
End Sub

Private Sub CmdRegistrar_Click()
 If Not BuscarActuaciones(TxtNro.Text) Then
        db.rsACTUACIONES.AddNew
 End If
If MsgBox("¿DESEA REGISTRAR NUEVA ACTUACIÓN?", vbQuestion + vbYesNo, "ACTUACIÓN REGISTRADA") = vbYes Then
   db.rsACTUACIONES.Fields!Nro.Value = TxtNro.Text
   db.rsACTUACIONES.Fields!Fecha.Value = CboAnio.Text
   db.rsACTUACIONES.Fields!Jerarquia.Value = CboJerarquia.Text
   db.rsACTUACIONES.Fields!FechaIngreso.Value = TxtFechaIngreso.Text
   db.rsACTUACIONES.Fields!FechaHecho.Value = TxtFechaHecho.Text
   db.rsACTUACIONES.Fields!Causante.Value = TxtCausante.Text
   db.rsACTUACIONES.Fields!Destino.Value = TxtDestino.Text
   db.rsACTUACIONES.Fields!Caratula.Value = TxtCaratula.Text
   db.rsACTUACIONES.Fields!CarAdmin.Value = TxtCarAdmin.Text
   db.rsACTUACIONES.Fields!Descripcion.Value = TxtDescripcion.Text

   db.rsACTUACIONES.Update
   db.rsACTUACIONES.Requery
End If

   CboAnio.Text = ""
   CboJerarquia.Text = ""
   TxtFechaIngreso.Text = ""
   TxtFechaHecho.Text = ""
   TxtCausante.Text = ""
   TxtDestino.Text = ""
   TxtCaratula.Text = ""
   TxtCarAdmin.Text = ""
   TxtDescripcion.Text = ""
   TxtBuscar.Text = ""

   CargarNro

End Sub


Private Sub CmdSalir_Click()
    If MsgBox("              ¿SALIR?", vbQuestion + vbYesNo, "SALIR") = vbYes Then
    Unload Me
    End If
End Sub

Private Sub CmdSiguiente_Click()
On Error Resume Next
db.rsACTUACIONES.MoveNext
If db.rsACTUACIONES.EOF Then
   db.rsACTUACIONES.MoveLast
End If

CargarDatos
End Sub

Private Sub CmdUltimo_Click()
On Error Resume Next
db.rsACTUACIONES.MoveNext
CargarDatos

End Sub

Private Sub Form_Load()

db.rsESCALAFON.Open
Do While Not db.rsESCALAFON.EOF
   CboJerarquia.AddItem (db.rsESCALAFON.Fields!Jerarquia.Value)
   db.rsESCALAFON.MoveNext
Loop

db.rsACTUACIONES.Open
CargarNro

CboAnio.AddItem "2000"
CboAnio.AddItem "2001"
CboAnio.AddItem "2002"
CboAnio.AddItem "2003"
CboAnio.AddItem "2004"
CboAnio.AddItem "2005"
CboAnio.AddItem "2006"
CboAnio.AddItem "2007"
CboAnio.AddItem "2008"
CboAnio.AddItem "2009"
CboAnio.AddItem "2010"
CboAnio.AddItem "2011"
CboAnio.AddItem "2012"
CboAnio.AddItem "2013"
CboAnio.AddItem "2014"
CboAnio.AddItem "2015"
CboAnio.AddItem "2016"
CboAnio.AddItem "2017"
CboAnio.AddItem "2018"
CboAnio.AddItem "2019"
CboAnio.AddItem "2020"
CboAnio.AddItem "2021"
CboAnio.AddItem "2022"
CboAnio.AddItem "2023"
CboAnio.AddItem "2024"
CboAnio.AddItem "2025"
CboAnio.AddItem "2026"
CboAnio.AddItem "2027"
CboAnio.AddItem "2028"
CboAnio.AddItem "2029"
CboAnio.AddItem "2030"
CboAnio.AddItem "2031"
CboAnio.AddItem "2032"
CboAnio.AddItem "2033"
CboAnio.AddItem "2034"
CboAnio.AddItem "2035"
CboAnio.AddItem "2036"
CboAnio.AddItem "2037"
CboAnio.AddItem "2038"
CboAnio.AddItem "2039"
CboAnio.AddItem "2040"
CboAnio.AddItem "2041"
CboAnio.AddItem "2042"
CboAnio.AddItem "2043"
CboAnio.AddItem "2044"
CboAnio.AddItem "2045"
CboAnio.AddItem "2046"
CboAnio.AddItem "2047"
CboAnio.AddItem "2048"
CboAnio.AddItem "2049"
CboAnio.AddItem "2050"
End Sub

