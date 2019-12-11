VERSION 5.00
Begin VB.Form FormActAdmin 
   BackColor       =   &H00FF8080&
   Caption         =   "ACTUACIONES ADMINISTRATIVAS"
   ClientHeight    =   8025
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10635
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   1335
      Left            =   16680
      TabIndex        =   44
      Top             =   1440
      Width           =   3255
      Begin VB.TextBox TxtBuscarTexto 
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
         TabIndex        =   46
         Top             =   600
         Width           =   2775
      End
      Begin VB.CommandButton CmdBuscarTexto 
         BackColor       =   &H00FFC0C0&
         Caption         =   "BUSCAR POR TEXTO"
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   240
         Width           =   2505
      End
   End
   Begin VB.ComboBox CboNi 
      BackColor       =   &H00E0E0E0&
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
      Left            =   13920
      TabIndex        =   43
      Top             =   2520
      Width           =   2535
   End
   Begin VB.ComboBox CboTipoRegistro 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3600
      TabIndex        =   9
      Top             =   7560
      Width           =   3495
   End
   Begin VB.TextBox TxtAnio 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   9600
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.ComboBox CboCausante 
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
      ItemData        =   "SISTEMA SSA UR XIX.frx":0000
      Left            =   3600
      List            =   "SISTEMA SSA UR XIX.frx":0002
      TabIndex        =   4
      Top             =   2520
      Width           =   9615
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
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7560
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
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8040
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
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8040
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
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7560
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
      TabIndex        =   26
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
      TabIndex        =   25
      Top             =   8280
      Width           =   2295
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
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7560
      Width           =   1335
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
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7560
      Width           =   1335
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   9600
      Width           =   1815
   End
   Begin VB.Frame FrameBuscar 
      BackColor       =   &H00FF8080&
      Height          =   1335
      Left            =   16680
      TabIndex        =   39
      Top             =   2880
      Width           =   3255
      Begin VB.CommandButton CmdBuscar 
         BackColor       =   &H00FFC0C0&
         Caption         =   "BUSCAR POR N°"
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   2505
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
         TabIndex        =   16
         Top             =   600
         Width           =   2775
      End
   End
   Begin VB.Frame FrameCmd 
      BackColor       =   &H00FF8080&
      Height          =   3015
      Left            =   16680
      TabIndex        =   38
      Top             =   4320
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   21
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
         TabIndex        =   20
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   9600
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
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   9600
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
      TabIndex        =   27
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
      TabIndex        =   7
      Top             =   4320
      Width           =   12855
   End
   Begin VB.TextBox TxtNro 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   8400
      TabIndex        =   28
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      Height          =   2415
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   4920
      Width           =   12855
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
      Left            =   1680
      TabIndex        =   42
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label2 
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
      Height          =   495
      Left            =   4560
      TabIndex        =   41
      Top             =   2400
      Width           =   15
   End
   Begin VB.Label LblTipoRegistro 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO DE REGISTRO"
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
      Left            =   1080
      TabIndex        =   40
      Top             =   7560
      Width           =   2295
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
      TabIndex        =   37
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
      TabIndex        =   36
      Top             =   1200
      Width           =   255
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
      Height          =   375
      Left            =   12360
      TabIndex        =   35
      Top             =   2640
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
      TabIndex        =   34
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label LblCaratula 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "CARÁTULA JUDICIAL"
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
      Left            =   1080
      TabIndex        =   33
      Top             =   3720
      Width           =   2295
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
      TabIndex        =   32
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
      TabIndex        =   31
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
      TabIndex        =   30
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
      TabIndex        =   29
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
   Begin VB.Menu ARCHIVO 
      Caption         =   "ARCHIVO"
      Begin VB.Menu Salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "FormActAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CargarDatos()
TxtNro.Text = db.rsACTUACIONES.Fields!idActuaciones.Value
TxtFechaIngreso.Text = db.rsACTUACIONES.Fields!FechaIngreso.Value
TxtFechaHecho.Text = db.rsACTUACIONES.Fields!FechaHecho.Value
TxtAnio.Text = db.rsACTUACIONES.Fields!anio.Value
CboCausante.Text = db.rsACTUACIONES.Fields!causante.Value
CboTipoRegistro.Text = db.rsACTUACIONES.Fields!TipoRegistro.Value
TxtDestino.Text = db.rsACTUACIONES.Fields!destino.Value
TxtCaratula.Text = db.rsACTUACIONES.Fields!caratula.Value
TxtCarAdmin.Text = db.rsACTUACIONES.Fields!carAdmin.Value
TxtDescripcion.Text = db.rsACTUACIONES.Fields!descripcion.Value
End Sub

Function CargarNro()
On Error Resume Next
db.rsACTUACIONES.MoveLast
Codigo = db.rsACTUACIONES.Fields!idActuaciones.Value
C = Codigo + 1
TxtNro.Text = C
End Function

Function BuscarActuaciones(X As Long) As Boolean
BuscarActuaciones = False
If db.rsACTUACIONES.RecordCount > 0 Then
   db.rsACTUACIONES.MoveFirst
   Do While Not db.rsACTUACIONES.EOF

   If X = db.rsACTUACIONES.Fields!idActuaciones.Value Then
      BuscarActuaciones = True
   Exit Do
   Else
   db.rsACTUACIONES.MoveNext
   End If
Loop

End If
End Function

Function BuscarTexto(X As String)
BuscarTexto = False
If db.rsACTUACIONES.Source > "" Then
   db.rsACTUACIONES.MoveFirst
   Do While Not db.rsACTUACIONES.EOF

   If X = db.rsACTUACIONES.Fields!causante.Value Then
      BuscarTexto = True
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
   db.rsACTUACIONES.Fields!anio.Value = TxtAnio.Text
   db.rsACTUACIONES.Fields!causante.Value = CboCausante.Text
   db.rsACTUACIONES.Fields!TipoRegistro.Value = CboTipoRegistro.Text
   db.rsACTUACIONES.Fields!destino.Value = TxtDestino.Text
   db.rsACTUACIONES.Fields!caratula.Value = TxtCaratula.Text
   db.rsACTUACIONES.Fields!carAdmin.Value = TxtCarAdmin.Text
   db.rsACTUACIONES.Fields!descripcion.Value = TxtDescripcion.Text

    db.rsACTUACIONES.Update
   db.rsACTUACIONES.Requery
End If

CboCausante.Enabled = False
CboTipoRegistro.Enabled = False
TxtFechaIngreso.Enabled = False
TxtFechaHecho.Enabled = False
TxtAnio.Enabled = False
TxtDestino.Enabled = False
TxtCaratula.Enabled = False
TxtCarAdmin.Enabled = False
TxtDescripcion.Enabled = False

CmdAlta.Enabled = True
CmdBaja.Enabled = True
CmdBuscar.Enabled = True
CmdModificar.Enabled = True
CmdActualizar.Enabled = False
CmdBorrar.Enabled = False
CmdRegistrar.Enabled = False
End Sub

Private Sub CmdAlta_Click()
   TxtFechaIngreso.Text = ""
   TxtFechaHecho.Text = ""
   CboCausante.Text = ""
   CboNi.Text = ""
   CboTipoRegistro.Text = ""
   TxtDestino.Text = ""
   TxtCaratula.Text = ""
   TxtCarAdmin.Text = ""
   TxtDescripcion.Text = ""
   TxtBuscar.Text = ""

   CboNi.Enabled = True
   CboCausante.Enabled = True
   CboTipoRegistro.Enabled = True
   TxtAnio.Enabled = True
   TxtFechaIngreso.Enabled = True
   TxtFechaHecho.Enabled = True
   TxtDestino.Enabled = True
   TxtCaratula.Enabled = True
   TxtCarAdmin.Enabled = True
   TxtDescripcion.Enabled = True
   TxtBuscar.Enabled = True
   CmdModificar.Enabled = False
   CmdActualizar.Enabled = False
   
   TxtAnio.BackColor = &H80000005
   CboNi.BackColor = &H80000005
   CboCausante.BackColor = &H80000005
   CboTipoRegistro.BackColor = &H80000005
   TxtFechaIngreso.BackColor = &H80000005
   TxtFechaHecho.BackColor = &H80000005
   TxtDestino.BackColor = &H80000005
   TxtCaratula.BackColor = &H80000005
   TxtCarAdmin.BackColor = &H80000005
   TxtDescripcion.BackColor = &H80000005
   
   CargarNro
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
   TxtAnio.Text = ""
   CboCausante.Text = ""
   TxtFechaIngreso.Text = ""
   TxtFechaHecho.Text = ""
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
   
      TxtAnio.Text = ""
      CboCausante.Text = ""
      CboTipoRegistro.Text = ""
      TxtFechaIngreso.Text = ""
      TxtFechaHecho.Text = ""
      TxtDestino.Text = ""
      TxtCaratula.Text = ""
      TxtCarAdmin.Text = ""
      TxtDescripcion.Text = ""
      TxtBuscar.Text = ""
      TxtBuscarTexto.Text = ""
   End If
End If
End Sub

Private Sub CmdBuscarTexto_Click()
If TxtBuscarTexto.Text > "" Then
   BuscarTexto (TxtBuscarTexto.Text)
   If BuscarTexto(TxtBuscarTexto.Text) = True Then
      CargarDatos
   End If
   
   If BuscarTexto(TxtBuscarTexto.Text) = False Then
      MsgBox "ACTUACIÓN NO REGISTRADA"
      TxtBuscarTexto.Text = ""
      TxtBuscarTexto.SetFocus
   
      TxtAnio.Text = ""
      CboCausante.Text = ""
      CboTipoRegistro.Text = ""
      TxtFechaIngreso.Text = ""
      TxtFechaHecho.Text = ""
      TxtDestino.Text = ""
      TxtCaratula.Text = ""
      TxtCarAdmin.Text = ""
      TxtDescripcion.Text = ""
      TxtBuscar.Text = ""
      TxtBuscarTexto.Text = ""
   End If
End If
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
   CboCausante.Enabled = True
   CboTipoRegistro.Enabled = True
   TxtAnio.Enabled = True
   TxtFechaIngreso.Enabled = True
   TxtFechaHecho.Enabled = True
   TxtDestino.Enabled = True
   TxtCaratula.Enabled = True
   TxtCarAdmin.Enabled = True
   TxtDescripcion.Enabled = True
   CmdActualizar.Enabled = True
   CmdRegistrar.Enabled = False
   CmdAlta.Enabled = False
   CmdBaja.Enabled = False
   
   CboCausante.BackColor = &H80000005
   CboTipoRegistro.BackColor = &H80000005
   TxtFechaIngreso.BackColor = &H80000005
   TxtAnio.BackColor = &H80000005
   TxtFechaHecho.BackColor = &H80000005
   TxtDestino.BackColor = &H80000005
   TxtCaratula.BackColor = &H80000005
   TxtCarAdmin.BackColor = &H80000005
   TxtDescripcion.BackColor = &H80000005
End Sub

Private Sub CmdPersonalPolicial_Click()
FormPersonalPolicial.Show
End Sub

Private Sub CmdRegistrar_Click()
 If Not BuscarActuaciones(TxtNro.Text) Then
        db.rsACTUACIONES.AddNew
 End If
If MsgBox("¿DESEA REGISTRAR NUEVA ACTUACIÓN?", vbQuestion + vbYesNo, "ACTUACIÓN REGISTRADA") = vbYes Then
   db.rsACTUACIONES.Fields!idActuaciones.Value = TxtNro.Text
   db.rsACTUACIONES.Fields!causante.Value = CboCausante.Text
   db.rsACTUACIONES.Fields!anio.Value = TxtAnio.Text
   db.rsACTUACIONES.Fields!FechaIngreso.Value = TxtFechaIngreso.Text
   db.rsACTUACIONES.Fields!FechaHecho.Value = TxtFechaHecho.Text
   db.rsACTUACIONES.Fields!destino.Value = TxtDestino.Text
   db.rsACTUACIONES.Fields!caratula.Value = TxtCaratula.Text
   db.rsACTUACIONES.Fields!carAdmin.Value = TxtCarAdmin.Text
   db.rsACTUACIONES.Fields!descripcion.Value = TxtDescripcion.Text
   db.rsACTUACIONES.Fields!TipoRegistro.Value = CboTipoRegistro.Text

   db.rsACTUACIONES.Update
   db.rsACTUACIONES.Requery
End If

   CboCausante.Text = ""
   CboTipoRegistro.Text = ""
   TxtAnio.Text = ""
   TxtFechaIngreso.Text = ""
   TxtFechaHecho.Text = ""
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

Private Sub CmdPrimero_Click()
db.rsACTUACIONES.MoveFirst
CargarDatos
End Sub

Private Sub CmdAnterior_Click()
db.rsACTUACIONES.MovePrevious
If db.rsACTUACIONES.BOF Then
   db.rsACTUACIONES.MoveFirst
End If
CargarDatos
End Sub

Private Sub CmdSiguiente_Click()
db.rsACTUACIONES.MoveNext
If db.rsACTUACIONES.EOF Then
   db.rsACTUACIONES.MoveLast
End If
CargarDatos
End Sub

Private Sub CmdUltimo_Click()
db.rsACTUACIONES.MoveLast
CargarDatos

End Sub

Private Sub Form_Load()
On Error Resume Next

db.rsACTUACIONES.Open
CargarNro

db.rsESCALAFON.Open
db.rsPERSONAL.Open
db.rsTIPOREGISTRO.Open

Do While Not db.rsESCALAFON.EOF
   CboCausante.AddItem (db.rsESCALAFON.Fields!Jerarquia.Value & db.rsPERSONAL.Fields!apeNom.Value)
   CboNi.AddItem (db.rsPERSONAL.Fields!Ni.Value)

db.rsESCALAFON.MoveNext
db.rsPERSONAL.MoveNext
Loop

Do While Not db.rsPERSONAL.EOF
   CboNi.AddItem (db.rsPERSONAL.Fields!Ni.Value)
   db.rsPERSONAL.MoveNext
Loop

Do While Not db.rsTIPOREGISTRO.EOF
   CboTipoRegistro.AddItem (db.rsTIPOREGISTRO.Fields!descripcion.Value)
   db.rsTIPOREGISTRO.MoveNext
Loop

End Sub

Private Sub Salir_Click()
   If MsgBox("              ¿SALIR?", vbQuestion + vbYesNo, "SALIR") = vbYes Then
      Unload Me
   End If
End Sub
