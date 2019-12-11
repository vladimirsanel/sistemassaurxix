VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormListadoGral 
   BackColor       =   &H00FF8080&
   Caption         =   "LISTADO GENERAL DE ACTUACIONES"
   ClientHeight    =   8535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdActualizar 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&ACTUALIZAR"
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
      Left            =   17280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox TxtTipoRegistro 
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
      Height          =   375
      Left            =   9480
      TabIndex        =   8
      Top             =   1800
      Width           =   5415
   End
   Begin VB.TextBox TxtCausante 
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
      Height          =   375
      Left            =   9480
      TabIndex        =   7
      Top             =   1800
      Width           =   5415
   End
   Begin VB.TextBox TxtFechaHecho 
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
      Height          =   375
      Left            =   9480
      TabIndex        =   6
      Top             =   1800
      Width           =   5415
   End
   Begin VB.TextBox TxtFechaIngreso 
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
      Height          =   375
      Left            =   9480
      TabIndex        =   5
      Top             =   1800
      Width           =   5415
   End
   Begin VB.TextBox TxtNro 
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
      Height          =   375
      Left            =   9480
      TabIndex        =   3
      Top             =   1800
      Width           =   5415
   End
   Begin VB.ComboBox CboTipo 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Malgun Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1800
      Width           =   3975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FormListados.frx":0000
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   20295
      _ExtentX        =   35798
      _ExtentY        =   5953
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "ACTUACIONES"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "idActuaciones"
         Caption         =   "   N°"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "anio"
         Caption         =   "  AÑO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "fechaIngreso"
         Caption         =   "FECHA DE INGRESO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "d/MMM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "fechaHecho"
         Caption         =   " FECHA DEL HECHO"
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
         DataField       =   "causante"
         Caption         =   "                                                 CAUSANTE"
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
         DataField       =   "destino"
         Caption         =   "                      DESTINO"
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
      BeginProperty Column06 
         DataField       =   "carAdmin"
         Caption         =   "       CARÁTULA ADMINISTRATIVA"
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
      BeginProperty Column07 
         DataField       =   "caratula"
         Caption         =   "         CARÁTULA JUDICIAL"
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
      BeginProperty Column08 
         DataField       =   "idTipoRegistro"
         Caption         =   "   TIPO DE REGISTRO"
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
            DividerStyle    =   5
            ColumnWidth     =   599,811
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1769,953
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1769,953
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   5699,906
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   2910,047
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   3270,047
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   3674,835
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   2520
         EndProperty
      EndProperty
   End
   Begin VB.Label LblFiltrar 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "FILTRAR"
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
      Left            =   5760
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label LblListados 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "LISTADO GENERAL DE ACTUACIONES"
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
      Left            =   5760
      TabIndex        =   1
      Top             =   240
      Width           =   9735
   End
End
Attribute VB_Name = "FormListadoGral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CboTipo_Change()
If CboTipo.Text = "<SELECCIONE EL TIPO DE CONSULTA>" Then
   LblFiltrar.Visible = False
   TxtNro.Visible = False
   TxtFechaIngreso.Visible = False
   TxtFechaHecho.Visible = False
   TxtCausante.Visible = False
   TxtTipoRegistro.Visible = False
End If

If CboTipo.Text = "POR NÚMERO DE REGISTRO" Then
   LblFiltrar.Visible = True
   LblFiltrar = "NÚMERO DE REGISTRO:"
   TxtNro.Visible = True
   TxtFechaIngreso.Visible = False
   TxtFechaHecho.Visible = False
   TxtCausante.Visible = False
   TxtTipoRegistro.Visible = False
   TxtNro.SetFocus
End If

If CboTipo.Text = "POR FECHA DE INGRESO" Then
   LblFiltrar.Visible = True
   LblFiltrar = "FECHA DE INGRESO:"
   TxtNro.Visible = False
   TxtFechaIngreso.Visible = True
   TxtFechaHecho.Visible = False
   TxtCausante.Visible = False
   TxtTipoRegistro.Visible = False
   TxtFechaIngreso.SetFocus
End If

If CboTipo.Text = "POR FECHA DEL HECHO" Then
   LblFiltrar.Visible = True
   LblFiltrar = "FECHA DEL HECHO:"
   TxtNro.Visible = False
   TxtFechaIngreso.Visible = False
   TxtFechaHecho.Visible = True
   TxtCausante.Visible = False
   TxtTipoRegistro.Visible = False
   TxtFechaHecho.SetFocus
End If

If CboTipo.Text = "POR CAUSANTE" Then
   LblFiltrar.Visible = True
   LblFiltrar = "CAUSANTE:"
   TxtNro.Visible = False
   TxtFechaIngreso.Visible = False
   TxtFechaHecho.Visible = False
   TxtCausante.Visible = True
   TxtTipoRegistro.Visible = False
   TxtCausante.SetFocus
End If

If CboTipo.Text = "POR TIPO DE REGISTRO" Then
   LblFiltrar.Visible = True
   LblFiltrar = "TIPO DE REGISTRO:"
   TxtNro.Visible = False
   TxtFechaIngreso.Visible = False
   TxtFechaHecho.Visible = False
   TxtCausante.Visible = False
   TxtTipoRegistro.Visible = True
   TxtTipoRegistro.SetFocus
End If
End Sub

Private Sub CboTipo_Click()
If CboTipo.Text = "<SELECCIONE EL TIPO DE CONSULTA>" Then
   LblFiltrar.Visible = False
   TxtNro.Visible = False
   TxtFechaIngreso.Visible = False
   TxtFechaHecho.Visible = False
   TxtCausante.Visible = False
   TxtTipoRegistro.Visible = False
End If

If CboTipo.Text = "POR NÚMERO DE REGISTRO" Then
   LblFiltrar.Visible = True
   LblFiltrar = "NÚMERO DE REGISTRO:"
   TxtNro.Visible = True
   TxtFechaIngreso.Visible = False
   TxtFechaHecho.Visible = False
   TxtCausante.Visible = False
   TxtNro.SetFocus
End If

If CboTipo.Text = "POR FECHA DE INGRESO" Then
   LblFiltrar.Visible = True
   LblFiltrar = "FECHA DE INGRESO:"
   TxtNro.Visible = False
   TxtFechaIngreso.Visible = True
   TxtFechaHecho.Visible = False
   TxtCausante.Visible = False
   TxtTipoRegistro.Visible = False
   TxtFechaIngreso.SetFocus
End If

If CboTipo.Text = "POR FECHA DEL HECHO" Then
   LblFiltrar.Visible = True
   LblFiltrar = "FECHA DEL HECHO:"
   TxtNro.Visible = False
   TxtFechaIngreso.Visible = False
   TxtFechaHecho.Visible = True
   TxtCausante.Visible = False
   TxtTipoRegistro.Visible = False
   TxtFechaHecho.SetFocus
End If

If CboTipo.Text = "POR CAUSANTE" Then
   LblFiltrar.Visible = True
   LblFiltrar = "CAUSANTE:"
   TxtNro.Visible = False
   TxtFechaIngreso.Visible = False
   TxtFechaHecho.Visible = False
   TxtCausante.Visible = True
   TxtTipoRegistro.Visible = False
   TxtCausante.SetFocus
End If

If CboTipo.Text = "POR TIPO DE REGISTRO" Then
   LblFiltrar.Visible = True
   LblFiltrar = "TIPO DE REGISTRO:"
   TxtNro.Visible = False
   TxtFechaIngreso.Visible = False
   TxtFechaHecho.Visible = False
   TxtCausante.Visible = False
   TxtTipoRegistro.Visible = True
   TxtTipoRegistro.SetFocus
End If
End Sub

Private Sub TxtNro_Change()
If TxtNro.Text <> "" Then
   'db.rsACTUACIONES.Filter = "idActuaciones LIKE '*" & Trim(TxtNro.Text) & "*'"
    db.rsACTUACIONES.Filter = "idActuaciones = #" & Trim(TxtNro) & "#"
   Else
      Set DataGrid1.DataSource = db
          DataGrid1.Refresh
End If
End Sub

Private Sub Form_Load()
CboTipo.AddItem "<SELECCIONE EL TIPO DE CONSULTA>"
CboTipo.AddItem "POR NÚMERO DE REGISTRO"
CboTipo.AddItem "POR FECHA DE INGRESO"
CboTipo.AddItem "POR FECHA DEL HECHO"
CboTipo.AddItem "POR CAUSANTE"
CboTipo.AddItem "POR TIPO DE REGISTRO"
CboTipo = "<SELECCIONE EL TIPO DE CONSULTA>"
End Sub

