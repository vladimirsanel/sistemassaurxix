VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form FormPDF 
   BackColor       =   &H00FF8080&
   Caption         =   "REGLAMENTACIÓN"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
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
      Left            =   18840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Cmd10703 
      BackColor       =   &H000000FF&
      Caption         =   "Ley N° 10.703/15"
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
      Left            =   18960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Cmd1166 
      BackColor       =   &H000000FF&
      Caption         =   "Decreto N° 1166/18"
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
      Left            =   18960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton Cmd4174 
      BackColor       =   &H000000FF&
      Caption         =   "Decreto N° 4174/15"
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
      Left            =   18960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton Cmd0461 
      BackColor       =   &H000000FF&
      Caption         =   "Decreto N° 0461/15"
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
      Left            =   18960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton Cmd12521 
      BackColor       =   &H000000FF&
      Caption         =   "Ley N° 12.521/06"
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
      Left            =   18960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton CmdRSA 
      BackColor       =   &H000000FF&
      Caption         =   "R.S.A."
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
      Left            =   18960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   1335
   End
   Begin AcroPDFLibCtl.AcroPDF AcroPDF1 
      Height          =   10575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   18735
      _cx             =   5080
      _cy             =   5080
   End
   Begin VB.Image Image1 
      Height          =   5790
      Left            =   6480
      Picture         =   "FormPDF.frx":0000
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   6555
   End
End
Attribute VB_Name = "FormPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdRSA_Click()
AcroPDF1.LoadFile "C:\SistemaSSA\sistemassaurxix\LEYES Y DECRETOS\R.S.A.pdf"
AcroPDF1.Visible = True
End Sub

Private Sub Cmd12521_Click()
AcroPDF1.LoadFile "C:\SistemaSSA\sistemassaurxix\LEYES Y DECRETOS\Ley Nº 12.521.pdf"
AcroPDF1.Visible = True
End Sub

Private Sub Cmd0461_Click()
AcroPDF1.LoadFile "C:\SistemaSSA\sistemassaurxix\LEYES Y DECRETOS\DECRETO 0461.pdf"
AcroPDF1.Visible = True
End Sub

Private Sub Cmd4174_Click()
AcroPDF1.LoadFile "C:\SistemaSSA\sistemassaurxix\LEYES Y DECRETOS\DECRETO 4174-15.pdf"
AcroPDF1.Visible = True
End Sub

Private Sub Cmd1166_Click()
AcroPDF1.LoadFile "C:\SistemaSSA\sistemassaurxix\LEYES Y DECRETOS\DECRETO 1166-18.pdf"
AcroPDF1.Visible = True
End Sub

Private Sub Cmd10703_Click()
AcroPDF1.LoadFile "C:\SistemaSSA\sistemassaurxix\LEYES Y DECRETOS\Ley N° 10.703.pdf"
AcroPDF1.Visible = True
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub
