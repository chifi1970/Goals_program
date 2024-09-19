VERSION 5.00
Begin VB.Form forma_nota 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8640
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.lvButtons_H btnborrar 
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "&Clear"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btnok 
      Height          =   615
      Left            =   7560
      TabIndex        =   1
      Top             =   3600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1085
      Caption         =   "&OK"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin VB.TextBox txtnota 
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2775
      Left            =   240
      MaxLength       =   150
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   8175
   End
End
Attribute VB_Name = "forma_nota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnborrar_Click()
txtnota.Text = ""
End Sub

Private Sub btnOK_Click()
On Error Resume Next
form1.txtnotes.Text = txtnota.Text
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
txtnota.Text = form1.txtnotes.Text

Left = ((Screen.Width - Width) / 2) + 2500
Top = 500 '((Screen.Height - Height) / 2) - 2200


End Sub

Private Sub txtnota_Change()
On Error Resume Next
form1.txtnotes.Text = txtnota.Text

End Sub


