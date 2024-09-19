VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "mscal.ocx"
Begin VB.Form form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Weekly goals"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   20250
   Icon            =   "Forma_Main_goals.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Forma_Main_goals.frx":16B92
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ListBox List9 
      Height          =   1035
      Left            =   17760
      TabIndex        =   102
      Top             =   8520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   17400
      Sorted          =   -1  'True
      TabIndex        =   23
      Top             =   6720
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   19320
      Sorted          =   -1  'True
      TabIndex        =   24
      Top             =   5040
      Visible         =   0   'False
      Width           =   8415
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   80
      Left            =   11400
      ScaleHeight     =   45
      ScaleWidth      =   5985
      TabIndex        =   202
      Top             =   7560
      Width           =   6015
   End
   Begin VB.PictureBox msg 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   5640
      ScaleHeight     =   945
      ScaleWidth      =   9945
      TabIndex        =   6
      Top             =   4320
      Visible         =   0   'False
      Width           =   9975
      Begin VB.Label lblmsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Please, wait a moment... loading all data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Width           =   9855
      End
   End
   Begin VB.ListBox Lista_empleados 
      Height          =   255
      Left            =   16080
      Sorted          =   -1  'True
      TabIndex        =   200
      Top             =   9360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   4920
      ScaleHeight     =   3705
      ScaleWidth      =   5445
      TabIndex        =   197
      Top             =   4560
      Visible         =   0   'False
      Width           =   5480
      Begin RichTextLib.RichTextBox txtresultado 
         Height          =   2775
         Left            =   120
         TabIndex        =   199
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   4895
         _Version        =   393217
         BackColor       =   12640511
         ScrollBars      =   3
         TextRTF         =   $"Forma_Main_goals.frx":3BDF8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.lvButtons_H btnclose 
         Height          =   375
         Left            =   4440
         TabIndex        =   198
         Top             =   3120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Caption         =   "Close"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
   End
   Begin Project1.lvButtons_H btnreportetotal 
      Height          =   180
      Left            =   19440
      TabIndex        =   196
      Top             =   6000
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   318
      Caption         =   "Totals"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12640511
   End
   Begin MSFlexGridLib.MSFlexGrid grid8 
      Height          =   1215
      Left            =   18720
      TabIndex        =   193
      Top             =   8280
      Visible         =   0   'False
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   2143
      _Version        =   393216
      Rows            =   10
      Cols            =   22
      BackColor       =   -2147483636
      BackColorSel    =   16761024
      BackColorBkg    =   -2147483633
      GridColor       =   4210752
      GridColorFixed  =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   13440
      ScaleHeight     =   2535
      ScaleWidth      =   975
      TabIndex        =   162
      Top             =   600
      Width           =   975
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   163
         Top             =   0
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "B"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   164
         Top             =   0
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "C"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   165
         Top             =   240
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "D"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   4
         Left            =   300
         TabIndex        =   166
         Top             =   240
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "E"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   5
         Left            =   600
         TabIndex        =   167
         Top             =   240
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "F"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   168
         Top             =   480
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "G"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   7
         Left            =   300
         TabIndex        =   169
         Top             =   480
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "H"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   8
         Left            =   600
         TabIndex        =   170
         Top             =   480
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "I"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   171
         Top             =   720
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "J"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   10
         Left            =   300
         TabIndex        =   172
         Top             =   720
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "K"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   11
         Left            =   600
         TabIndex        =   173
         Top             =   720
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "L"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   12
         Left            =   0
         TabIndex        =   174
         Top             =   960
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "M"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   13
         Left            =   300
         TabIndex        =   175
         Top             =   960
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "N"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   14
         Left            =   600
         TabIndex        =   176
         Top             =   960
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "O"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   15
         Left            =   0
         TabIndex        =   177
         Top             =   1200
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "P"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   16
         Left            =   300
         TabIndex        =   178
         Top             =   1200
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "Q"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   17
         Left            =   600
         TabIndex        =   179
         Top             =   1200
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "R"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   18
         Left            =   0
         TabIndex        =   180
         Top             =   1440
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "S"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   19
         Left            =   300
         TabIndex        =   181
         Top             =   1440
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "T"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   20
         Left            =   600
         TabIndex        =   182
         Top             =   1440
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "U"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   21
         Left            =   0
         TabIndex        =   183
         Top             =   1680
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "V"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   22
         Left            =   300
         TabIndex        =   184
         Top             =   1680
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "W"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   23
         Left            =   600
         TabIndex        =   185
         Top             =   1680
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "X"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   24
         Left            =   0
         TabIndex        =   186
         Top             =   1920
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "Y"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   25
         Left            =   300
         TabIndex        =   187
         Top             =   1920
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "Z"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   26
         Left            =   0
         TabIndex        =   188
         Top             =   2160
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         Caption         =   "A-Z"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   -1  'True
         cBack           =   -2147483633
      End
      Begin Project1.lvButtons_H Btnletra 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   189
         Top             =   0
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   450
         Caption         =   "A"
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
         cBhover         =   8421504
         cGradient       =   8421504
         Gradient        =   3
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   1095
      Left            =   17040
      TabIndex        =   153
      Top             =   9720
      Visible         =   0   'False
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   1931
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.ListBox List7 
      Height          =   255
      Left            =   9000
      Sorted          =   -1  'True
      TabIndex        =   93
      Top             =   5040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin Project1.lvButtons_H btnDMV 
      Height          =   345
      Left            =   10800
      TabIndex        =   160
      Top             =   5715
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   609
      Caption         =   "DMV goals"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   4210752
      cGradient       =   4210752
      Gradient        =   1
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1440
      TabIndex        =   157
      Top             =   6560
      Width           =   2055
      Begin VB.OptionButton op_invoice 
         BackColor       =   &H00000000&
         Caption         =   "60 days"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   159
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton op_invoice 
         BackColor       =   &H00000000&
         Caption         =   "30 days"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   158
         Top             =   120
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   11760
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Project1.lvButtons_H btncargafile 
      Height          =   600
      Left            =   9000
      TabIndex        =   156
      Top             =   7560
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   1058
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "Forma_Main_goals.frx":3BE81
      ImgSize         =   48
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtfile 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5040
      TabIndex        =   155
      Top             =   7680
      Width           =   4095
   End
   Begin Project1.lvButtons_H btnupdate 
      Height          =   495
      Left            =   9000
      TabIndex        =   154
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "&Recalculate"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   32896
      cGradient       =   32896
      Gradient        =   1
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnend 
      Height          =   375
      Left            =   19200
      TabIndex        =   2
      Top             =   7560
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Caption         =   "En&d"
      CapAlign        =   2
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16777215
      cFHover         =   8421504
      cBhover         =   8421504
      LockHover       =   3
      cGradient       =   4210752
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      cBack           =   4210752
   End
   Begin Project1.lvButtons_H btn_attributos 
      Height          =   375
      Left            =   12120
      TabIndex        =   116
      Top             =   240
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Caption         =   "&Attributes"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   4210752
      cGradient       =   4210752
      Gradient        =   1
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   40
      cBack           =   12632256
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2295
      Left            =   3480
      TabIndex        =   134
      Top             =   960
      Visible         =   0   'False
      Width           =   3615
      _Version        =   524288
      _ExtentX        =   6376
      _ExtentY        =   4048
      _StockProps     =   1
      BackColor       =   14737632
      Year            =   2021
      Month           =   3
      Day             =   26
      DayLength       =   0
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   0   'False
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.lvButtons_H btn_NB 
      Height          =   255
      Left            =   10200
      TabIndex        =   1
      Top             =   840
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      Caption         =   "&NB"
      CapAlign        =   2
      BackStyle       =   2
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
      cBack           =   8421504
   End
   Begin Project1.lvButtons_H btn_GI 
      Height          =   255
      Left            =   10200
      TabIndex        =   0
      Top             =   2280
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      Caption         =   " &GI"
      CapAlign        =   2
      BackStyle       =   2
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
      cBack           =   8421504
   End
   Begin Project1.lvButtons_H btn_INV 
      Height          =   255
      Left            =   10200
      TabIndex        =   32
      Top             =   3480
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      Caption         =   "IN&V"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      cBack           =   8421504
   End
   Begin Project1.lvButtons_H btnorden 
      Height          =   375
      Left            =   10200
      TabIndex        =   151
      Top             =   2520
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      Image           =   "Forma_Main_goals.frx":3DDCC
      ImgSize         =   40
      cBack           =   8421504
   End
   Begin VB.ListBox List12 
      Height          =   255
      Left            =   17280
      TabIndex        =   147
      Top             =   8160
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.ListBox List4 
      Height          =   255
      Left            =   12120
      Sorted          =   -1  'True
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox List11 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   690
      Left            =   17520
      TabIndex        =   143
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10800
      Top             =   8160
   End
   Begin Project1.lvButtons_H btnacercade 
      Height          =   255
      Left            =   9960
      TabIndex        =   133
      Top             =   4560
      Visible         =   0   'False
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   450
      Caption         =   "?"
      CapAlign        =   2
      BackStyle       =   4
      Shape           =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   4210752
      cFHover         =   4210752
      cGradient       =   12632256
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin VB.ComboBox cboimpre 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   315
      Left            =   17160
      TabIndex        =   131
      Top             =   13300
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton btnsql 
      Caption         =   "SQL"
      Height          =   375
      Left            =   8760
      TabIndex        =   130
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtnotes 
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   15960
      MaxLength       =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   127
      Top             =   5760
      Width           =   2775
   End
   Begin VB.OptionButton op_week 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Week 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   126
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton op_week 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Week 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   125
      Top             =   480
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtdatepayday 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7200
      TabIndex        =   123
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtdateto 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5040
      TabIndex        =   121
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtdatefrom 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3480
      TabIndex        =   119
      Top             =   120
      Width           =   1095
   End
   Begin Project1.lvButtons_H btncargar_excel 
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Caption         =   "&Start"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   4210752
      cGradient       =   4210752
      Gradient        =   1
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   40
      cBack           =   16777215
   End
   Begin Project1.lvButtons_H btnexcel 
      Height          =   735
      Left            =   20640
      TabIndex        =   114
      Top             =   8520
      Visible         =   0   'False
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1296
      Caption         =   "&Export to Excel"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   4210752
      cGradient       =   4210752
      Gradient        =   1
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   40
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btnNB_deduction 
      Height          =   255
      Left            =   12960
      TabIndex        =   101
      Top             =   4200
      Visible         =   0   'False
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   450
      Caption         =   "Change"
      CapAlign        =   1
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   65535
   End
   Begin Project1.lvButtons_H btndeduction 
      Height          =   255
      Left            =   12960
      TabIndex        =   100
      Top             =   3960
      Visible         =   0   'False
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   450
      Caption         =   "Change"
      CapAlign        =   1
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   65535
   End
   Begin VB.CheckBox chklock1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Lock"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   12360
      TabIndex        =   113
      Top             =   3980
      Width           =   735
   End
   Begin VB.CheckBox chklock2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Lock"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   12360
      TabIndex        =   112
      Top             =   4220
      Width           =   735
   End
   Begin Project1.lvButtons_H btnporcentaje 
      Height          =   255
      Left            =   12960
      TabIndex        =   110
      Top             =   4920
      Visible         =   0   'False
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   450
      Caption         =   "Change"
      CapAlign        =   1
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   16744703
   End
   Begin VB.CheckBox chklock3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Lock"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   12360
      TabIndex        =   111
      Top             =   4940
      Width           =   735
   End
   Begin VB.ListBox List10 
      Height          =   840
      Left            =   18600
      TabIndex        =   105
      Top             =   6960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox lista_PhoneSales 
      Height          =   255
      Left            =   2280
      TabIndex        =   99
      Top             =   4560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ListBox List8 
      Height          =   255
      Left            =   9000
      TabIndex        =   96
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grid5 
      Height          =   960
      Left            =   5040
      TabIndex        =   94
      Top             =   5640
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1693
      _Version        =   393216
      Rows            =   3
      Cols            =   3
      BackColor       =   12632256
      BackColorSel    =   16761024
      GridLines       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox imagen_goals 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   11280
      ScaleHeight     =   1575
      ScaleWidth      =   8895
      TabIndex        =   41
      Top             =   6720
      Width           =   8895
      Begin VB.Label tablax_salary 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   5760
         TabIndex        =   68
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label tablax_salary 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   3
         Left            =   4320
         TabIndex        =   67
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label tablax_salary 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   66
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label tablax_salary 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   65
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label tablax_salary 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   64
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label tablax_comm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$500 Extra"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   5
         Left            =   7320
         TabIndex        =   63
         Top             =   480
         Width           =   1455
         WordWrap        =   -1  'True
      End
      Begin VB.Label tablax_comm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "25.00 %"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   5760
         TabIndex        =   62
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label tablax_comm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Index           =   3
         Left            =   4680
         TabIndex        =   61
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label tablax_comm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   60
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label tablax_comm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   59
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label tablax_comm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   58
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label tablax_NB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "11,500 BF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   7320
         TabIndex        =   57
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label tablax_NB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   5760
         TabIndex        =   56
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label tablax_NB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   4560
         TabIndex        =   55
         Top             =   315
         Width           =   1455
      End
      Begin VB.Label tablax_NB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   54
         Top             =   315
         Width           =   1455
      End
      Begin VB.Label tablax_NB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   53
         Top             =   315
         Width           =   1455
      End
      Begin VB.Label tablax_NB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   52
         Top             =   315
         Width           =   1335
      End
      Begin VB.Label Tablax_BF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monthly goal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   7320
         TabIndex        =   51
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Tablax_BF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "$2,999 - UP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   5760
         TabIndex        =   50
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
         WordWrap        =   -1  'True
      End
      Begin VB.Label Tablax_BF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   4560
         TabIndex        =   49
         Top             =   75
         Width           =   1455
         WordWrap        =   -1  'True
      End
      Begin VB.Label Tablax_BF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   48
         Top             =   75
         Width           =   1455
         WordWrap        =   -1  'True
      End
      Begin VB.Label Tablax_BF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   47
         Top             =   75
         Width           =   1455
         WordWrap        =   -1  'True
      End
      Begin VB.Label Tablax_BF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   42
         Top             =   75
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image2 
         Height          =   1215
         Index           =   0
         Left            =   120
         Picture         =   "Forma_Main_goals.frx":3E4EB
         Stretch         =   -1  'True
         Top             =   -45
         Width           =   1635
      End
      Begin VB.Image Image2 
         Height          =   1215
         Index           =   1
         Left            =   1560
         Picture         =   "Forma_Main_goals.frx":41BF3
         Stretch         =   -1  'True
         Top             =   -45
         Width           =   1635
      End
      Begin VB.Image Image2 
         Height          =   1215
         Index           =   2
         Left            =   3000
         Picture         =   "Forma_Main_goals.frx":45FD6
         Stretch         =   -1  'True
         Top             =   -45
         Width           =   1635
      End
      Begin VB.Image Image2 
         Height          =   1215
         Index           =   3
         Left            =   4440
         Picture         =   "Forma_Main_goals.frx":4A692
         Stretch         =   -1  'True
         Top             =   -45
         Width           =   1635
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   285
      Left            =   -300
      TabIndex        =   40
      Top             =   13560
      Width           =   150
   End
   Begin VB.ListBox List6 
      Height          =   255
      Left            =   15960
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox List5 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   690
      Left            =   17520
      TabIndex        =   37
      Top             =   1560
      Width           =   2655
   End
   Begin VB.ListBox lista_invoices30 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   360
      TabIndex        =   34
      Top             =   6960
      Width           =   4335
   End
   Begin VB.ComboBox cboexcepcion 
      Height          =   315
      Left            =   12840
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox lista_managers 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1500
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   29
      Top             =   4920
      Width           =   4575
   End
   Begin VB.ListBox List3 
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   6
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   18840
      Style           =   1  'Checkbox
      TabIndex        =   28
      Top             =   3240
      Width           =   1695
   End
   Begin VB.ListBox lista_users_shared 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   15960
      TabIndex        =   16
      Top             =   750
      Width           =   2655
   End
   Begin VB.ListBox lista_agentes 
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   10680
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   600
      Width           =   2775
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   1455
      Left            =   14880
      TabIndex        =   25
      Top             =   3240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   2566
      _Version        =   393217
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid grid4 
      Height          =   975
      Left            =   5040
      TabIndex        =   91
      Top             =   4560
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1720
      _Version        =   393216
      Rows            =   4
      Cols            =   3
      BackColor       =   12632256
      BackColorSel    =   16761024
      GridLines       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.lvButtons_H btn_errors 
      Height          =   375
      Left            =   9000
      TabIndex        =   92
      Top             =   4560
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Caption         =   "Load &MP Errors"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btn_UW 
      Height          =   375
      Left            =   9000
      TabIndex        =   95
      Top             =   5760
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Caption         =   "Load &UW ded."
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
      cBack           =   -2147483633
   End
   Begin Project1.lvButtons_H btnprint 
      Height          =   735
      Left            =   20520
      TabIndex        =   115
      Top             =   7920
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1296
      Caption         =   "&Print"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   4210752
      cGradient       =   4210752
      Gradient        =   1
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   40
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btnsave 
      Height          =   495
      Left            =   18960
      TabIndex        =   124
      Top             =   5440
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   873
      Caption         =   "&Save"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   4210752
      cGradient       =   4210752
      Gradient        =   1
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   32
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btnload 
      Height          =   495
      Left            =   18960
      TabIndex        =   144
      Top             =   4920
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   873
      Caption         =   "&Report"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   4210752
      cGradient       =   4210752
      Gradient        =   1
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   32
      cBack           =   12632256
   End
   Begin MSFlexGridLib.MSFlexGrid grid6 
      Height          =   840
      Left            =   5040
      TabIndex        =   145
      Top             =   6720
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1482
      _Version        =   393216
      Rows            =   3
      Cols            =   3
      BackColor       =   12648384
      BackColorSel    =   16761024
      GridLines       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.lvButtons_H btn_dmv 
      Height          =   375
      Left            =   9000
      TabIndex        =   146
      Top             =   6720
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Caption         =   "Load Wages"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid grid2 
      Height          =   1215
      Left            =   240
      TabIndex        =   190
      Top             =   2040
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   2143
      _Version        =   393216
      Rows            =   10
      Cols            =   22
      BackColor       =   -2147483636
      BackColorSel    =   16761024
      BackColorBkg    =   0
      GridColor       =   4210752
      GridColorFixed  =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   1215
      Left            =   240
      TabIndex        =   191
      Top             =   720
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   2143
      _Version        =   393216
      Rows            =   10
      Cols            =   22
      BackColor       =   -2147483636
      BackColorSel    =   16761024
      BackColorBkg    =   0
      GridColor       =   4210752
      GridColorFixed  =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid3 
      Height          =   1095
      Left            =   240
      TabIndex        =   192
      Top             =   3360
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   1931
      _Version        =   393216
      BackColor       =   -2147483636
      Rows            =   10
      Cols            =   22
      BackColorSel    =   16761024
      BackColorBkg    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   22
   End
   Begin VB.Label lbltotal_users_managers 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   19800
      TabIndex        =   201
      Top             =   1290
      Width           =   240
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000010&
      BorderWidth     =   3
      Height          =   315
      Index           =   2
      Left            =   19680
      Shape           =   4  'Rounded Rectangle
      Top             =   1230
      Width           =   495
   End
   Begin VB.Label lbltotal_unidades 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   165
      Left            =   19440
      TabIndex        =   195
      Top             =   4680
      Width           =   75
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   165
      Left            =   18960
      TabIndex        =   194
      Top             =   4680
      Width           =   360
   End
   Begin VB.Label lbltotal_final 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   435
      Left            =   16260
      TabIndex        =   161
      Top             =   4740
      Width           =   225
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Created by: Hector Navarro (2021-2024)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5040
      TabIndex        =   152
      Top             =   8040
      Width           =   2970
   End
   Begin VB.Label lbltotal_user 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   16680
      TabIndex        =   15
      Top             =   480
      Width           =   240
   End
   Begin VB.Label lbltotal_csr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   15360
      TabIndex        =   14
      Top             =   480
      Width           =   240
   End
   Begin VB.Label lblpenalty_dmv 
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   14160
      TabIndex        =   150
      Top             =   5760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Penalty:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Left            =   13320
      TabIndex        =   149
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lbldmv 
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12240
      TabIndex        =   148
      Top             =   5760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Programa en desarrollo"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   3720
      TabIndex        =   129
      Top             =   13320
      Width           =   3255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "(2021) Created by: Hector Navarro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   360
      TabIndex        =   39
      Top             =   13320
      Width           =   2775
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   -720
      Top             =   13080
      Width           =   22695
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Not applicable:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   17520
      TabIndex        =   142
      Top             =   2200
      Width           =   1065
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manager or Commercial:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   17520
      TabIndex        =   141
      Top             =   1365
      Width           =   1740
   End
   Begin VB.Image icon_mex 
      Height          =   600
      Left            =   17640
      Picture         =   "Forma_Main_goals.frx":4F078
      Stretch         =   -1  'True
      Top             =   320
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lbloficina 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "--------"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   225
      Left            =   17280
      TabIndex        =   140
      Top             =   3000
      Width           =   480
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Office:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   16560
      TabIndex        =   139
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblregistros 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000016&
      Height          =   255
      Left            =   11400
      TabIndex        =   138
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   10800
      TabIndex        =   137
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lbllae 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   19320
      TabIndex        =   136
      Top             =   540
      Width           =   735
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LAE:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   18840
      TabIndex        =   135
      Top             =   600
      Width           =   330
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Printer:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   16560
      TabIndex        =   132
      Top             =   13320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Shape Shape13 
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   13335
      Left            =   10240
      Top             =   -240
      Width           =   50
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   15240
      TabIndex        =   128
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Day:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6360
      TabIndex        =   122
      Top             =   120
      Width           =   660
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4680
      TabIndex        =   120
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2940
      TabIndex        =   118
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Final Total:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   15080
      TabIndex        =   26
      Top             =   4845
      Width           =   975
   End
   Begin VB.Label lbltotal_invoices 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   17000
      TabIndex        =   27
      Top             =   5280
      Width           =   150
   End
   Begin VB.Label lblfull_name 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "...."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   180
      Left            =   15765
      TabIndex        =   117
      Top             =   285
      Width           =   120
   End
   Begin VB.Shape Shape8 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   12240
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Shape Shape8 
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   12240
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Shape Shape8 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   12240
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   13725
      TabIndex        =   109
      Top             =   3720
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12920
      TabIndex        =   108
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblidentificacion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   19320
      TabIndex        =   107
      Top             =   285
      Width           =   225
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   18840
      TabIndex        =   106
      Top             =   360
      Width           =   225
   End
   Begin VB.Label lblinvoice2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   13080
      TabIndex        =   104
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblinvoice1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   12240
      TabIndex        =   103
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
      WordWrap        =   -1  'True
   End
   Begin VB.Image img_gerente 
      Height          =   495
      Left            =   14040
      Picture         =   "Forma_Main_goals.frx":4F4BA
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblinitials 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   19320
      TabIndex        =   98
      Top             =   45
      Width           =   225
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Initials:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   18720
      TabIndex        =   97
      Top             =   120
      Width           =   495
   End
   Begin VB.Image img_arrow2 
      Height          =   540
      Index           =   4
      Left            =   17640
      Picture         =   "Forma_Main_goals.frx":50232
      Stretch         =   -1  'True
      Top             =   6240
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image img_arrow2 
      Height          =   600
      Index           =   3
      Left            =   16320
      Picture         =   "Forma_Main_goals.frx":513C5
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Image img_arrow2 
      Height          =   600
      Index           =   2
      Left            =   14880
      Picture         =   "Forma_Main_goals.frx":549FD
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Image img_arrow2 
      Height          =   600
      Index           =   1
      Left            =   13440
      Picture         =   "Forma_Main_goals.frx":58326
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Image img_arrow2 
      Height          =   600
      Index           =   0
      Left            =   12000
      Picture         =   "Forma_Main_goals.frx":5BCDB
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblid 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   12240
      TabIndex        =   90
      Top             =   5400
      Width           =   2535
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   10680
      TabIndex        =   89
      Top             =   5400
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblcommission 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   12240
      TabIndex        =   88
      Top             =   5160
      Width           =   2535
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Commission"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   10680
      TabIndex        =   87
      Top             =   5160
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblporcentaje 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   13560
      TabIndex        =   86
      Top             =   4920
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Percentage"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   10680
      TabIndex        =   85
      Top             =   4920
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbltotal_bf 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   12240
      TabIndex        =   84
      Top             =   4680
      Width           =   2535
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total BF"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   10680
      TabIndex        =   83
      Top             =   4680
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbltotal_NB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   12240
      TabIndex        =   82
      Top             =   4440
      Width           =   2535
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total NB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   10680
      TabIndex        =   81
      Top             =   4440
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblnb_deduc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   13560
      TabIndex        =   80
      Top             =   4200
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NB Deduction"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   10680
      TabIndex        =   79
      Top             =   4200
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbldeduction 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   13560
      TabIndex        =   78
      Top             =   3960
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DEDUCTION"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   10680
      TabIndex        =   77
      Top             =   3960
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblinvoice 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   12240
      TabIndex        =   76
      Top             =   3720
      Width           =   2520
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Invoice"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   10680
      TabIndex        =   75
      Top             =   3720
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblbf 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   12240
      TabIndex        =   74
      Top             =   3480
      Width           =   2535
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BF"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   10680
      TabIndex        =   73
      Top             =   3480
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblnb 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   12240
      TabIndex        =   72
      Top             =   3240
      Width           =   2535
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   10680
      TabIndex        =   71
      Top             =   3240
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoices"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   70
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Image img_arrow 
      Height          =   600
      Index           =   0
      Left            =   11520
      Picture         =   "Forma_Main_goals.frx":5F1A1
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deduction DMV"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10800
      TabIndex        =   69
      Top             =   5760
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Image img_arrow 
      Height          =   360
      Index           =   5
      Left            =   18840
      Picture         =   "Forma_Main_goals.frx":621E4
      Stretch         =   -1  'True
      Top             =   6360
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Image img_arrow 
      Height          =   600
      Index           =   4
      Left            =   11520
      Picture         =   "Forma_Main_goals.frx":62B8D
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Image img_arrow 
      Height          =   600
      Index           =   3
      Left            =   15840
      Picture         =   "Forma_Main_goals.frx":65C3B
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Image img_arrow 
      Height          =   600
      Index           =   2
      Left            =   14400
      Picture         =   "Forma_Main_goals.frx":68E18
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Image img_arrow 
      Height          =   600
      Index           =   1
      Left            =   12960
      Picture         =   "Forma_Main_goals.frx":6C174
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Salary"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10440
      TabIndex        =   46
      Top             =   7680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Commission"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   165
      Left            =   10440
      TabIndex        =   45
      Top             =   7350
      Width           =   750
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NB Goal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   10440
      TabIndex        =   44
      Top             =   7060
      Width           =   555
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BF Goal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   10440
      TabIndex        =   43
      Top             =   6780
      Width           =   540
   End
   Begin VB.Image Image4 
      Height          =   660
      Left            =   3480
      Picture         =   "Forma_Main_goals.frx":6F54F
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label lblgrandtotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   17400
      TabIndex        =   36
      Top             =   7920
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total before ded.:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   15120
      TabIndex        =   35
      Top             =   5325
      Width           =   1560
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Exceptions:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   30
      Top             =   4540
      Width           =   2415
   End
   Begin VB.Label lbl_total_CSR_user 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   345
      Left            =   16290
      TabIndex        =   22
      Top             =   2060
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   15720
      TabIndex        =   21
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   2
      Left            =   16080
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label lblsuma_total_user 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   16365
      TabIndex        =   20
      Top             =   1680
      Width           =   285
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   15800
      TabIndex        =   19
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblsuma_total_CSR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   15165
      TabIndex        =   18
      Top             =   1680
      Width           =   285
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000C0C0&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   16080
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   15000
      TabIndex        =   17
      Top             =   2115
      Width           =   510
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000C0C0&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   14880
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   16080
      TabIndex        =   13
      Top             =   480
      Width           =   465
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000010&
      BorderWidth     =   3
      Height          =   315
      Index           =   1
      Left            =   15960
      Shape           =   4  'Rounded Rectangle
      Top             =   435
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000010&
      BorderWidth     =   3
      Height          =   320
      Index           =   0
      Left            =   14760
      Shape           =   4  'Rounded Rectangle
      Top             =   440
      Width           =   975
   End
   Begin VB.Label lblagent 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "...."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   15720
      TabIndex        =   12
      Top             =   0
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   15000
      TabIndex        =   11
      Top             =   80
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CSR:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   14840
      TabIndex        =   10
      Top             =   480
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "A g e n t s :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10800
      TabIndex        =   9
      Top             =   320
      Width           =   1335
   End
   Begin VB.Label lbl_count 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   9720
      TabIndex        =   5
      Top             =   3960
      Width           =   240
   End
   Begin VB.Label lbl_count 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   9720
      TabIndex        =   4
      Top             =   1560
      Width           =   240
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   12720
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H8000000A&
      BorderStyle     =   3  'Dot
      FillColor       =   &H80000011&
      FillStyle       =   0  'Solid
      Height          =   3015
      Left            =   10560
      Top             =   3120
      Width           =   4335
   End
   Begin VB.Shape Shape12 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   3
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   555
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   -105
      Width           =   5775
   End
   Begin VB.Shape Shape5 
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   -120
      Top             =   13200
      Width           =   22335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      FillStyle       =   0  'Solid
      Height          =   13335
      Left            =   -120
      Top             =   -240
      Width           =   10455
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   8
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   14920
      Shape           =   4  'Rounded Rectangle
      Top             =   4720
      Width           =   3855
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   15000
      Shape           =   4  'Rounded Rectangle
      Top             =   5160
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   14880
      Picture         =   "Forma_Main_goals.frx":70BF7
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Image Image8 
      Height          =   1240
      Left            =   14880
      Picture         =   "Forma_Main_goals.frx":732A1
      Stretch         =   -1  'True
      Top             =   675
      Width           =   675
   End
   Begin VB.Image Image6 
      Height          =   1125
      Left            =   16320
      Picture         =   "Forma_Main_goals.frx":73F64
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image7 
      Height          =   495
      Left            =   16800
      Picture         =   "Forma_Main_goals.frx":74AD6
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1335
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Dim columna As Integer, UltimaFila As Integer, Ultimacolumna As Integer, lineas_BF As Integer, lineas_GI  As Integer, lineas_NB  As Integer, lineas_INV  As Integer
Dim tabla_NB(2000, 6), tabla_GI(2000, 6), tabla_INV(2000, 6), letra$, agente$, tabla(20, 2), grandtotal As Single, total_facturas_propias As Single, total_facturas_ajenas As Single
Dim marca_porcentaje As Integer, notes(50) As String * 300, num_fila_agente As Integer, porcentaje(50) As String * 10
Dim cargado As Integer, seg As Integer, calen As Integer, ubicacion(200, 3), expiracion_invoices As Integer, marca As Integer, contador_usuarios As Integer
Dim matrix_NB(2000, 12)


Dim user_original$, csr_original$, concepto$, ID_Cliente$
         

Dim NB_Goal1, NB_Goal2, NB_Goal3, NB_Goal4, NB_Goal5
    Dim commission1, commission2, commission3, commission4, commission5
    Dim rango1a, rango1b, rango2a, rango2b, rango3a, rango3b, rango4a, rango4b, rango5a, rango5b
    Dim categoria_tier As Integer
    

 Dim DesignX As Integer
      Dim DesignY As Integer
Dim primeravez As Integer


Public Sub carga_impresoras()
On Error Resume Next

Dim cImprGen As String
    cImprGen = cboimpre.Text
    
cboimpre.Clear
ruta$ = "c:\goals\"
    
If Dir$(ruta$ + "printer") <> "" Then
 nf = FreeFile
 Open ruta$ + "printer" For Input Shared As #nf
 Lock #nf
 Line Input #nf, P1$
 Line Input #nf, P2$
 Unlock #nf
 Close #nf
 
 cImprGen = P1$
 cboimpre.Text = P1$

End If
    
    
    
    
For Each xprint In Printers
           If xprint.DeviceName = cImprGen Then
              ' La define como predeterminada del sistema.
              Set Printer = xprint
              DoEvents
              Exit For
           End If
Next
        
        
        
For Each xprint In Printers
        cboimpre.AddItem xprint.DeviceName
Next
        
        
nf = FreeFile
 Open ruta$ + "printer" For Output Shared As #nf
 Lock #nf
 Print #nf, Printer.DeviceName
 Print #nf, Printer.Port
 Unlock #nf
 Close #nf
 
 
 For t = 0 To cboimpre.ListCount - 1
   If cboimpre.List(t) = Printer.DeviceName Then
       cboimpre.ListIndex = t
       Exit For
   End If
 Next t
        
        
        
        
End Sub



Public Function Redondear(dNumero As Double, iDecimales As Integer) As Double
    Dim lMultiplicador As Long
    Dim dRetorno As Double
    
    If iDecimales > 9 Then iDecimales = 9
    lMultiplicador = 10 ^ iDecimales
    dRetorno = CDbl(CLng(dNumero * lMultiplicador)) / lMultiplicador
    
    Redondear = dRetorno
End Function
Public Sub Checa_status()


End Sub
' Load a TreeView control from a file that uses tabs
' to show indentation.
Private Sub LoadTreeViewFromFile(ByVal file_name As String, ByVal trv As TreeView)
Dim fnum As Integer
Dim text_line As String
Dim level As Integer
Dim tree_nodes() As Node
Dim num_nodes As Integer

    fnum = FreeFile
    Open file_name For Input As fnum

    TreeView1.Nodes.Clear
    Do While Not EOF(fnum)
        ' Get a line.
        Line Input #fnum, text_line

        ' Find the level of indentation.
        level = 1
        Do While Left$(text_line, 1) = vbTab
            level = level + 1
            text_line = Mid$(text_line, 2)
        Loop

        ' Make room for the new node.
        If level > num_nodes Then
            num_nodes = level
            ReDim Preserve tree_nodes(1 To num_nodes)
        End If

        ' Add the new node.
        If level = 1 Then
            Set tree_nodes(level) = TreeView1.Nodes.Add(, , , text_line)
        Else
            Set tree_nodes(level) = TreeView1.Nodes.Add(tree_nodes(level - 1), tvwChild, , text_line)
            tree_nodes(level).EnsureVisible
        End If
    Loop

    Close fnum
    If trv.Nodes.Count > 0 Then trv.Nodes(1).EnsureVisible
End Sub


Private Sub btn_attributos_Click()
On Error Resume Next
If lista_agentes.ListCount = 0 Then
  Exit Sub
End If

lblmsg.Caption = "Please, wait a moment... loading all data"
msg.Visible = True
msg.Refresh

Btnletra_Click (26)
Btnletra(26).Value = True
Load forma_atributos

msg.Visible = False
msg.Refresh

forma_atributos.Show 1




carga_ubicaciones
carga_attributos


End Sub

Private Sub btn_dmv_Click()
On Error Resume Next
Dim xlApp2 As Excel.Application
Dim xlLibro2 As Excel.Workbook
Dim xlHoja2 As Excel.Worksheet
Dim varMatriz As Variant
Dim lngUltimaFila As Long
If btn_dmv.Visible = False Then Exit Sub
lblmsg.Caption = "Please, wait a moment... loading hourly wages"
msg.Visible = True
msg.Refresh

btn_dmv.Visible = False
contador = 0

inicio:


If txtfile.Text <> "" Then
  n$ = txtfile.Text
Else
  n$ = "c:\goals\wages.xlsx"
End If



grid6.Clear
List12.Clear

If Dir$(n$) = "" Then
   'MsgBox "The file " + n$ + " has not been found", 64, "Attention"
   GoTo final
End If

'reset variables de los objetos
Set xlHoja2 = Nothing
Set xlLibro2 = Nothing
Set xlApp2 = Nothing

'abrir programa Excel
Set xlApp2 = New Excel.Application
'xl.Visible = True

'abrir el archivo Excel
'(archivo en la misma carpeta)
Set xlLibro2 = xlApp2.Workbooks.Open(FileName:=n$, ReadOnly:=True)

' Get the first worksheet.
 Set xlHoja2 = xlApp2.Worksheets(1)
' Set xlHoja = xlApp.Worksheets("bf")

'1. Si conoces el rango a leer
'varMatriz = xlHoja.Range(«A1:C10»).Value

'2. Si no conoces el rango
' lngUltimaFila = Columns("A:A").range("A65536").End(xlUp).row
lngUltimaFila = 2000

    ActiveCell.SpecialCells(xlLastCell).Select
    
    
    ultimafilax = ActiveCell.row
    ultimacolumnax = ActiveCell.Column

' sino carga el archivo entonces abrelo
If lngUltimaFila = 0 Then
  
  lngUltimaFila = lineas_uw
  contador = contador + 1
  
  
  If contador >= 2 Then
     Exit Sub
  End If
  
Else

  lineas_uw = lngUltimaFila

  
End If

continua:

 varMatriz = xlHoja2.range(xlHoja2.Cells(1, 1), xlHoja2.Cells(lngUltimaFila, 22))   ' cambie 10 por 22

'grid3.Clear
'utilizamos los datos…
'txtLlamadas.Text = varMatriz(10, 3)
grid6.Rows = lngUltimaFila + 2
grid6.cols = 28


cont = 0
For t = 1 To grid6.Rows - 1
  grid6.row = t - 1
  If t > 1 And t <= lngUltimaFila And varMatriz(t, 1) <> "" Then
     cont = cont + 1
     grid6.col = 0
     grid6.Text = cont
     
  End If
  
  For Y = 1 To 22
   grid6.col = Y
   grid6.Text = varMatriz(t, Y)
  Next Y
  
Next t
   
'grid3.Rows = cont + 2
ultima_linea = cont + 2
'lbl_count(1).Caption = grid3.Rows - 2
guarda_filas = cont + 2
grid6.Rows = guarda_filas

'cerramos el archivo Excel
xlLibro2.Close SaveChanges:=False
xlApp2.Quit

'reset variables de los objetos
Set xlHoja2 = Nothing
Set xlLibro2 = Nothing
Set xlApp2 = Nothing


grid6.ColWidth(0) = 600
grid6.ColWidth(1) = 1200  '1
grid6.ColWidth(2) = 1100  '2
grid6.ColWidth(3) = 1200  '3
grid6.ColWidth(4) = 1200  '4
grid6.ColWidth(5) = 1200  '5
grid6.ColWidth(6) = 1200  '6
grid6.ColWidth(7) = 1200  '7
grid6.ColWidth(8) = 1500  '8
grid6.ColWidth(9) = 1200  '9
grid6.ColWidth(10) = 1200 '10
grid6.ColWidth(11) = 1500 '11
grid6.ColWidth(12) = 1200 '12
grid6.ColWidth(13) = 1400 '13
grid6.ColWidth(14) = 1800 '14
grid6.ColWidth(15) = 1600 '15
grid6.ColWidth(16) = 1600 '16
grid6.ColWidth(17) = 1300 '17
grid6.ColWidth(18) = 1300 '18
grid6.ColWidth(19) = 1300 '19
grid6.ColWidth(20) = 1300 '20
grid6.ColWidth(21) = 1300 '21
grid6.ColWidth(22) = 1300 '22



grid6.Rows = grid6.Rows - 2
lin = 0


GoTo final

For t = 1 To grid6.Rows - 1
   grid6.row = t
   grid6.col = 1
   numrow$ = grid6.Text
   
   If Val(numrow$) <= 0 Or numrow$ = "" Then
      For z = 0 To grid6.cols
          grid6.col = z
          grid6.Text = ""
      Next z
      lin = lin + 1
   End If
   
   
Next t

grid6.Rows = grid6.Rows - lin + 1




GoTo final




'Exit Sub


List12.Clear


For t = 1 To grid6.Rows
  grid6.row = t
  
   grid6.col = 2
   n$ = RTrim(LTrim(UCase(grid6.Text)))
   
   
   grid6.col = 3
   cant1 = Val(Format(grid6.Text, "00000.00"))
  
   existe = 0
   c = 0
   If n$ <> "" Then
      If List12.ListCount = 0 Then existe = 2
      For Y = 0 To List12.ListCount - 1
         n2$ = UCase(Left(List12.List(Y), 3))
         c = Val(Mid(List12.List(Y), 5, 3))
         cant = Val(Right(List12.List(Y), 7))
         
         If n2$ = n$ Then
            existe = 1
            cantf = cant + cant1
            c = c + 1
            
            List12.RemoveItem Y
            Exit For
         End If
      
         
      Next Y
      
   End If
   
  If existe = 2 And List12.ListCount = 0 Then c = 1
  If existe = 2 And List12.ListCount = 1 Then c = 1
   
   
  If existe = 1 And c > 0 Then
    cantf = cant + cant1
    List12.AddItem Format(n$, "@@@") + Space(1) + Format(c, "000") + " " + Format(cantf, "0000.00")
  ElseIf existe = 2 And c > 0 Then
    cantf = cant + cant1
    List12.AddItem Format(n$, "@@@") + Space(1) + Format(c, "000") + " " + Format(cantf, "0000.00")
  ElseIf existe = 0 And c > 0 Then
    cantf = cant1
    c = 1
    List12.AddItem Format(n$, "@@@") + Space(1) + Format(c, "000") + " " + Format(cantf, "0000.00")
  End If
  
  cantf = 0
  cant = 0
  
Next t





final:


' separa por fechas. pone solo las fechas correctas

grid.Clear
grid.Rows = grid6.Rows

grid6.cols = grid6.cols + 2
grid.cols = grid6.cols




linea = 0
For t = 0 To grid6.Rows
   grid6.row = t
   grid6.col = 4
   fecha_reporte$ = Format(grid6.Text, "mm/dd/yyyy")
   
   If Format(txtdatefrom.Text, "mm/dd/yyyy") = fecha_reporte$ Or grid6.row = 0 Then
      grid.row = linea
      grid6.row = linea
      For z = 1 To grid6.cols
           grid6.col = z
           grid.col = z
           grid.Text = grid6.Text
      Next z
      linea = linea + 1
   
   End If
   
Next t




grid6.Rows = linea

'grid6.Clear
'For t = 0 To grid.Rows + 1
'   grid6.row = t + 1
'   grid6.col = 0
'   grid6.Text = Str(t + 1)
   
'   grid.row = t
'   grid6.row = t
   
  
'   For z = 1 To grid.cols
'            grid.col = z
'      grid6.col = z
'       grid6.Text = grid.Text
'   Next z
'Next t



grid6.col = 2
grid6.Sort = flexSortStringAscending
For t = 1 To grid6.Rows - 1
  grid6.row = t
  grid6.col = 0
  grid6.Text = t
Next t




msg.Visible = False
msg.Refresh


' separa_campos
btn_dmv.Visible = True


End Sub

Private Sub btn_errors_Click()
On Error Resume Next

GoTo start

grid.Clear
grid4.Clear

Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    Set Rs = New ADODB.Recordset
    'Checa_status
   
    user1$ = ""
    
    
    comando1$ = "DECLARE @p0 DateTime = '" + Format(txtdatefrom.Text, "yyyy-mm-dd") + "' " + Chr$(13) & _
    "DECLARE @p1 DateTime = '" + Format(txtdateto.Text, "yyyy-mm-dd") + "' " + Chr$(13) & _
    "DECLARE @p2 Int = 5 " + Chr$(13) & _
    "DECLARE @p3 Int = 7 " + Chr$(13) & _
    "DECLARE @p4 DateTime = '" + Format(txtdatefrom.Text, "yyyy-mm-dd") + "' " + Chr$(13) & _
    "DECLARE @p5 DateTime = '" + Format(txtdateto.Text, "yyyy-mm-dd") + "' " + Chr$(13) & _
    "DECLARE @p6 Int = 3 " + Chr$(13) & _
    "DECLARE @p7 DateTime = '" + Format(txtdatefrom.Text, "yyyy-mm-dd") + "' " + Chr$(13) & _
    "DECLARE @p8 DateTime = '" + Format(txtdateto.Text, "yyyy-mm-dd") + "' " + Chr$(13) & _
    "DECLARE @p9 Decimal(1,0) = 0 " + Chr$(13) & _
    "SELECT [t7].[IdMPPF] AS [IdMppf], [t0].[IDReceiptHDR] AS [IdreceiptHdr], [t7].[IdEmployee], [t11].[Username] AS [Username], [t17].[IdPoliciesHDR] AS [IdPoliciesHdr], " & _
    "[t17].[PolicyNumber], [t0].[Date], [t0].[IdOffice], [t19].[Office], [t15].[IDEmployee] AS [USRIdEmployee], [t15].[Username] AS [USRUserName], [t14].[IdCustomer], " & _
    "(CASE WHEN [t6].[test] IS NULL THEN CONVERT(Decimal(33,4),@p9) ELSE CONVERT(Decimal(33,4),[t6].[Amount]) END) AS [FeeAmount], [t7].[Active] " & _
    "FROM [ReceiptsHDR] AS [t0] " & _
    "LEFT OUTER JOIN (SELECT 1 AS [test], [t1].[IDReceiptHDR], [t2].[IdInvoiceItem], [t1].[Date], [t1].[Active], [t1].[Void], " & _
    "[t2].[Active] AS [Active2] FROM [ReceiptsHDR] AS [t1] " & _
    "INNER JOIN [ReceiptsDTL] AS [t2] ON [t1].[IDReceiptHDR] = [t2].[IdReceiptHDR]) AS [t3] ON ([t0].[IDReceiptHDR] = [t3].[IDReceiptHDR]) AND (CONVERT(DATE, [t3].[Date]) >= @p0) " & _
    "AND (CONVERT(DATE, [t3].[Date]) <= @p1) AND ([t3].[IdInvoiceItem] IN (@p2, @p3)) AND ([t3].[Active] = 1) AND (NOT ([t3].[Void] = 1)) AND ([t3].[Active2] = 1) " & _
    "LEFT OUTER JOIN (    SELECT 1 AS [test], [t4].[IDReceiptHDR], [t5].[IdInvoiceItem], [t5].[Amount], [t4].[Date], [t4].[Active], [t4].[Void], [t5].[Active] AS [Active2] " & _
    "FROM [ReceiptsHDR] AS [t4]"



    
    '"SELECT [t7].[IdMPPF] AS [IdMppf], [t0].[IDReceiptHDR] AS [IdreceiptHdr], [t7].[IdEmployee], [t11].[Username] AS [Username], [t17].[IdPoliciesHDR] AS [IdPoliciesHdr], " + Chr$(13) & _
    '"[t17].[PolicyNumber], [t0].[Date], [t0].[IdOffice], [t19].[Office], [t15].[IDEmployee] AS [USRIdEmployee], [t15].[Username] AS [USRUserName], [t14].[IdCustomer], " + Chr$(13) & _
    '"(CASE WHEN [t6].[test] IS NULL THEN CONVERT(Decimal(33,4),@p9) ELSE CONVERT(Decimal(33,4),[t6].[Amount]) END) AS [FeeAmount], [t7].[Active] " + Chr$(13) & _
    '"FROM [ReceiptsHDR] AS [t0] " + Chr$(13) & _
    '"LEFT OUTER JOIN ( " + Chr$(13) & _
    '"SELECT 1 AS [test], [t1].[IDReceiptHDR], [t2].[IdInvoiceItem], [t1].[Date], [t1].[Active], [t1].[Void], [t2].[Active] AS [Active2] " + Chr$(13) & _
    '"FROM [ReceiptsHDR] AS [t1] " + Chr$(13) & _
    '"INNER JOIN [ReceiptsDTL] AS [t2] ON [t1].[IDReceiptHDR] = [t2].[IdReceiptHDR] " + Chr$(13) & _
    '") AS [t3] ON ([t0].[IDReceiptHDR] = [t3].[IDReceiptHDR]) AND (CONVERT(DATE, [t3].[Date]) >= @p0) AND (CONVERT(DATE, [t3].[Date]) <= @p1) AND ([t3].[IdInvoiceItem] IN (@p2, @p3)) AND ([t3].[Active] = 1) AND (NOT ([t3].[Void] = 1)) AND ([t3].[Active2] = 1) " + Chr$(13) & _
    '"LEFT OUTER JOIN ( " + Chr$(13) & _
    '"SELECT 1 AS [test], [t4].[IDReceiptHDR], [t5].[IdInvoiceItem], [t5].[Amount], [t4].[Date], [t4].[Active], [t4].[Void], [t5].[Active] AS [Active2] " + Chr$(13) & _
    '"FROM [ReceiptsHDR] AS [t4] " + Chr$(13)

    
'    comando2$ = "INNER JOIN [ReceiptsDTL] AS [t5] ON [t4].[IDReceiptHDR] = [t5].[IdReceiptHDR] " + Chr$(13) & _
'    ") AS [t6] ON ([t0].[IDReceiptHDR] = [t6].[IDReceiptHDR]) AND (CONVERT(DATE, [t6].[Date]) >= @p4) AND (CONVERT(DATE, [t6].[Date]) <= @p5) AND ([t6].[IdInvoiceItem] IN (@p6)) AND ([t6].[Active] = 1) AND (NOT ([t6].[Void] = 1)) AND ([t6].[Active2] = 1) " + Chr$(13) & _
'    "INNER JOIN [MPPFCalc] AS [t7] ON [t0].[IDReceiptHDR] = [t7].[IdReceiptsHDR] " + Chr$(13) & _
'    "LEFT OUTER JOIN [AGIConfig] AS [t8] ON [t7].[IdStatus] = ([t8].[IdAGIConfig]) " + Chr$(13) & _
'    "LEFT OUTER JOIN [AGIConfig] AS [t9] ON [t7].[IdDeduction] = ([t9].[IdAGIConfig]) " + Chr$(13) & _
'    "LEFT OUTER JOIN [ErrorTagTypeCatalog] AS [t10] ON [t7].[IdTypeErrorTag] = ([t10].[IdTypeErrorTag]) " + Chr$(13) & _
'    "LEFT OUTER JOIN [EmployeeInfo] AS [t11] ON [t7].[IdEmployee] = [t11].[IDEmployee] " + Chr$(13) & _
'    "LEFT OUTER JOIN [InvoiceItemCatalog] AS [t12] ON [t3].[IdInvoiceItem] = [t12].[IdInvoiceItem] " + Chr$(13) & _
'    "LEFT OUTER JOIN [InvoiceItemCatalog] AS [t13] ON [t6].[IdInvoiceItem] = [t13].[IdInvoiceItem] " + Chr$(13) & _
'    "INNER JOIN [Customers] AS [t14] ON [t0].[IdCustomer] = [t14].[IdCustomer] " + Chr$(13) & _
'    "INNER JOIN [EmployeeInfo] AS [t15] ON [t0].[IdEmployeeUSR] = [t15].[IDEmployee] " + Chr$(13) & _
'    "INNER JOIN [EmployeeInfo] AS [t16] ON [t0].[IdEmployeeCSR1] = [t16].[IDEmployee] " + Chr$(13) & _
'    "INNER JOIN [PoliciesHDR] AS [t17] ON [t0].[IdPoliciesHDR] = [t17].[IdPoliciesHDR] " + Chr$(13) & _
'    "INNER JOIN [InsuranceCatalog] AS [t18] ON [t17].[IdCompany] = [t18].[IdCompany] " + Chr$(13) & _
'    "INNER JOIN [OfficesCatalog] AS [t19] ON [t0].[IdOffice] = [t19].[IdOffice] " + Chr$(13) & _
'    "WHERE (CONVERT(DATE, [t0].[Date]) >= @p7) AND (CONVERT(DATE, [t0].[Date]) <= @p8) AND ([t7].[Active] = 1) AND (([t3].[test] IS NOT NULL) OR ([t6].[test] IS NOT NULL)) " + Chr$(13) & _
'    "and t7.IdDeduction=16 ORDER BY [t0].[IDReceiptHDR]"

    
     comando2$ = "INNER JOIN [ReceiptsDTL] AS [t5] ON [t4].[IDReceiptHDR] = [t5].[IdReceiptHDR]) AS [t6] ON ([t0].[IDReceiptHDR] = [t6].[IDReceiptHDR]) AND (CONVERT(DATE, [t6].[Date]) >= @p4) " & _
     "AND (CONVERT(DATE, [t6].[Date]) <= @p5) AND ([t6].[IdInvoiceItem] IN (@p6)) AND ([t6].[Active] = 1) AND (NOT ([t6].[Void] = 1)) AND ([t6].[Active2] = 1) " & _
     "INNER JOIN [MPPFCalc] AS [t7] ON [t0].[IDReceiptHDR] = [t7].[IdReceiptsHDR] " & _
     "LEFT OUTER JOIN [AGIConfig] AS [t8] ON [t7].[IdStatus] = ([t8].[IdAGIConfig]) " & _
     "LEFT OUTER JOIN [AGIConfig] AS [t9] ON [t7].[IdDeduction] = ([t9].[IdAGIConfig]) " & _
     "LEFT OUTER JOIN [ErrorTagTypeCatalog] AS [t10] ON [t7].[IdTypeErrorTag] = ([t10].[IdTypeErrorTag]) " & _
     "LEFT OUTER JOIN [EmployeeInfo] AS [t11] ON [t7].[IdEmployee] = [t11].[IDEmployee] " & _
     "LEFT OUTER JOIN [InvoiceItemCatalog] AS [t12] ON [t3].[IdInvoiceItem] = [t12].[IdInvoiceItem] " & _
     "LEFT OUTER JOIN [InvoiceItemCatalog] AS [t13] ON [t6].[IdInvoiceItem] = [t13].[IdInvoiceItem] " & _
     "INNER JOIN [Customers] AS [t14] ON [t0].[IdCustomer] = [t14].[IdCustomer] " & _
     "INNER JOIN [EmployeeInfo] AS [t15] ON [t0].[IdEmployeeUSR] = [t15].[IDEmployee] " & _
     "INNER JOIN [EmployeeInfo] AS [t16] ON [t0].[IdEmployeeCSR1] = [t16].[IDEmployee] " & _
     "INNER JOIN [PoliciesHDR] AS [t17] ON [t0].[IdPoliciesHDR] = [t17].[IdPoliciesHDR] " & _
     "INNER JOIN [InsuranceCatalog] AS [t18] ON [t17].[IdCompany] = [t18].[IdCompany] " & _
     "INNER JOIN [OfficesCatalog] AS [t19] ON [t0].[IdOffice] = [t19].[IdOffice] " & _
     "WHERE (CONVERT(DATE, [t0].[Date]) >= @p7) AND (CONVERT(DATE, [t0].[Date]) <= @p8) AND ([t7].[Active] = 1) AND (([t3].[test] IS NOT NULL) OR ([t6].[test] IS NOT NULL)) " & _
     "ORDER BY [t0].[IDReceiptHDR]"



     








    

    
    
    sSelect = comando1$ + comando2$
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    grid.AllowUserResizing = flexResizeColumns

    
    ' Asignar el recordset al FlexGrid
    Set grid.DataSource = Rs
    ' user1$ = Rs(0)
          
                        
                         
    Rs.Close
    
    
    
    
grid4.ColWidth(0) = 500
grid4.ColWidth(1) = 1200
grid4.ColWidth(2) = 1300

List7.Clear

grid4.Rows = grid.Rows

linea = 0
For t = 1 To grid.Rows - 1
  grid.row = t
  grid.col = 11
  n$ = RTrim(UCase(grid.Text))
  
  grid.col = 13
  cant = Val(grid.Text)
  
  If n$ <> "" Then
    If linea = 0 Then
       grid4.row = 0
       grid4.col = 1
       grid4.Text = "User"
       
       grid4.col = 2
       grid4.Text = "Fee Amount"
     
    End If
  
    ' checa si ya esta el usuario
    existe = 0
    For Y = 0 To List7.ListCount - 1
       n2$ = RTrim(UCase(Left(List7.List(Y), 20)))
       If n2$ = n$ Then
          cant2 = Val(Right(List7.List(Y), Len(List7.List(Y)) - 20))
          List7.RemoveItem Y
          existe = 1
          Exit For
       End If
    Next Y
    
    
    If existe = 0 Then
          List7.AddItem Format(n$, "!@@@@@@@@@@@@@@@@@@@@") + Space(1) + Format(cant, "##0.00")
    Else
          List7.AddItem Format(n$, "!@@@@@@@@@@@@@@@@@@@@") + Space(1) + Format(cant + cant2, "##0.00")
    End If
    
    
    linea = linea + 1
      
  End If
Next t
  

grid4.Rows = List7.ListCount

For t = 0 To List7.ListCount - 1
    n$ = Left(List7.List(t), 20)
    cant = Val(Right(List7.List(t), Len(List7.List(t)) - 20))
    
    grid4.row = t + 1
    
    grid4.col = 0
    grid4.Text = t + 1
        
    grid4.col = 1
    grid4.Text = n$
    
    grid4.col = 2
    grid4.Text = Str(cant)
Next t

grid4.row = t + 1
grid4.col = 0
grid4.Text = t - 1

msg.Visible = False
msg.Refresh
    






Exit Sub


start:


Dim xlApp2 As Excel.Application
Dim xlLibro2 As Excel.Workbook
Dim xlHoja2 As Excel.Worksheet
Dim varMatriz As Variant
Dim lngUltimaFila As Long
If btn_errors.Visible = False Then Exit Sub
lblmsg.Caption = "Please, wait a moment... loading errors"
msg.Visible = True
msg.Refresh

btn_errors.Visible = False
contador = 0

inicio:
n$ = "c:\goals\errorsmp.xlsx" '.csv"
grid4.Clear
List7.Clear

If Dir$(n$) = "" Then
  ' MsgBox "The file " + n$ + " has not been found", 64, "Attention"
   GoTo final
End If

'reset variables de los objetos
Set xlHoja2 = Nothing
Set xlLibro2 = Nothing
Set xlApp2 = Nothing

'abrir programa Excel
Set xlApp2 = New Excel.Application
'xl.Visible = True

'abrir el archivo Excel
'(archivo en la misma carpeta)
Set xlLibro2 = xlApp2.Workbooks.Open(FileName:=n$, ReadOnly:=True)

' Get the first worksheet.
 Set xlHoja2 = xlApp2.Worksheets(1)
' Set xlHoja = xlApp.Worksheets("bf")

'1. Si conoces el rango a leer
'varMatriz = xlHoja.Range(«A1:C10»).Value

'2. Si no conoces el rango
' lngUltimaFila = Columns("A:A").range("A65536").End(xlUp).row
lngUltimaFila = 2000

    ActiveCell.SpecialCells(xlLastCell).Select
    
    
    ultimafilax = ActiveCell.row
    ultimacolumnax = ActiveCell.Column

' sino carga el archivo entonces abrelo
If lngUltimaFila = 0 Then
  
  lngUltimaFila = lineas_err
  contador = contador + 1
  
  
  If contador >= 2 Then
     Exit Sub
  End If
  
Else

  lineas_err = lngUltimaFila

  
End If

continua:

 varMatriz = xlHoja2.range(xlHoja2.Cells(1, 1), xlHoja2.Cells(lngUltimaFila, 22))   ' cambie 10 por 22

'grid3.Clear
'utilizamos los datos…
'txtLlamadas.Text = varMatriz(10, 3)
grid4.Rows = lngUltimaFila + 2
grid4.cols = 3

cont = 0
For t = 1 To grid4.Rows - 2
  grid4.row = t - 1
  If t > 1 And t <= lngUltimaFila And varMatriz(t, 1) <> "" Then
     cont = cont + 1
     grid4.col = 0
     grid4.Text = cont
     
  End If
  For Y = 1 To 3
   grid4.col = Y
   grid4.Text = varMatriz(t, Y)
  Next Y
Next t
   
'grid3.Rows = cont + 2
ultima_linea = cont + 2
'lbl_count(1).Caption = grid3.Rows - 2
guarda_filas = cont + 2
grid4.Rows = guarda_filas

'cerramos el archivo Excel
xlLibro2.Close SaveChanges:=False
xlApp2.Quit

'reset variables de los objetos
Set xlHoja2 = Nothing
Set xlLibro2 = Nothing
Set xlApp2 = Nothing


grid4.ColWidth(0) = 500
grid4.ColWidth(1) = 1200
grid4.ColWidth(2) = 1300

List7.Clear

For t = 1 To grid4.Rows
  grid4.row = t
  grid4.col = 1
  n$ = grid4.Text
  grid4.col = 2
  cant = Val(grid4.Text)
  If n$ <> "" Then
    List7.AddItem Format(n$, "!@@@@@@@@@@@@@@@@@@@@") + Space(1) + Format(cant, "##0.00")
  End If
Next t
  


final:
msg.Visible = False
msg.Refresh

' separa_campos
btn_errors.Visible = True
End Sub

Private Sub btn_GI_Click()
On Error Resume Next
If txtdatefrom.Text = "" Then
  MsgBox "The start date is invalid", 64, "Attention"
  Exit Sub
End If

If txtdateto.Text = "" Then
  MsgBox "The end date is invalid", 64, "Attention"
  Exit Sub
End If

If txtdatepayday.Text = "" Then
  MsgBox "The pay date is invalid", 64, "Attention"
  Exit Sub
End If

lblmsg.Caption = "Please, wait a moment... loading GI data"
msg.Visible = True
msg.Refresh
carga_GI

msg.Visible = False
msg.Refresh

Exit Sub


Dim xlApp2 As Excel.Application
Dim xlLibro2 As Excel.Workbook
Dim xlHoja2 As Excel.Worksheet
Dim varMatriz As Variant
Dim lngUltimaFila As Long
If btn_GI.Visible = False Then Exit Sub


btn_GI.Visible = False
contador = 0


inicio:
n$ = "c:\goals\GI.xlsx" '.csv"
Grid2.Clear

If Dir$(n$) = "" Then
   MsgBox "The file " + n$ + " has not been found", 64, "Attention"
   GoTo final
End If

'reset variables de los objetos
Set xlHoja2 = Nothing
Set xlLibro2 = Nothing
Set xlApp2 = Nothing

'abrir programa Excel
Set xlApp2 = New Excel.Application
'xl.Visible = True

'abrir el archivo Excel
'(archivo en la misma carpeta)
Set xlLibro2 = xlApp2.Workbooks.Open(FileName:=n$, ReadOnly:=True)

' Get the first worksheet.
 Set xlHoja2 = xlApp2.Worksheets(1)
' Set xlHoja = xlApp.Worksheets("bf")

'1. Si conoces el rango a leer
'varMatriz = xlHoja.Range(«A1:C10»).Value

'2. Si no conoces el rango
' lngUltimaFila = Columns("A:A").range("A65536").End(xlUp).row
lngUltimaFila = 2000

    ActiveCell.SpecialCells(xlLastCell).Select
    
    
    ultimafilax = ActiveCell.row
    ultimacolumnax = ActiveCell.Column

' sino carga el archivo entonces abrelo
If lngUltimaFila = 0 Then
  
  lngUltimaFila = lineas_GI
  contador = contador + 1
  
  
  If contador >= 2 Then
     Exit Sub
  End If
  
Else

  lineas_GI = lngUltimaFila

  
End If

continua:

 varMatriz = xlHoja2.range(xlHoja2.Cells(1, 1), xlHoja2.Cells(lngUltimaFila, 22))   ' cambie 10 por 22

Grid2.Clear
'utilizamos los datos…
'txtLlamadas.Text = varMatriz(10, 3)
Grid2.Rows = lngUltimaFila + 2
Grid2.cols = 22

cont = 0
For t = 1 To Grid2.Rows - 2
  Grid2.row = t
  If t > 1 And t <= lngUltimaFila And varMatriz(t, 1) <> "" Then
     cont = cont + 1
     Grid2.col = 0
     Grid2.Text = cont
     
  End If
  For Y = 1 To 21
   Grid2.col = Y
   Grid2.Text = varMatriz(t, Y)
  Next Y
Next t
   
Grid2.Rows = cont + 2
lbl_count(1).Caption = Grid2.Rows - 2


'cerramos el archivo Excel
xlLibro2.Close SaveChanges:=False
xlApp2.Quit

'reset variables de los objetos
Set xlHoja2 = Nothing
Set xlLibro2 = Nothing
Set xlApp2 = Nothing


final:
msg.Visible = False
msg.Refresh

separa_campos_GI
btn_GI.Visible = True
CARGA_AGENTES
End Sub










Private Sub btn_INV_Click()
On Error Resume Next
If txtdatefrom.Text = "" Then
  MsgBox "The start date is invalid", 64, "Attention"
  Exit Sub
End If

If txtdateto.Text = "" Then
  MsgBox "The end date is invalid", 64, "Attention"
  Exit Sub
End If

If txtdatepayday.Text = "" Then
  MsgBox "The pay date is invalid", 64, "Attention"
  Exit Sub
End If

lblmsg.Caption = "Please, wait a moment... loading INV data"
msg.Visible = True
msg.Refresh
carga_inv
msg.Visible = False
msg.Refresh
Exit Sub


Dim xlApp2 As Excel.Application
Dim xlLibro2 As Excel.Workbook
Dim xlHoja2 As Excel.Worksheet
Dim varMatriz As Variant
Dim lngUltimaFila As Long
If btn_INV.Visible = False Then Exit Sub


btn_INV.Visible = False
contador = 0
lista_invoices30.Clear

inicio:
n$ = "c:\goals\INV30.xlsx" '.csv"
Grid3.Clear

If Dir$(n$) = "" Then
   MsgBox "The file " + n$ + " has not been found", 64, "Attention"
   GoTo final
End If

'reset variables de los objetos
Set xlHoja2 = Nothing
Set xlLibro2 = Nothing
Set xlApp2 = Nothing

'abrir programa Excel
Set xlApp2 = New Excel.Application
'xl.Visible = True

'abrir el archivo Excel
'(archivo en la misma carpeta)
Set xlLibro2 = xlApp2.Workbooks.Open(FileName:=n$, ReadOnly:=True)

' Get the first worksheet.
 Set xlHoja2 = xlApp2.Worksheets(1)
' Set xlHoja = xlApp.Worksheets("bf")

'1. Si conoces el rango a leer
'varMatriz = xlHoja.Range(«A1:C10»).Value

'2. Si no conoces el rango
' lngUltimaFila = Columns("A:A").range("A65536").End(xlUp).row
lngUltimaFila = 2000

    ActiveCell.SpecialCells(xlLastCell).Select
    
    
    ultimafilax = ActiveCell.row
    ultimacolumnax = ActiveCell.Column

' sino carga el archivo entonces abrelo
If lngUltimaFila = 0 Then
  
  lngUltimaFila = lineas_INV
  contador = contador + 1
  
  
  If contador >= 2 Then
     Exit Sub
  End If
  
Else

  lineas_INV = lngUltimaFila

  
End If

continua:

 varMatriz = xlHoja2.range(xlHoja2.Cells(1, 1), xlHoja2.Cells(lngUltimaFila, 22))   ' cambie 10 por 22

'grid3.Clear
'utilizamos los datos…
'txtLlamadas.Text = varMatriz(10, 3)
Grid3.Rows = lngUltimaFila + 2
Grid3.cols = 22

cont = 0
For t = 1 To Grid3.Rows - 2
  Grid3.row = t - 1
  If t > 1 And t <= lngUltimaFila And varMatriz(t, 1) <> "" Then
     cont = cont + 1
     Grid3.col = 0
     Grid3.Text = cont
     
  End If
  For Y = 1 To 21
   Grid3.col = Y
   Grid3.Text = varMatriz(t, Y)
  Next Y
Next t
   
'grid3.Rows = cont + 2
ultima_linea = cont + 2
'lbl_count(1).Caption = grid3.Rows - 2
guarda_filas = cont + 2


'cerramos el archivo Excel
xlLibro2.Close SaveChanges:=False
xlApp2.Quit

'reset variables de los objetos
Set xlHoja2 = Nothing
Set xlLibro2 = Nothing
Set xlApp2 = Nothing




' *****************************************************************************************

n$ = "c:\goals\INV.xlsx" '.csv"
' grid3.Clear

If Dir$(n$) = "" Then
   MsgBox "The file " + n$ + " has not been found", 64, "Attention"
   GoTo final
End If

'reset variables de los objetos
Set xlHoja2 = Nothing
Set xlLibro2 = Nothing
Set xlApp2 = Nothing

'abrir programa Excel
Set xlApp2 = New Excel.Application
'xl.Visible = True

'abrir el archivo Excel
'(archivo en la misma carpeta)
Set xlLibro2 = xlApp2.Workbooks.Open(FileName:=n$, ReadOnly:=True)

' Get the first worksheet.
 Set xlHoja2 = xlApp2.Worksheets(1)
' Set xlHoja = xlApp.Worksheets("bf")

'1. Si conoces el rango a leer
'varMatriz = xlHoja.Range(«A1:C10»).Value

'2. Si no conoces el rango
' lngUltimaFila = Columns("A:A").range("A65536").End(xlUp).row
lngUltimaFila = 2000

    ActiveCell.SpecialCells(xlLastCell).Select
    
    
    ultimafilax = ActiveCell.row
    ultimacolumnax = ActiveCell.Column

' sino carga el archivo entonces abrelo
If lngUltimaFila = 0 Then
  
  lngUltimaFila = lineas_INV
  contador = contador + 1
  
  
  If contador >= 2 Then
     Exit Sub
  End If
  
Else

  lineas_INV = lngUltimaFila

  
End If

continua2:

 varMatriz = xlHoja2.range(xlHoja2.Cells(1, 1), xlHoja2.Cells(lngUltimaFila, 22))   ' cambie 10 por 22

'grid3.Clear
'utilizamos los datos…
'txtLlamadas.Text = varMatriz(10, 3)
Grid3.Rows = lngUltimaFila + 2
'grid3.Cols = 22



'cont = 0
linea = ultima_linea - 1
For t = 2 To Grid3.Rows - 2
  Grid3.row = linea ' + 1
  If t > 1 And t <= lngUltimaFila And varMatriz(t, 1) <> "" Then
     cont = cont + 1
     Grid3.col = 0
     Grid3.Text = cont
     
  End If
  
  For Y = 1 To 21
   Grid3.col = Y
   Grid3.Text = varMatriz(t, Y)
  Next Y
  linea = linea + 1
Next t
   
Grid3.Rows = cont + 1

'lbl_count(2).Caption = grid3.Rows - 2


'cerramos el archivo Excel
xlLibro2.Close SaveChanges:=False
xlApp2.Quit

'reset variables de los objetos
Set xlHoja2 = Nothing
Set xlLibro2 = Nothing
Set xlApp2 = Nothing


final:
msg.Visible = False
msg.Refresh

separa_campos_inv
btn_INV.Visible = True
End Sub


Private Sub btn_NB_Click()
On Error Resume Next
    

calcula_NB



End Sub





Private Sub btn_UW_Click()
On Error Resume Next

GoTo excelaqui



 Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset
    
    
If grid5.Rows < 3 Then Exit Sub

Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
    
    
lblmsg.Caption = "Exporting the information to Excel"
mensaje.Visible = True
mensaje.Refresh

btnexcel.Visible = False
    
    
'Start a new workbook in Excel
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add
      
'Add data to cells of the first worksheet in the new workbook
Set oSheet = oBook.Worksheets(1)
'oSheet.range("A1").Value = "Last Name"
'oSheet.range("B1").Value = "First Name"
'oSheet.range("A1:B1").Font.Bold = True
'oSheet.range("A2").Value = "Doe"
'oSheet.range("B2").Value = "John"
'oSheet.range("A3").Value = "Vazquez"
'oSheet.range("B3").Value = "Maria"


'Create an array with 11 columns and (grid5.Rows + 4) rows
ReDim DataArray(grid5.Rows + 4, 1 To 11) As Variant
'Dim r As Integer



' ******************************************************************** AQUI EMPIEZA A PONER LA INFO *********************************************************


Dim valor_suspenso As Boolean


     
   Dim tabla$(2000, 13)
   Dim campo$(2, 20)
   
   Erase tabla$
    
   Set Rs = New ADODB.Recordset
   
   grid6.Clear
   grid6.Rows = 2
   grid6.Visible = False

   
   pagina = 0


   sSelect = "select username from employeeinfo where idemployee='" + usuario_DMV$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
   Rs.Open sSelect, base, adOpenUnspecified
  
    
   Username_dmv$ = RTrim(LTrim(Rs(0)))
   Rs.Close
   
   f1$ = Format(txtdatefrom.Text, "mm/dd/yyyy")
   f2$ = Format(txtdateto.Text, "mm/dd/yyyy")
   
   
   oficina1$ = ""
   If cbo_oficina.ListIndex >= 0 Then
     oficina1$ = UCase(LTrim(RTrim(Right(cbo_oficina.List(cbo_oficina.ListIndex), 25))))
   End If
   
    
   grid6.Visible = False
   grid6.Refresh
  
   If opcion_depto = 0 Then
     tipo_doc$ = "2,9"
     fecha_UW$ = "rechdr.DateReviewedForUWI"
    
   Else
     tipo_doc$ = "11"
     fecha_UW$ = "rechdr.DateReviewedForUWII"
    
   End If
  
  
   tag_errors$ = ""
   For Y = 0 To List1.ListCount - 1
    sSelect = "select idtypeerrortag from ErrorTagTypeCatalog where typeerrorname='" + UCase(List1.List(Y)) + "'"
    Rs.Open sSelect, base, adOpenUnspecified
   
    numerr$ = Rs(0)
    Rs.Close
    
    tag_errors$ = tag_errors$ + numerr$ + ","
   Next Y
    
   If Right(tag_errors$, 1) = "," Then
     tag_errors$ = Left(tag_errors$, Len(tag_errors$) - 1)
   End If
  
    
    
   If opcion_depto = 1 Then   ' 0=NB  1=ENDO
      tipolog$ = "9"  ' Error UW II
      departamento$ = "12"   ' ENDO
   Else
      tipolog$ = "8"   ' Error UW I
      departamento$ = "4"  ' NB
   End If
   
   
   
  ' carga todos los empleados o solo la seleccionada
  
   If cbo_users.ListIndex = -1 Then
      lista_de_usuarios$ = ""
      For t = 0 To cbo_users.ListCount - 1
         n$ = cbo_users.List(t)
    
         sSelect = "select idemployee from employeeinfo where username='" + n$ + "'"
    
         Rs.Open sSelect, base, adOpenUnspecified
         
         IdEmployee$ = RTrim(LTrim(Rs(0)))
         Rs.Close
      
         lista_de_usuarios$ = lista_de_usuarios$ + IdEmployee$ + ","
      
      Next t
        
      lista_de_usuarios$ = Left(lista_de_usuarios$, Len(lista_de_usuarios$) - 1)
   
   Else
  
  
      n$ = cbo_users.List(cbo_users.ListIndex)
      
      sSelect = "select idemployee from employeeinfo where username='" + n$ + "'"
    
      Rs.Open sSelect, base, adOpenUnspecified
      
      IdEmployee$ = RTrim(LTrim(Rs(0)))
      Rs.Close
      
      lista_de_usuarios$ = IdEmployee$
  
  
  
   End If
   
   
   
   
  ' carga todas las oficinas o solo la seleccionada
  
   If cbo_oficina.ListIndex = -1 Then
  
     lista_de_oficinas$ = ""
     For t = 0 To cbo_oficina.ListCount - 1
  
       n$ = LTrim(RTrim(Right(cbo_oficina.List(t), 30)))
      
        ' si esta activada la casilla de Omitir Phone Sales
       If chk_no_phonesales.Value = 1 Or chk_no_phonesales.Value = True Then
         If n$ = "JA - PHONE SALES" Then
          GoTo brinca_PS
         End If
       End If
      
       sSelect = "select idoffice from OfficesCatalog where Office='" + n$ + "'"
    
       Rs.Open sSelect, base, adOpenUnspecified
      
       idoficina$ = RTrim(LTrim(Rs(0)))
       Rs.Close
      
       lista_de_oficinas$ = lista_de_oficinas$ + idoficina$ + ","
brinca_PS:
      
     Next t
      
     lista_de_oficinas$ = Left(lista_de_oficinas$, Len(lista_de_oficinas$) - 1)
   
   Else
   
      n$ = LTrim(RTrim(Right(cbo_oficina.List(cbo_oficina.ListIndex), 30)))
    
      
      sSelect = "select idoffice from OfficesCatalog where Office='" + n$ + "'"
    
      Rs.Open sSelect, base, adOpenUnspecified
     
      lista_de_oficinas$ = RTrim(LTrim(Rs(0)))
      Rs.Close
   
   
   End If
   
   
   
 
  
  
  
   sSelect = "select polhdr.IdPoliciesHDR as PolicyID, polhdr.IdCustomer as CustID, polhdr.PolicyNumber,cast(rechdr.date as date) as Date, " & _
  "empusr.Username,rechdr.IDReceiptHDR as Receipt, ofc.Office , ii.InvoiceItemName, " & _
  "recdtl.amount, rechdr.DateReviewedForUWI, polnotes.Note, tipoerror.TypeErrorName , polerror.isdeduction, polerror.exception, polnotes.IdEmployee as UW_Agent, polnotes.CreatedDate " & _
  "From ReceiptsHDR rechdr " & _
  "inner join PoliciesHDR polhdr on polhdr.IdPoliciesHDR=rechdr.IdPoliciesHDR " & _
  "inner join Customers cus on cus.IdCustomer=polhdr.IdCustomer " & _
  "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
  "inner join ProgramsCatalog prog on prog.IdProgram=polhdr.IdProgram " & _
  "inner join EmployeeInfo empusr on empusr.IDEmployee=rechdr.IdEmployeeUSR " & _
  "inner join OfficesCatalog ofc on ofc.IdOffice=rechdr.IdOffice " & _
  "inner join ReceiptsDTL recdtl on recdtl.IdReceiptHDR = rechdr.IDReceiptHDR " & _
  "inner join InvoiceItemCatalog ii on ii.IdInvoiceItem=recdtl.IdInvoiceItem " & _
  "left join EmplGoalsCalc goal on goal.IdReceiptHDR=rechdr.IDReceiptHDR " & _
  "inner join PolicyErrorTagRel polerror on polerror.IdPoliciesHDR=polhdr.IdPoliciesHDR " & _
  "inner join PolicyErrorTagNotesRel polnotes on polnotes.IdPolicyErrorTagRel=polerror.IdPolicyErrorTagRel " & _
  "inner join ErrorTagTypeCatalog tipoerror on tipoerror.IdTypeErrorTag=polerror.IdTypeErrorTag " & _
  "where cast (rechdr.Date as date) >= '" + f1$ + "' and cast (rechdr.Date as date) <= '" + f2$ + "' and empusr.IdEmployee in (" + lista_de_usuarios$ + ") " & _
  "and ii.IdInvoiceItem in (" + tipo_doc$ + ") and polerror.isdeduction='1' and (polerror.exception is null or polerror.exception='0') " & _
  "and ofc.IdOffice in (" + lista_de_oficinas$ + ") and polerror.IdTypeErrorTag in (" + tag_errors$ + ") and rechdr.Active=1 and polerror.active=1 " & _
  "order by rechdr.IDReceiptHDR, tipoerror.TypeErrorName, polnotes.CreatedDate, polnotes.Note, ofc.Office"
  
  
  
  
  
  ' and empusr.IDEmployee in (" + lista_de_usuarios$ + ")   se quito esta linea
  
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
   
        
     ' Permitir redimensionar las columnas
    grid6.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid6.DataSource = Rs
                         
    Rs.Close
    
    
    
    
    ' elimina los registros repetidos
    ' ********************************************
    grid7.Clear
    grid7.Rows = grid6.Rows
    grid7.cols = grid6.cols
    
    filas = 0
    For t = 1 To grid6.Rows
       
       grid6.row = t
       
       grid6.col = 6
       recibo_actual$ = grid6.Text
       
       grid6.col = 11
       comentario_actual$ = UCase(grid6.Text)
       
       grid6.col = 12
       error_actual$ = UCase(grid6.Text)
       
           
       comentario_siguiente$ = ""
              
       grid6.row = t + 1
       
       grid6.col = 6
       recibo_siguiente$ = grid6.Text
       
       grid6.col = 11
       comentario_siguiente$ = UCase(grid6.Text)
       
       grid6.col = 12
       error_siguiente$ = UCase(grid6.Text)
       
       
       
       
       
       If filas = 0 Then
          'GoTo graba_fila
       End If
       
            
       If recibo_actual$ = recibo_siguiente$ And comentario_actual$ = comentario_siguiente$ And error_actual$ = error_siguiente$ Then
           GoTo ve_siguiente
       Else
           GoTo graba_fila
       End If
       
       
       GoTo ve_siguiente
       
graba_fila:
       filas = filas + 1
       grid6.row = t
       grid7.row = filas
       
       For Y = 1 To grid6.cols - 1
          grid6.col = Y
          grid7.col = Y
          grid7.Text = grid6.Text
       Next Y
       
       
       
       
ve_siguiente:
       
    Next t
    
    
       
       
       
       filas = filas + 1
       grid6.row = t - 1
       grid7.row = filas
       
       For Y = 1 To grid6.cols - 1
          grid6.col = Y
          grid7.col = Y
          grid7.Text = grid6.Text
       Next Y
       
       
    
    
    grid6.Clear
    grid7.Rows = filas + 1
    grid6.Rows = filas + 1
    grid6.cols = grid7.cols
    
    For t = 0 To grid7.Rows
       grid6.row = t
       grid7.row = t
       For Y = 1 To grid7.cols - 1
           grid7.col = Y
           grid6.col = Y
           grid6.Text = grid7.Text
       Next Y
    Next t
    
    
   
    
    ' ordena_tabla6
    
    Erase tabla$
    
    
    
   fila = 0
   
   For t = 1 To grid6.Rows - 1
        
        existe = 0
        grid6.row = t
        
        grid6.col = 6
        recibo$ = grid6.Text
        
        commen$ = ""
        grid6.col = 11
        commen$ = grid6.Text
        
        grid6.col = 12
        tipo_error$ = grid6.Text
        
        For Y = 1 To fila
            If tabla$(Y, 0) = recibo$ And tabla$(Y, 7) = tipo_error$ And commen$ <> "" Then
                     
                     tabla$(Y, 8) = tabla$(Y, 8) + commen$ + ". "
                     existe = 1
                     Exit For
            End If
        Next Y
  
 
        If existe = 0 Then
            fila = fila + 1
            tabla$(fila, 0) = recibo$
            
            grid6.col = 5
            tabla$(fila, 1) = grid6.Text  ' Agente
            
            grid6.col = 6
            tabla$(fila, 2) = grid6.Text ' recibo
            
           
            
            grid6.col = 4
            tabla$(fila, 3) = Format(grid6.Text, "mm/dd/yyyy")  ' fecha
            
                         
            grid6.col = 10
            tabla$(fila, 4) = Format(grid6.Text, "mm/dd/yyyy")  ' fecha de revision
            
            
            
            
            grid6.col = 15
            id_emp_uw$ = grid6.Text
            
            sSelect = "select username from employeeinfo where idemployee='" + id_emp_uw$ + "'"
    
            Rs.Open sSelect, base, adOpenUnspecified
           
            Username_UW$ = RTrim(LTrim(Rs(0)))
            Rs.Close
            
            
            tabla$(fila, 5) = Username_UW$  ' UW agente
            
            
            
            
            grid6.col = 9
            tabla$(fila, 6) = Format(grid6.Text, "$###,##0.00")  ' GI
            
            grid6.col = 12
            tabla$(fila, 7) = grid6.Text  ' Tipo de ERROR
            
            grid6.col = 11
            tabla$(fila, 8) = grid6.Text + ". " ' Comentarios
            
            grid6.col = 13
            tabla$(fila, 9) = grid6.Text   ' ES DEDUCCION
            
            grid6.col = 14
            tabla$(fila, 10) = grid6.Text   ' Exception
            
            grid6.col = 7
            tabla$(fila, 11) = grid6.Text   ' oficina
            
            
            grid6.col = 2
            tabla$(fila, 12) = grid6.Text   ' Cust_id
            
            
            grid6.col = 3
            tabla$(fila, 13) = grid6.Text   ' poliza
            
            
            
            
        Else
            
            ' tabla$(fila, 8) = tabla$(Y, 8) ' Comentarios
            
            
        End If
        
   Next t
 
 
 
 
 
 
   ' depura la tabla y elimina los que no tienen comentario
   ' -------------------------------------------------------------
   
   
   For Y = 0 To fila
      sSelect = "select * from PolicyErrorTagRel where IdReceiptHDR='" + tabla$(Y, 2) + "'"
      Rs.Open sSelect, base, adOpenUnspecified
      
        
     grid7.AllowUserResizing = flexResizeColumns
     Set grid7.DataSource = Rs
                         
     Rs.Close
     
     
     If grid7.Rows <= 1 Then
           For z = 0 To 13
              tabla(Y, z) = ""
           Next z
     End If
    
    
     
    
   Next Y
   

   
   ' ------------ TERMINA DE DEPURAR -----------------------------
 
 
    
   grid7.Clear
  
   grid7.Rows = fila + 1
   grid7.cols = 14
   
   
   For t = 1 To fila
      
      grid7.row = t
      
      For Y = 0 To 13
         grid7.col = Y
         grid7.Text = tabla$(t, Y)
      Next Y
   Next t
    
    
   
   
 MAX_filas = grid7.Rows
 List5.Clear
 
 
 ' carga las oficinas con errorres
 List4.Clear
 For Y = 1 To grid7.Rows - 1
    grid7.row = Y
    grid7.col = 11
    uw_oficina$ = UCase(grid7.Text)
    existe = 0
    For w = 0 To List4.ListCount - 1
       If List4.List(w) = uw_oficina$ Then
            existe = 1
            Exit For
       End If
    Next w
 
    If existe = 0 Then
          List4.AddItem uw_oficina$
    End If
    
 Next Y
 
 
 grand_total = 0
 grand_suma_total = 0
 
 
 
 
 For X = 0 To List4.ListCount - 1
 
    
     
  OFICINA_impresa$ = UCase(List4.List(X))
  
  'Printer.Print Space(3) + OFICINA_impresa$
  
  
  linea = linea + 2
  
  suma_total = 0
  Total = 0
 
   
  For t = 1 To grid7.Rows - 1
   
   
   
   grid7.row = t
   
   grid7.col = 1
   UW_Agente$ = Format(Left(grid7.Text, 11), "!@@@@@@@@@@@")
   
   
   grid7.col = 2
   uw_recibo$ = Format(Left(grid7.Text, 6), "@@@@@@")
   
   grid7.col = 3
   uw_Efec_Date$ = Format(grid7.Text, "mm/dd/yyyy")
   
   grid7.col = 4
   uw_Rev_date$ = Format(grid7.Text, "mm/dd/yyyy")
   
   If uw_Rev_date$ = "" Then
      uw_Rev_date$ = Space(10)
   End If
   
   
   grid7.col = 5
   uw_Inspector$ = Format(Left(grid7.Text, 10), "!@@@@@@@@@@")
   
   grid7.col = 6
   uw_cantidad$ = Format(Format(grid7.Text, "$###,##0.00"), "@@@@@@@@@@@")
   cantidad = Val(Format(uw_cantidad$, "000000.00"))
   
   grid7.col = 7
   uw_error$ = Format(Left(grid7.Text, 22), "!@@@@@@@@@@@@@@@@@@@@@@")
   
   grid7.col = 8
   uw_Comentario$ = grid7.Text
   
   
   
   
   
   grid7.col = 11
   uw_oficina$ = UCase(grid7.Text)
   
   If uw_oficina$ <> OFICINA_impresa$ Then
      GoTo brinca
   End If
   
   
   grid7.col = 12
   uw_custid$ = Format(UCase(grid7.Text), "@@@@@@")
   
   
   grid7.col = 13
   uw_poliza$ = Format(UCase(grid7.Text), "!@@@@@@@@@@@@@@@@@@@@")
   
   
    
    
   grid7.col = 2
   folio$ = grid7.Text
   
   existe = 0
   r1$ = Format(folio$, "00000000") + Space(1) + Format(cantidad, "00000.00")
   For q = 0 To List5.ListCount - 1
      If Val(Left(List5.List(q), 8)) = Val(folio$) Then
         existe = 1
         Exit For
      End If
   Next q
   
   
   
   If existe = 0 And folio$ <> "" Then
      List5.AddItem r1$
      suma_total = suma_total + cantidad
      Total = Total + 1
   
   End If
      
      
      
      
   
   
   
   
   num_lineas_tipo_error = 1
   
   If Len(uw_error$) > 24 Then
       num1 = Len(uw_error$) / 24
       num2 = Int(Len(uw_error$) / 24)
       residuo = num1 - num2
       
       If residuo > 0 Then
          num_lineas_tipo_error = num2 + 1
       Else
          num_lineas_tipo_error = num2
       End If
   End If
   
   
   
   num_lineas_comentario = 1
   
   If Len(uw_Comentario$) > 36 Then
       num1 = Len(uw_Comentario$) / 36
       num2 = Int(Len(uw_Comentario$) / 36)
       residuo = num1 - num2
       If residuo > 0.5 Then
          num_lineas_comentario = num2 + 2
       ElseIf residuo > 0 Then
          num_lineas_comentario = num2 + 1
          
       Else
          num_lineas_comentario = num2
       End If
   End If
   
   
   
   ' checa cuantas lineas se agregaran
   total_lin = 1
   
   If num_lineas_tipo_error > 1 Or num_lineas_comentario > 1 Then
       
       If num_lineas_tipo_error > num_lineas_comentario Then
              total_lin = num_lineas_tipo_error '+ 1
       ElseIf num_lineas_tipo_error < num_lineas_comentario Then
              total_lin = num_lineas_comentario '+ 1
       End If
   
   End If
   
   
       
   
   Erase campo$
   
   linea = linea + total_lin

  ' If num_lineas_tipo_error = 1 And num_lineas_comentario = 1 Then
      
      If RTrim(uw_Comentario$) = "." Then
        GoTo brinca
      End If
      
      
      r$ = Space(3) + UW_Agente$ + " " + uw_custid$ + "  " + uw_recibo$ + "  " + uw_poliza$ + "  " + Format(uw_Efec_Date$, "@@@@@@@@@@") + "  " + Format(uw_Rev_date$, "@@@@@@@@@@") + "  " + uw_Inspector$ + Space(1) + uw_cantidad$ + Space(1) + Format(Left(uw_error$, 22), "!@@@@@@@@@@@@@@@@@@@@@@") + Space(2) + uw_Comentario$
      If LTrim(RTrim(r$)) <> "" And folio$ <> "" Then
      
      ' r$ = Space(3) + UW_Agente$ + " " + uw_custid$ + "  " + uw_recibo$ + "  " + uw_poliza$ + "  " + Format(uw_Efec_Date$, "@@@@@@@@@@") + "  " + Format(uw_Rev_date$, "@@@@@@@@@@") + "  " + uw_Inspector$ + Space(1) + uw_cantidad$ + Space(1) + Format(Left(uw_error$, 22), "!@@@@@@@@@@@@@@@@@@@@@@") + Space(2) + uw_Comentario$
      
   
       num = num + 1
       DataArray(num, 1) = UW_Agente$
            
       DataArray(num, 2) = uw_custid$
      
       DataArray(num, 3) = uw_recibo$
      
       DataArray(num, 4) = uw_poliza$
      
       DataArray(num, 5) = uw_Efec_Date$
      
       DataArray(num, 6) = uw_Rev_date$
      
       DataArray(num, 7) = uw_Inspector$
      
       DataArray(num, 8) = uw_cantidad$
      
       DataArray(num, 9) = uw_error$
      
       DataArray(num, 10) = uw_Comentario$
      
       DataArray(num, 11) = uw_oficina$
       
       linea = linea + 1
      
      End If
   
   
 
   
   
   
   
   
   
brinca:

 Next t
 
 
  
  r$ = Space(3) + UW_Agente$ + " " + uw_custid$ + "  " + uw_recibo$ + "  " + uw_poliza$ + "  " + Format(uw_Efec_Date$, "@@@@@@@@@@") + "  " + Format(uw_Rev_date$, "@@@@@@@@@@") + "  " + uw_Inspector$ + Space(1) + uw_cantidad$ + Space(1) + Format(Left(uw_error$, 22), "!@@@@@@@@@@@@@@@@@@@@@@") + Space(2) + uw_Comentario$
  If LTrim(RTrim(r$)) <> "" And folio$ <> "" Then
  
      linea = linea + 2
 'Printer.Print Space(3) + "TOTAL: " + Format(Format(total, "##0"), "!@@@") + Space(2) + "Total BF: " + Format(suma_total, "$###,##0.00")
 
      num = num + 1
      DataArray(num, 1) = "TOTAL: "
            
      DataArray(num, 2) = Format(Format(Total, "##0"), "!@@@")
      
      DataArray(num, 3) = ""
      
      DataArray(num, 4) = ""
      
      DataArray(num, 5) = "Total BF: "
      
      DataArray(num, 6) = Format(suma_total, "$###,##0.00")
      
      DataArray(num, 7) = ""
      
      DataArray(num, 8) = ""
      
      DataArray(num, 9) = ""
      
      DataArray(num, 10) = ""
      
      DataArray(num, 11) = ""
      
      
      
 
 
      linea = linea + 2
 
      grand_total = grand_total + Total
      grand_suma_total = grand_suma_total + suma_total
 
  End If
 
 

Next X



' printer.Print Space(3) + ">>> Grand Total: " + Format(grand_total, "###,##0") + Space(8) + "Grand Total BF: " + Format(grand_suma_total, "$###,##0.00")
linea = linea + 3

num = num + 1
      DataArray(num, 1) = ">>> Grand Total: "
            
      DataArray(num, 2) = Format(grand_total, "###,##0")
      
      DataArray(num, 3) = ""
      
      DataArray(num, 4) = ""
      
      DataArray(num, 5) = "Grand Total BF: "
      
      DataArray(num, 6) = Format(grand_suma_total, "$###,##0.00")
      
      DataArray(num, 7) = ""
      
      DataArray(num, 8) = ""
      
      DataArray(num, 9) = ""
      
      DataArray(num, 10) = ""
      
      DataArray(num, 11) = ""
      
 

       
'End If











' ************************************************************************************************************************************************************



'Add headers to the worksheet on row 1
Set oSheet = oBook.Worksheets(1)
oSheet.range("A1:K1").Value = Array("Agent", "Cust-ID", "Receipt", "Policy#", "Efect_date", "Rev_date", "UW", "Amount", "Error", "Comment", "Office")

'Transfer the array to the worksheet starting at cell A2, -- I changed A2 by A1

oSheet.range("A2").Resize(grid5.Rows + 4, 11).Value = DataArray





'GoTo brincaesto

cd1.DialogTitle = "Save File"
    cd1.InitDir = "c:\UW"
    cd1.Filter = "Excel Files (*.xls)|*.xls|All " & _
        "Files (*.*)|*.*"
    cd1.FilterIndex = 1
    cd1.flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNExplorer
    cd1.CancelError = True

  cd1.ShowSave
  n$ = cd1.FileName
  
  
If n$ = "" Then
  btnexcel.Visible = True
  mensaje.Visible = False
  oExcel.Quit
  Exit Sub
End If


brincaesto:





'Save the Workbook and Quit Excel

oExcel.DisplayAlerts = False
'n$ = "c:\uw\temp.csv"
oBook.SaveAs FileName:=n$, FileFormat:=xlExcel8    'xlExcel8=excel 97-2003  xlCSV para CVS
wb.Close (True)
oExcel.DisplayAlerts = True
oExcel.Quit

'Name "c:\uw\temp.csv" As n$


mensaje.Visible = False
btnexcel.Visible = True









 Exit Sub
 
 
 
 
' carga datos desde excel
excelaqui:
 
 

Dim xlApp2 As Excel.Application
Dim xlLibro2 As Excel.Workbook
Dim xlHoja2 As Excel.Worksheet
Dim varMatriz As Variant
Dim lngUltimaFila As Long
If btn_UW.Visible = False Then Exit Sub
lblmsg.Caption = "Please, wait a moment... loading UW DED."
msg.Visible = True
msg.Refresh

btn_UW.Visible = False
contador = 0

inicio:
n$ = "c:\goals\errorsuw.xlsx" '.csv"
grid5.Clear
List8.Clear

If Dir$(n$) = "" Then
   'MsgBox "The file " + n$ + " has not been found", 64, "Attention"
   GoTo final
End If

'reset variables de los objetos
Set xlHoja2 = Nothing
Set xlLibro2 = Nothing
Set xlApp2 = Nothing

'abrir programa Excel
Set xlApp2 = New Excel.Application
'xl.Visible = True

'abrir el archivo Excel
'(archivo en la misma carpeta)
Set xlLibro2 = xlApp2.Workbooks.Open(FileName:=n$, ReadOnly:=True)

' Get the first worksheet.
 Set xlHoja2 = xlApp2.Worksheets(1)
' Set xlHoja = xlApp.Worksheets("bf")

'1. Si conoces el rango a leer
'varMatriz = xlHoja.Range(«A1:C10»).Value

'2. Si no conoces el rango
' lngUltimaFila = Columns("A:A").range("A65536").End(xlUp).row
lngUltimaFila = 2000

    ActiveCell.SpecialCells(xlLastCell).Select
    
    
    ultimafilax = ActiveCell.row
    ultimacolumnax = ActiveCell.Column

' sino carga el archivo entonces abrelo
If lngUltimaFila = 0 Then
  
  lngUltimaFila = lineas_uw
  contador = contador + 1
  
  
  If contador >= 2 Then
     Exit Sub
  End If
  
Else

  lineas_uw = lngUltimaFila

  
End If

continua:

 varMatriz = xlHoja2.range(xlHoja2.Cells(1, 1), xlHoja2.Cells(lngUltimaFila, 22))   ' cambie 10 por 22

'grid3.Clear
'utilizamos los datos…
'txtLlamadas.Text = varMatriz(10, 3)
grid5.Rows = lngUltimaFila + 2
grid5.cols = 4

cont = 0
For t = 1 To grid5.Rows - 2
  grid5.row = t - 1
  If t > 1 And t <= lngUltimaFila And varMatriz(t, 1) <> "" Then
     cont = cont + 1
     grid5.col = 0
     grid5.Text = cont
     
  End If
  For Y = 1 To 3
   grid5.col = Y
   grid5.Text = varMatriz(t, Y)
  Next Y
Next t
   
'grid3.Rows = cont + 2
ultima_linea = cont + 2
'lbl_count(1).Caption = grid3.Rows - 2
guarda_filas = cont + 2
grid5.Rows = guarda_filas

'cerramos el archivo Excel
xlLibro2.Close SaveChanges:=False
xlApp2.Quit

'reset variables de los objetos
Set xlHoja2 = Nothing
Set xlLibro2 = Nothing
Set xlApp2 = Nothing


grid5.ColWidth(0) = 500
grid5.ColWidth(1) = 1100
grid5.ColWidth(2) = 800
grid5.ColWidth(3) = 1200


List8.Clear


For t = 1 To grid5.Rows
  grid5.row = t
   grid5.col = 2
   n$ = RTrim(LTrim(UCase(grid5.Text)))
   
   
   grid5.col = 3
   cant1 = Val(Format(grid5.Text, "00000.00"))
  
   existe = 0
   c = 0
   If n$ <> "" Then
      If List8.ListCount = 0 Then existe = 2
      For Y = 0 To List8.ListCount - 1
         n2$ = UCase(Left(List8.List(Y), 3))
         c = Val(Mid(List8.List(Y), 5, 3))
         cant = Val(Right(List8.List(Y), 7))
         
         If n2$ = n$ Then
            existe = 1
            cantf = cant + cant1
            c = c + 1
            
            List8.RemoveItem Y
            Exit For
         End If
      
         
      Next Y
      
   End If
   
  If existe = 2 And List8.ListCount = 0 Then c = 1
  If existe = 2 And List8.ListCount = 1 Then c = 1
   
   
  If existe = 1 And c > 0 Then
    cantf = cant + cant1
    List8.AddItem Format(n$, "@@@") + Space(1) + Format(c, "000") + " " + Format(cantf, "0000.00")
  ElseIf existe = 2 And c > 0 Then
    cantf = cant + cant1
    List8.AddItem Format(n$, "@@@") + Space(1) + Format(c, "000") + " " + Format(cantf, "0000.00")
  ElseIf existe = 0 And c > 0 Then
    cantf = cant1
    c = 1
    List8.AddItem Format(n$, "@@@") + Space(1) + Format(c, "000") + " " + Format(cantf, "0000.00")
  End If
  
  cantf = 0
  cant = 0
  
Next t



final:
msg.Visible = False
msg.Refresh


' separa_campos
btn_UW.Visible = True
End Sub

Private Sub btnacercade_Click()
On Error Resume Next

msg1$ = "           G O A L S   " + Chr$(13)
msg1$ = msg1$ + Chr$(13)
msg1$ = msg1$ + "Copyright(C) 2021-2024 by the author" + Chr$(13)
msg1$ = msg1$ + "JUST AUTO INSURANCE" + Chr$(13) + Chr$(13)

msg1$ = msg1$ + "Warning: This computer program is protected by" + Chr$(13)
msg1$ = msg1$ + "copyright law and international treaties." + Chr$(13)
msg1$ = msg1$ + "Unauthorized reproduction or distribution of this" + Chr$(13)
msg1$ = msg1$ + "program, or any portion of it, may result in severe" + Chr$(13)
msg1$ = msg1$ + "civil and criminal panalties, and will be prosecuted" + Chr$(13)
msg1$ = msg1$ + "to the maximum extent possible under law." + Chr$(13)
msg1$ = msg1$ + Chr$(13) + Chr$(13)
msg1$ = msg1$ + "Created by: Hector Navarro"
MsgBox msg1$, 64, "About... "
Text1.SetFocus
End Sub



Private Sub btncargafile_Click()
On Error Resume Next

n$ = ""
cd1.FileName = ""
cd1.DialogTitle = "Open File"
    cd1.InitDir = "c:\goals"
    cd1.Filter = "ALL Files (*.xlsx)|*.xlsx|XLS Files (*.xls)"
    cd1.FilterIndex = 1
    cd1.flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNExplorer
    cd1.CancelError = True

  cd1.ShowOpen
  n$ = cd1.FileName
  
  If n$ = "" Then
    Exit Sub
  End If
  
  txtfile.Text = n$

End Sub

Private Sub btncargar_excel_Click()
On Error Resume Next
If txtdatefrom.Text = "" Then
  MsgBox "The start date is invalid", 64, "Attention"
  Exit Sub
End If

If txtdateto.Text = "" Then
  MsgBox "The end date is invalid", 64, "Attention"
  Exit Sub
End If

If txtdatepayday.Text = "" Then
  MsgBox "The pay date is invalid", 64, "Attention"
  Exit Sub
End If

txtresultado.Text = ""

lblmsg.Caption = "Please, wait a moment... loading all data"
msg.Visible = True
msg.Refresh
lblexcel.Visible = False
Shape10.Visible = False

'Grid1.Visible = False
'Grid2.Visible = False
'grid3.Visible = False
'grid4.Visible = False
'grid5.Visible = False
'grid6.Visible = False


Grid1.Clear
Grid2.Clear
Grid3.Clear
grid4.Clear
lista_invoices30.Clear
List7.Clear
List8.Clear
List12.Clear

 lbl_count(0).Caption = "00"
  lbl_count(1).Caption = "00"

btncargar_excel.Visible = False

btn_NB_Click

btn_GI_Click


btn_INV_Click

calcula_invoices

btn_errors_Click
btn_UW_Click
btn_dmv_Click

calcula_grandtotal

lblexcel.Visible = True
Shape10.Visible = True
btncargar_excel.Visible = True

'Grid1.Visible = True
'Grid2.Visible = True
'grid3.Visible = True
'grid4.Visible = True
'grid5.Visible = True
'grid6.Visible = True


msg.Visible = False
msg.Refresh

lblregistros.Caption = Format(lista_agentes.ListCount, "##0")

End Sub

Private Sub btnclose_Click()
On Error Resume Next
Picture2.Visible = False

End Sub

Private Sub btndeduction_Click()
On Error Resume Next

r$ = InputBox("Type the new authorized deduction:", "NEW DEDUCTION")
r$ = UCase$(r$)
If LTrim(r$) = "" Then Exit Sub


lbldeduction.Caption = r$

chklock1.Enabled = True
'chklock1.Value = True


End Sub

Private Sub btnDMV_Click()
On Error Resume Next
Load forma_DMV
forma_DMV.Show 1

End Sub

Private Sub btnend_Click()
On Error Resume Next
Dim file_name As String


    file_name = "c:\goals" 'App.Path
    
    If Right$(file_name, 1) <> "\" Then file_name = file_name & "\"
    file_name = file_name & "test.txt"
    SaveTreeViewIntoFile file_name, TreeView1
    
    base.Close
    
End
End Sub

Private Sub btnexcel_Click()
On Error Resume Next

If lista_agentes.ListCount = 0 Then Exit Sub

Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
    
    
lblmsg.Caption = "Please, wait a moment... transferring data to Excel"

msg.Visible = True
msg.Refresh




    
    
'Start a new workbook in Excel
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add
      
'Add data to cells of the first worksheet in the new workbook
Set oSheet = oBook.Worksheets(1)
'oSheet.range("A1").Value = "Last Name"
'oSheet.range("B1").Value = "First Name"
'oSheet.range("A1:B1").Font.Bold = True
'oSheet.range("A2").Value = "Doe"
'oSheet.range("B2").Value = "John"
'oSheet.range("A3").Value = "Vazquez"
'oSheet.range("B3").Value = "Maria"


'Create an array with 20 columns and 100 rows
Dim DataArray(1 To 100, 1 To 20) As Variant
Dim r As Integer


grandtotal = 0
num = 0
For t = 0 To lista_agentes.ListCount - 1
   
   lista_agentes.ListIndex = t
   grandtotal = grandtotal + Val(Format(lbltotal_invoices.Caption, "000000.00"))

   n$ = lblinitials.Caption + " - " + RTrim(lblfull_name.Caption)
   
   
   ' detecta si es manager y lo marca
   existe = 0
   For Y = 0 To lista_managers.ListCount - 1
      n2$ = UCase(RTrim(Left(lista_managers.List(Y), 20)))
      If n2$ = agente$ Then
          puesto$ = Right(lista_managers.List(Y), Len(lista_managers.List(Y)) - 20)
          If LTrim(UCase(RTrim(puesto$))) = "MANAGER" Then
               existe = 1
               Exit For
          End If
          
          If LTrim(UCase(RTrim(puesto$))) = "MONTERREY" Then
               existe = 1
               Exit For
          End If
          
          If LTrim(UCase(RTrim(puesto$))) = "COMMERCIAL" Then
               existe = 1
               Exit For
          End If
          
      End If
   Next Y



   If existe = 0 Then
    num = num + 1
    DataArray(num, 1) = " "
    DataArray(num, 2) = n$
    DataArray(num, 3) = txtdatefrom.Text
    DataArray(num, 4) = txtdateto.Text
    DataArray(num, 5) = txtdatepayday.Text
    DataArray(num, 6) = "0"
    DataArray(num, 7) = "0"
    DataArray(num, 8) = "$0.00"
    DataArray(num, 9) = lblnb.Caption
    DataArray(num, 10) = lblbf.Caption
    DataArray(num, 11) = lblinvoice.Caption
    DataArray(num, 12) = Format(lbldeduction.Caption, "$###,##0.00")
    DataArray(num, 13) = Format(lblnb_deduc.Caption, "#0")
    DataArray(num, 14) = lbltotal_NB.Caption
    DataArray(num, 15) = lbltotal_bf.Caption
    DataArray(num, 16) = lblporcentaje.Caption
    DataArray(num, 17) = lblcommission.Caption
    DataArray(num, 18) = "$0.00"
    DataArray(num, 19) = lblid.Caption
    
    r1$ = RTrim(Left(txtnotes.Text, Len(txtnotes.Text) - 1))
    
    
    DataArray(num, 20) = RTrim(r1$)
   
   End If
   r1$ = ""
Next t

'Add headers to the worksheet on row 1
Set oSheet = oBook.Worksheets(1)
oSheet.range("A1:T1").Value = Array(" ", "Employee", "From", "To", "Pay Day", "R. Hour", "OT", "A. Hour", "NB", "BF", "Invoice", "Deduction", "NB Deduc.", "Total NB", "Total BF", "%", "Comission", "Total", "ID", "Notes")

'Transfer the array to the worksheet starting at cell A2
oSheet.range("A2").Resize(100, 20).Value = DataArray


lblgrandtotal.Caption = Format(grandtotal, "$###,##0.00")



'Save the Workbook and Quit Excel
oBook.SaveAs "C:\goals\HR.xlsx"
oExcel.Quit


msg.Visible = False
msg.Refresh

End Sub

Private Sub Btnletra_Click(Index As Integer)
 On Error Resume Next

letra$ = Btnletra(Index).Caption
If Index = 26 Then letra$ = "-"

CARGA_AGENTES
Text1.SetFocus


End Sub

Private Sub btnload_Click()
On Error Resume Next
Load HR
HR.Show 1

End Sub

Private Sub btnNB_deduction_Click()
On Error Resume Next
r$ = InputBox("Type the new authorized NB deduction:", "NEW NB DEDUCTION")
r$ = UCase$(r$)
If LTrim(r$) = "" Then Exit Sub


lblnb_deduc.Caption = r$

chklock2.Enabled = True
'chklock2.Value = True

End Sub

Private Sub btnorden_Click()
On Error Resume Next
If btnorden.Value = False Then
 Grid2.col = 0
 
Else
 Grid2.col = 3
 
End If


Grid2.Sort = flexSortGenericAscending


End Sub

Private Sub btnporcentaje_Click()
On Error Resume Next
r$ = InputBox("Type the new authorized percentage:", "More information")
r$ = UCase$(r$)
If LTrim(r$) = "" Then Exit Sub


lblporcentaje.Caption = r$
If Right(r$, 1) <> "%" Then r$ = r$ + "%"

lblporcentaje.Caption = r$
'chklock3.Enabled = True
chklock3.Value = True

porcentaje(lista_agentes.ListIndex) = r$

lista_agentes_Click
'calcula_todo

End Sub

Private Sub btnprint_Click()
On Error Resume Next
If lista_agentes.ListCount = 0 Then Exit Sub


X$ = MsgBox("Do you want to print the report?", 4, "Attention")
If X$ = "7" Then Exit Sub


lblmsg.Caption = "Please, wait a moment... preparing data for printing"

msg.Visible = True
msg.Refresh

'Create an array with 20 columns and 100 rows
Dim DataArray(100, 20)
Dim r As Integer

' Printer.NewPage
Printer.FontName = "courier new"
Printer.FontBold = True
Printer.FontSize = "12"
Printer.Orientation = 2

Printer.Print Space(1)
Printer.Print Space(40) + "COMMISSION REPORT"
Printer.Print Space(1)

grandtotal = 0
num = 0
For t = 0 To lista_agentes.ListCount - 1
   
   lista_agentes.ListIndex = t
   grandtotal = grandtotal + Val(Format(lbltotal_invoices.Caption, "000000.00"))

   n$ = lblinitials.Caption + " - " + RTrim(lblfull_name.Caption)
   
   
   ' detecta si es manager y lo marca
   existe = 0
   For Y = 0 To lista_managers.ListCount - 1
      n2$ = UCase(RTrim(Left(lista_managers.List(Y), 20)))
      If n2$ = agente$ Then
          puesto$ = Right(lista_managers.List(Y), Len(lista_managers.List(Y)) - 20)
          If LTrim(UCase(RTrim(puesto$))) = "MANAGER" Then
               existe = 1
               Exit For
          End If
          
          If LTrim(UCase(RTrim(puesto$))) = "MONTERREY" Then
               existe = 1
               Exit For
          End If
          
          If LTrim(UCase(RTrim(puesto$))) = "COMMERCIAL" Then
               existe = 1
               Exit For
          End If
          
      End If
   Next Y



   If existe = 0 Then
    num = num + 1
    DataArray(num, 1) = " "
    DataArray(num, 2) = n$
    DataArray(num, 3) = txtdatefrom.Text
    DataArray(num, 4) = txtdateto.Text
    DataArray(num, 5) = txtdatepayday.Text
    DataArray(num, 6) = "0"
    DataArray(num, 7) = "0"
    DataArray(num, 8) = "$0.00"
    
    If lblnb.Caption = "" Then lblnb.Caption = "0"
    DataArray(num, 9) = lblnb.Caption
    
    DataArray(num, 10) = lblbf.Caption
    DataArray(num, 11) = lblinvoice.Caption
    DataArray(num, 12) = Format(lbldeduction.Caption, "$###,##0.00")
    
    If lblnb_deduc.Caption = "" Then
       lblnb_deduc.Caption = "0"
    End If
    DataArray(num, 13) = Format(lblnb_deduc.Caption, "#0")
    
    DataArray(num, 14) = lbltotal_NB.Caption
    DataArray(num, 15) = lbltotal_bf.Caption
    
    If lblporcentaje.Caption = "" Then lblporcentaje.Caption = "-"
    If lblporcentaje.Caption = "$50.00" Then lblporcentaje.Caption = "$50"
    
    DataArray(num, 16) = lblporcentaje.Caption
    
    If lblcommission.Caption = "" Then lblcommission.Caption = "$0.00"
    DataArray(num, 17) = lblcommission.Caption
    DataArray(num, 18) = "$0.00"
    DataArray(num, 19) = lblid.Caption
    DataArray(num, 20) = txtnotes.Text
   
   End If



Next t





DataArray(0, 1) = " "
DataArray(0, 2) = "Employee"
DataArray(0, 3) = "From"
DataArray(0, 4) = "To"
DataArray(0, 5) = "Pay Day"
DataArray(0, 6) = "R. Hour"
DataArray(0, 7) = "OT"
DataArray(0, 8) = "A. Hour"
DataArray(0, 9) = "NB"
DataArray(0, 10) = "BF"
DataArray(0, 11) = "Invoice"
DataArray(0, 12) = "Deduction"
DataArray(0, 13) = "NB Deduc."
DataArray(0, 14) = "Total NB"
DataArray(0, 15) = "Total BF"
DataArray(0, 16) = "%"
DataArray(0, 17) = "Comission"
DataArray(0, 18) = "Total"
DataArray(0, 19) = "ID"
DataArray(0, 20) = "Notes"


Printer.FontSize = "7"

Printer.Print " Employee                   From       To         Pay-Day    R.Hour   OT  A.Hour   NB      BF      Invoice Deduction NBded TotNB  TotBF      %  Commision     Total    ID   "
Printer.Print Space(1)

For t = 1 To num
 'For Y = 1 To 19
 Printer.Print Format(DataArray(t, 1), "@");
 Printer.Print Format(Left(DataArray(t, 2), 25), "!@@@@@@@@@@@@@@@@@@@@@@@@@@") + Space(1); ' employee
 If DataArray(t, 3) = "" Then DataArray(t, 3) = "  /  /    "
  Printer.Print Format(DataArray(t, 3), "!@@@@@@@@@@") + Space(1); ' from
 If DataArray(t, 4) = "" Then DataArray(t, 4) = "  /  /    "
 Printer.Print Format(DataArray(t, 4), "!@@@@@@@@@@") + Space(1); ' to
 If DataArray(t, 5) = "" Then DataArray(t, 5) = "  /  /    "
 Printer.Print Format(DataArray(t, 5), "!@@@@@@@@@@") + Space(1); 'pay day
 
 Printer.Print Format(DataArray(t, 6), "@@@@@") + Space(1);    'R. Hour
 Printer.Print Format(DataArray(t, 7), "@@@@@") + Space(1);    ' OT
 Printer.Print Format(DataArray(t, 8), "@@@@@@@") + Space(1);    ' A. Hour
 Printer.Print Format(DataArray(t, 9), "@@@@@") + Space(2);      ' NB
 Printer.Print Format(DataArray(t, 10), "@@@@@@@@@") + Space(2); ' BF
 Printer.Print Format(DataArray(t, 11), "@@@@@@@@") + Space(1);  ' Invoice
 Printer.Print Format(DataArray(t, 12), "@@@@@@@@") + Space(1); ' Deduction
 Printer.Print Format(LTrim(RTrim(DataArray(t, 13))), "@@@@") + Space(1); ' NB Deduc
 Printer.Print Format(DataArray(t, 14), "@@@@@") + Space(2); ' Total NB
 Printer.Print Format(DataArray(t, 15), "@@@@@@@@@") + Space(1); ' Total BF
 Printer.Print Format(DataArray(t, 16), "@@@@") + Space(2); ' %
 
 If DataArray(t, 16) = "$50.00" Then
   DataArray(t, 16) = "$50"
 End If
 
 Printer.Print Format(DataArray(t, 17), "@@@@@@@@") + Space(2); ' Commision
 Printer.Print Format(DataArray(t, 18), "@@@@@@@@") + Space(3); ' Total
 Printer.Print Format(DataArray(t, 19), "@@@@") + Space(1); 'ID
 
 
 
 Printer.Print " "
 Printer.Print " "
 
 Printer.FontBold = Not Printer.FontBold
 
Next t



' imprime los comentarios
 
 Printer.FontBold = True

 Printer.Print " "
 Printer.Print " "
  Printer.Print "       N O T E S:"
 Printer.Print " "

For t = 1 To num
 
 Printer.Print Format(DataArray(t, 1), "@");
 r1$ = Left(DataArray(t, 2), 30)
 ' pone ---
 
 a$ = r1$
 For Y = 1 To (30 - Len(a$))
   a$ = a$ + "-"
 Next Y
 
 
 Printer.Print Format(a$, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") + "- "; ' employee
 If DataArray(t, 20) = "" Then
    Printer.Print Left(DataArray(t, 20), 150)   ' note
 Else
     Printer.Print Left(DataArray(t, 20), 150);   ' note
 End If
 Printer.FontBold = Not Printer.FontBold

  
Next t

Printer.Print " "
Printer.Print " "
Printer.Print " "
Printer.Print " "
Printer.FontName = "Arial"
Printer.FontSize = 8
Printer.Print Format(Now, "mm/dd/yyyy") + " " + Format(Now, "hh:mm:ss am/pm")

Printer.EndDoc

msg.Visible = False
msg.Refresh
MsgBox "The report has been printed", 64, "Attention"

End Sub

Private Sub btnreportetotal_Click()
On Error Resume Next

    If txtresultado.Text <> "" Then
       Picture2.Visible = True

       
       Exit Sub
    End If



        
    
    If lista_agentes.ListCount = 0 Then Exit Sub
    
    
    If txtdatefrom.Text = "" Then
       MsgBox "You have not captured the start date", 64, "Attention"
       Exit Sub
    End If
     
    If txtdateto.Text = "" Then
       MsgBox "You have not captured the end date", 64, "Attention"
       Exit Sub
    End If
     
    If txtdatepayday.Text = "" Then
       MsgBox "You have not captured the payday date", 64, "Attention"
       Exit Sub
    End If
     
     
     
    Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

lblmsg.Caption = "Please, wait a moment... Generating the report"

msg.Visible = True
msg.Refresh

 
' GoTo salta
 
txtresultado.Text = ""
For t = 0 To lista_agentes.ListCount - 1
    
    lista_agentes.ListIndex = t
    openforms = DoEvents
   
   empleado_lae$ = lbllae.Caption
   
   
   sSelect = "select ofc.IdOffice, emp.IdEmployee from EmployeeInfo emp " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where emp.Active=1 and empofc.active=1 and IdJobTitle in (16,17,28,2,29,24) and emp.IdEmployee='" + empleado_lae$ + "' and empjob.Active='1'"

   Rs.Open sSelect, base, adOpenUnspecified
   id_oficina$ = Rs(0)
   id_empleado$ = Rs(1)
    Rs.Close
   
  
    
    n$ = ""
    sSelect = "select firstname, Lastname1 from employeeinfo where idemployee='" + id_empleado$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    nombre$ = Rs(0)
    apellido$ = Rs(1)
    Rs.Close
    
    n$ = nombre$ + " " + apellido$
    
    
      txtresultado.Text = txtresultado.Text + "ID: " + Format(empleado_lae$, "0000") + "  " + Format(n$, "!@@@@@@@@@@@@@@@@@@@@") + "  NB: " + Format(lblnb.Caption, "00.0") + "  Total: " + Format(Format(lbltotal_bf.Caption, "$###,##0.00"), "@@@@@@@@@@@") + Chr$(13) + Chr$(10)
      
    
    
      
    msg.Refresh

Next t
     
     
     
   Picture2.Visible = True

   msg.Visible = False
End Sub

Private Sub btnsave_Click()
On Error Resume Next


    If user_sistema$ = "ONLY READ" Then
       MsgBox "This function is only for Administrator", 16, "Access denied"
       Exit Sub
    
    End If
    
    
    If lista_agentes.ListCount = 0 Then Exit Sub
    
    X$ = MsgBox("Do you want to save the work?", 4, "Attention")
    If X$ = "7" Then Exit Sub
    
    If txtdatefrom.Text = "" Then
       MsgBox "You have not captured the start date", 64, "Attention"
       Exit Sub
    End If
     
    If txtdateto.Text = "" Then
       MsgBox "You have not captured the end date", 64, "Attention"
       Exit Sub
    End If
     
    If txtdatepayday.Text = "" Then
       MsgBox "You have not captured the payday date", 64, "Attention"
       Exit Sub
    End If
     
     
     
    Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

lblmsg.Caption = "Please, wait a moment... saving all the information"

msg.Visible = True
msg.Refresh

 
' GoTo salta
 
 
For t = 0 To lista_agentes.ListCount - 1
    
    lista_agentes.ListIndex = t
     openforms = DoEvents
    
   ' GoTo label
   
   empleado_lae$ = lbllae.Caption
   
   
       sSelect = "select ofc.IdOffice, emp.IdEmployee from EmployeeInfo emp " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where emp.Active=1 and empofc.active=1 and IdJobTitle in (16,17,28,2,29,24) and emp.IdEmployee='" + empleado_lae$ + "' and empjob.Active='1'"

   Rs.Open sSelect, base, adOpenUnspecified
   id_oficina$ = Rs(0)
   id_empleado$ = Rs(1)
    Rs.Close
   


 ' verifica si ya existe el registro en SQL
    a$ = ""
    sSelect = "select comm from employeePayee where idemployee='" + id_empleado$ + "' and payday=convert(datetime,'" + txtdatepayday.Text + "') and datefrom= convert(datetime,'" + txtdatefrom.Text + "') and dateto= convert(datetime,'" + txtdateto.Text + "')"
    Rs.Open sSelect, base, adOpenUnspecified
    a$ = Rs(0)
    Rs.Close
    
    
    
   
   
        
    ' If lbllae.Caption = "161" Then Stop
    
     
       If Right(lblporcentaje.Caption, 1) = "%" Then
      var_por = Val(Left(lblporcentaje.Caption, Len(lblporcentaje.Caption) - 1))
    Else
      var_por = 0
    End If
     
   
     
     
  '  If a$ = "" Then
     
  '    sSelect = "INSERT INTO payrollgoals (idemployee, fromdate, todate, payday, rhour, ot, ahour, nb, bf, invoice, deduction, nbdeduction, totalnb, totalbf, percentage, commision, total, notes)  VALUES ('" + _
  '    Format(lbllae.Caption, "####0") + "', convert(datetime,'" + txtdatefrom.Text + "'), convert(datetime,'" + txtdateto.Text + "'), convert(datetime,'" + txtdatepayday.Text + "'), '" + "0" + "', '" + "0" + "', '" + "0.00" + _
  '    "', '" + lblnb.Caption + "', convert(money,'" + lblbf.Caption + "'), convert(money,'" + lblinvoice.Caption + "'), convert(money,'" + lbldeduction.Caption + "'), '" + lblnb_deduc.Caption + "', '" + lbltotal_NB.Caption + _
  '    "', convert(money,'" + lbltotal_bf.Caption + "'), '" + Format(var_por, "##0") + "', convert(money,'" + lblcommission.Caption + "'), convert(money,'" + "$0.00" + "'), '" + txtnotes.Text + "')"
    
  '  Else
    
  '    sSelect = "update payrollgoals set idemployee='" + Format(lbllae.Caption, "####0") + "', fromdate= convert(datetime,'" + txtdatefrom.Text + "')" + _
  '    ", todate= convert(datetime,'" + txtdateto.Text + "'), payday=convert(datetime,'" + txtdatepayday.Text + "'), rhour='0', ot='0', ahour='0.00', " + _
  '    "nb='" + lblnb.Caption + "', bf=convert(money,'" + lblbf.Caption + "'), invoice='" + lblinvoice.Caption + "', deduction=convert(money,'" + lbldeduction.Caption + "')" + _
  '    ", nbdeduction='" + lblnb_deduc.Caption + "', totalnb='" + lbltotal_NB.Caption + "', totalbf=convert(money,'" + lbltotal_bf.Caption + "'), percentage='" + Format(var_por, "##0") + _
  '    "', commision=convert(money,'" + lblcommission.Caption + "'), total=convert(money,'" + "$0.00" + "'), notes='" + txtnotes.Text + "' " + _
  '    "where idemployee='" + lbllae.Caption + "' and payday='" + txtdatepayday.Text + "' and fromdate='" + txtdatefrom.Text + "' and todate='" + txtdateto.Text + "'"
    
  '  End If
        
                      
  '  Rs.Open sSelect, base, adOpenUnspecified
    
  '  Rs.Close
    
    
    If a$ = "" Then
    
    
      
              
      sSelect = "INSERT INTO employeePayee (idemployee, idoffice, idpaytype, datefrom, dateto, payday, " & _
      "regularhours, OvertimeHours, doubletimehours, sickhours, vacationhours, HourlyPay, OTpay, Totalhours, TotalhoursAmount, totalOThours, TotalOTamount,  " & _
      "NB, bf, invoice, nbdeduction, bfdeduction, percentage, Penality, bonus, comm, otbonus, totalamountbf, totalamountbyhr, totalpaymentempl, notes )  VALUES (" + _
      "'" + Format(empleado_lae$, "####0") + "', '" + id_oficina$ + "', '1', convert(datetime,'" + txtdatefrom.Text + "'), convert(datetime,'" + txtdateto.Text + "'), " & _
      "convert(datetime,'" + txtdatepayday.Text + "'), '0', '0', '0', '0', " & _
      "'0', convert(money,'" + Pago_hr$ + "'), convert(money,'" + Pago_OT$ + "'), '0', " & _
      "'0', '0', '0', '" + lblnb.Caption + "', '" + lbltotal_bf.Caption + "', convert(money,'" + lblinvoice.Caption + "'), '" + lblnb_deduc.Caption + "'" & _
      ", convert(money, '" + lbldeduction.Caption + "'), '" + Str(var_por) + "', '0', '0', '" + lblcommission.Caption + "'" & _
      ", '0', '" + Str(Val(lbltotal_bf.Caption) + Val(lblinvoice.Caption)) + "', '" + total_de_hrs_REG_OT$ + "', '" + total_pago$ + "', '" + notas$ + "')"
      
      
      
      
      
    Else
      
      
      
      
      sSelect = "update employeePayee set idemployee='" + Format(empleado_lae$, "####0") + "', " & _
      "idoffice='" + id_oficina$ + "', idpaytype='1', datefrom=convert(datetime,'" + txtdatefrom.Text + "'), dateto=convert(datetime,'" + txtdateto.Text + "'), " & _
      "payday=convert(datetime,'" + txtdatepayday.Text + "'), regularhours='0', OvertimeHours='0',  doubletimehours='0', " & _
      "sickhours='0', vacationhours='0', HourlyPay=convert(money,'" + Pago_hr$ + "'), OTpay=convert(money,'" + Pago_OT$ + "'), " & _
      "Totalhours='0', TotalhoursAmount='0', totalOThours='0', TotalOTamount='0', " & _
      "NB='" + lblnb.Caption + "', BF=convert(money,'" + lbltotal_bf.Caption + "'), invoice=convert(money,'" + lblinvoice.Caption + "'), nbdeduction='" + lblnb_deduc.Caption + "', bfdeduction=convert(money, '" + Format(lbldeduction.Caption, "####0.00") + "'), " & _
      "percentage='" + Str(var_por) + "', Penality='0', bonus='0', comm='" + lblcommission.Caption + "', otbonus='0', totalamountbf=convert(money,'" + Str(Val(lbltotal_bf.Caption) + Val(lblinvoice.Caption)) + "')," & _
      "totalamountbyhr=convert(money,'" + total_de_hrs_REG_OT$ + "'), totalpaymentempl=convert(money,'" + total_pago$ + "'), notes='" + notas$ + "' " & _
      "where idemployee='" + Format(empleado_lae$, "####0") + "' and payday='" + txtdatepayday.Text + "' and datefrom='" + txtdatefrom.Text + "' and dateto='" + txtdateto.Text + "'"
    
      
    End If
        
                      
    Rs.Open sSelect, base, adOpenUnspecified
    
    Rs.Close
    
    
    
label:
    msg.Refresh

Next t
     
     
     
     
salta:
     


   For t = 1 To grid6.Rows - 1
  
  
    grid6.row = t
  
   
    grid6.col = 28
    empleado_lae$ = grid6.Text
    
    'If Val(empleado_lae$) = 165 Then Stop
    
    
     sSelect = "select ofc.IdOffice, emp.IdEmployee from EmployeeInfo emp " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where emp.Active=1 and empofc.active=1 and IdJobTitle in (16,17,28,2,29,24) and emp.IdEmployee='" + empleado_lae$ + "' and empjob.Active='1'"

   Rs.Open sSelect, base, adOpenUnspecified
   id_oficina$ = Rs(0)
   id_empleado$ = Rs(1)
    Rs.Close
   


 ' verifica si ya existe el registro en SQL
    a$ = ""
    sSelect = "select comm from employeePayee where idemployee='" + id_empleado$ + "' and payday=convert(datetime,'" + txtdatepayday.Text + "') and datefrom= convert(datetime,'" + txtdatefrom.Text + "') and dateto= convert(datetime,'" + txtdateto.Text + "')"
    Rs.Open sSelect, base, adOpenUnspecified
    a$ = Rs(0)
    Rs.Close
    
    
  
  
    grid6.col = 7
     regular_HR$ = Format(grid6.Text, "####0.00")
    
    grid6.col = 8
    OT_hours$ = Format(grid6.Text, "####0.00")
    
    grid6.col = 9
    Double_time$ = Format(grid6.Text, "####0.00")
    
    grid6.col = 10
    sick_hours$ = Format(grid6.Text, "####0.00")
    
    If sick_hours$ = "-" Then
         sick_hours$ = "0"
    End If
    
    
    grid6.col = 11
    Vacations$ = Format(grid6.Text, "####0.00")
    
    If Vacations$ = "-" Then
         Vacations$ = "0"
    End If
    
    
    grid6.col = 12
    Pago_hr$ = Format(grid6.Text, "####0.00")
    
    grid6.col = 13
    Pago_OT$ = Format(grid6.Text, "####0.00")
    
    grid6.col = 19
    Invoice$ = Format(grid6.Text, "####0.00")
    
    grid6.col = 20
    Deduccion$ = Format(grid6.Text, "####0.00")
    
    grid6.col = 21
     NB_DED$ = Format(grid6.Text, "####0.00")
    
    
    grid6.col = 22
    bf$ = Format(grid6.Text, "####0.00")
    
    grid6.col = 23
    nb$ = Format(grid6.Text, "####0.00")
    
    grid6.col = 24
    percentage$ = Format(grid6.Text, "####0.00")
    
    grid6.col = 25
    comm$ = Format(grid6.Text, "####0.00")
    
    
    grid6.col = 26
    totalbf$ = Format(grid6.Text, "####0.00")
    
    grid6.col = 27
    totalnb$ = Format(grid6.Text, "####0.00")
  
    grid6.col = 29
    notas$ = grid6.Text
    
  
      total_de_hrs$ = Str(Val(regular_HR$) + Val(sick_hours$) + Val(Vacations$))
      total_hours_amount$ = Str(Val(total_de_hrs$) * Val(Pago_hr$))
      
      total_OT_amount$ = Str(Val(OT_hours$) * Val(Pago_OT$))
      
      total_de_hrs_REG_OT$ = Str(Val(total_OT_amount$) + Val(total_hours_amount$))     ' en pesos
      
      
      total_pago$ = Str(Val(total_hours_amount$) + Val(total_OT_amount$) + Val(comm$))
      
  
  
   If a$ = "" Then
   
      
      
     
     
      sSelect = "INSERT INTO employeePayee (idemployee, idoffice, idpaytype, datefrom, dateto, payday, " & _
      "regularhours, OvertimeHours, doubletimehours, sickhours, vacationhours, HourlyPay, OTpay, Totalhours, TotalhoursAmount, totalOThours, TotalOTamount,  " & _
      "NB, bf, invoice, nbdeduction, bfdeduction, percentage, Penality, bonus, comm, otbonus, totalamountbf, totalamountbyhr, totalpaymentempl, notes )  VALUES (" + _
      "'" + Format(empleado_lae$, "####0") + "', '" + id_oficina$ + "', '1', convert(datetime,'" + txtdatefrom.Text + "'), convert(datetime,'" + txtdateto.Text + "'), " & _
      "convert(datetime,'" + txtdatepayday.Text + "'), '" + regular_HR$ + "', '" + OT_hours$ + "', '" + Double_time$ + "', '" + sick_hours$ + "', " & _
      "'" + Vacations$ + "', convert(money,'" + Pago_hr$ + "'), convert(money,'" + Pago_OT$ + "'), '" + total_de_hrs$ + "', " & _
      "'" + total_hours_amount$ + "', '" + OT_hours$ + "', '" + total_OT_amount$ + "', '" + nb$ + "', '" + bf$ + "', convert(money,'" + Invoice$ + "'), '" & _
      Format(NB_DED$, "####0.00") + "', convert(money, '" + Format(Deduccion$, "####0.00") + "'), '" + percentage$ + "', '0', '0', '" & _
      comm$ + "', '0', '" + Str(Val(bf$) + Val(Invoice$)) + "', '" + total_de_hrs_REG_OT$ + "', '" + total_pago$ + "', '" + notas$ + "')"
      
      
      
    Else
    
    
      
      sSelect = "update employeePayee set idemployee='" + Format(empleado_lae$, "####0") + "', " & _
      "idoffice='" + id_oficina$ + "', idpaytype='1', datefrom=convert(datetime,'" + txtdatefrom.Text + "'), dateto=convert(datetime,'" + txtdateto.Text + "'), " & _
      "payday=convert(datetime,'" + txtdatepayday.Text + "'), regularhours='" + regular_HR$ + "', OvertimeHours='" + OT_hours$ + "',  doubletimehours='" + Double_time$ + "', " & _
      "sickhours='" + sick_hours$ + "', vacationhours='" + Vacations$ + "', HourlyPay=convert(money,'" + Pago_hr$ + "'), OTpay=convert(money,'" + Pago_OT$ + "'), " & _
      "Totalhours='" + Format(Val(total_de_hrs$), "####0.00") + "', TotalhoursAmount='" + Format(total_hours_amount$, "00000.00") + "', totalOThours='" + OT_hours$ + "', TotalOTamount='" + total_OT_amount$ + "', " & _
      "NB='" + nb$ + "', BF=convert(money,'" + bf$ + "'), invoice=convert(money,'" + Invoice$ + "'), nbdeduction='" + Format(NB_DED$, "####0.00") + "', bfdeduction=convert(money, '" + Format(Deduccion$, "####0.00") + "'), " & _
      "percentage='" + percentage$ + "', Penality='0', bonus='0', comm='" + comm$ + "', otbonus='0', totalamountbf=convert(money,'" + Str(Val(bf$) + Val(Invoice$)) + "')," & _
      "totalamountbyhr=convert(money,'" + total_de_hrs_REG_OT$ + "'), totalpaymentempl=convert(money,'" + total_pago$ + "'), notes='" + notas$ + "' " & _
      "where idemployee='" + Format(empleado_lae$, "####0") + "' and payday='" + txtdatepayday.Text + "' and datefrom='" + txtdatefrom.Text + "' and dateto='" + txtdateto.Text + "'"
    
      
      
      
      
         
    End If
        
                      
    Rs.Open sSelect, base, adOpenUnspecified
    
    Rs.Close
  
  Next t
     
     
     
   msg.Visible = False
     
    
End Sub

Private Sub btnupdate_Click()
On Error Resume Next
If txtdatefrom.Text = "" Then
  MsgBox "The start date is invalid", 64, "Attention"
  Exit Sub
End If

If txtdateto.Text = "" Then
  MsgBox "The end date is invalid", 64, "Attention"
  Exit Sub
End If

If txtdatepayday.Text = "" Then
  MsgBox "The pay date is invalid", 64, "Attention"
  Exit Sub
End If



lblmsg.Caption = "Please, wait a moment... Updating all data"
msg.Visible = True
msg.Refresh
lblexcel.Visible = False
Shape10.Visible = False


btn_errors_Click
btn_UW_Click
btn_dmv_Click

calcula_grandtotal

lblexcel.Visible = True
Shape10.Visible = True
btncargar_excel.Visible = True
msg.Visible = False
msg.Refresh

lblregistros.Caption = Format(lista_agentes.ListCount, "##0")

End Sub

Private Sub Calendar1_Click()
On Error Resume Next

Select Case calen
Case 0
  txtdatefrom.Text = Calendar1.Value
Case 1
  txtdateto.Text = Calendar1.Value
Case 2
  txtdatepayday.Text = Calendar1.Value
End Select
Calendar1.Visible = False


End Sub

Private Sub cboimpre_Click()
On Error Resume Next


For Each xprint In Printers
           If xprint.DeviceName = cboimpre.Text Then
              ' La define como predeterminada del sistema.
              Set Printer = xprint
              DoEvents
              Exit For
           End If
Next


nf = FreeFile
 Open "c:\goals\printer" For Output Shared As #nf
 Lock #nf
 Print #nf, Printer.DeviceName
 Print #nf, Printer.Port
 Unlock #nf
 Close #nf
 Text1.SetFocus

End Sub


' Add a reference to "Microsoft Excel 11.0 Object Library"
' (or whatever version you have installed on your system).

Private Sub Form_Load()
On Error Resume Next
Dim file_name As String


If (App.PrevInstance = True) Then
  'base.Close
  End
End If

txtresultado.Text = ""

' convierte_tabla

expiracion_invoices = 0

Conecta_SQL


  

carga_impresoras

carga_tiers


   ' ******   estas lineas leen el archivo de excel
   
   ' file_name = Application.StartupPath
   ' If Right$(file_name, 1) <> "\" Then file_name = file_name & "\"
   ' file_name = file_name & "Items.xls"
   ' txtFile.Text = GetSetting("howto_read_excel", "Settings", "File", file_name)
    
    
   columna = 0
   Top = 0
   Left = (Screen.Width - Width) / 2
   letra$ = "-"
   
   
   
   Dim ScaleFactorX As Single, ScaleFactorY As Single  ' Scaling factors
      ' Size of Form in Pixels at design resolution
      
      'If Screen.Width <= 12000 Then
         ' DesignX =  800
      'Else
          DesignX = 1366 '1024
      'End If
      
      'If Screen.Height <= 9000 Then
      '      DesignY = 600  '800
      'Else
            DesignY = 768 '940 '1024
      'End If
      
      
      RePosForm = True   ' Flag for positioning Form
      DoResize = False   ' Flag for Resize Event
      ' Set up the screen values
      Xtwips = Screen.TwipsPerPixelX
      Ytwips = Screen.TwipsPerPixelY
      Ypixels = Screen.Height / Ytwips ' Y Pixel Resolution
      Xpixels = Screen.Width / Xtwips  ' X Pixel Resolution

      ' Determine scaling factors
      If DesignX = 800 Then
        ScaleFactorX = (Xpixels / DesignX)  ' 0.78
        ScaleFactorY = (Ypixels / DesignY)  ' 0.78
      Else
        ScaleFactorX = (Xpixels / DesignX)
        ScaleFactorY = (Ypixels * 1.3 / DesignY)
      
        'ScaleFactorX = 1360 / DesignX
        'ScaleFactorY = 1024 / DesignY
      End If
      
      ScaleMode = 1  ' twips
      'Exit Sub  ' uncomment to see how Form1 looks without resizing
      Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
      'Label1.Caption = "Current resolution is " & Str$(Xpixels) + _
       '"  by " + Str$(Ypixels)
      If DesignX = 800 Then
        forma_main.Height = 9000 'Me.Height ' Remember the current size
        forma_main.Width = 12000 'Me.Width
      Else
        Height = Me.Height ' Remember the current size
        Width = Me.Width
      
      End If
primeravez = 0

   

    file_name = "c:\goals" 'App.Path
    
   ' nf = FreeFile
   ' Open file_name + "\test.txt" For Output Shared As #nf
   ' Print #nf, "Nombres"
   ' Print #nf, Chr$(9) + "hector"
   ' Print #nf, Chr$(9) + "Elena"
   ' Print #nf, Chr$(9) + "Carlos"
   ' Print #nf, "Apellidos"
   ' Print #nf, Chr$(9) + "Navarro"
   ' Print #nf, Chr$(9) + "Vazquez"
    
   ' Print #nf, "Condados"
   ' Print #nf, Chr$(9) + "San Bernardino"
   ' Print #nf, Chr$(9) + Chr$(9) + "Ontario"
   ' Print #nf, Chr$(9) + Chr$(9) + "Fontana"
   ' Print #nf, Chr$(9) + Chr$(9) + "Upland"
   ' Print #nf, Chr$(9) + Chr$(9) + "Rancho Cucamonga"
    
   ' Print #nf, Chr$(9) + "Los angeles"
   ' Print #nf, Chr$(9) + Chr$(9) + "Whittier"
   ' Print #nf, Chr$(9) + Chr$(9) + "Compton"
   ' Print #nf, Chr$(9) + Chr$(9) + "Echo Park"
   ' Print #nf, Chr$(9) + "Riverside"
   ' Print #nf, Chr$(9) + Chr$(9) + "Moreno Valley"
    
   ' Close nf
    
     With TreeView1
        .Style = tvwTreelinesPlusMinusText
        .LineStyle = tvwRootLines
        .PathSeparator = "\"
        '.Indentation = Screen.TwipsPerPixelX * 5 '256
        '
        ' No permitir la edición automática del texto
        .LabelEdit = tvwManual
        ' Para que se pueda expandir al seleccionar un nodo,
        ' cambia este valor a True,
        ' si se deja en False, tendrás que hacer doble-click
        .SingleSel = True
        ' Para que al perder el foco,
        ' se siga viendo el que está seleccionado
        .HideSelection = False
        '
        .Refresh
    End With
    
    
    
  '  If Right$(file_name, 1) <> "\" Then file_name = file_name & "\"
  '  file_name = file_name & "test.txt"
  '  LoadTreeViewFromFile file_name, TreeView1
   
   
cboexcepcion.Clear
cboexcepcion.AddItem "BF"
cboexcepcion.AddItem "BF 12 Paid In Full"
cboexcepcion.AddItem "BF Call Center"
cboexcepcion.AddItem "BF Commercial"
cboexcepcion.AddItem "BF DMV"
cboexcepcion.AddItem "BF Endo Fee"
cboexcepcion.AddItem "BF Payment Fee"
cboexcepcion.AddItem "BF NSD"
cboexcepcion.AddItem "Invoice"
Checa_status


carga_ubicaciones

carga_attributos

   
End Sub

' Write tabs indicating this node's depth in
' the tree followed by the node's text.
' Then save its children and its siblings.
Private Sub SaveNode(ByVal fnum As Integer, ByVal n As Node, ByVal level As Integer)
    If n Is Nothing Then Exit Sub

    ' Save the node.
    Print #fnum, String$(level, vbTab) & n.Text

    ' Save its children.
    SaveNode fnum, n.Child, level + 1

    ' Save its next sibling.
    SaveNode fnum, n.Next, level
End Sub
' Save a TreeView control into a file that uses tabs
' to show indentation.
Private Sub SaveTreeViewIntoFile(ByVal file_name As String, ByVal trv As TreeView)
Dim fnum As Integer

    fnum = FreeFile
    Open file_name For Output As fnum

    ' Find the root nodes.
    If TreeView1.Nodes.Count > 0 Then SaveNode fnum, TreeView1.Nodes(1), 0

    Close fnum
End Sub

Private Sub Form_Resize()
   On Error Resume Next
Dim ScaleFactorX As Single, ScaleFactorY As Single

If primeravez = 0 Then


primeravez = 1
      If Not DoResize Then  ' To avoid infinite loop
         DoResize = True
         Exit Sub
      End If

      RePosForm = False
      ScaleFactorX = Me.Width / MyForm.Width   ' How much change?
      ScaleFactorY = Me.Height / MyForm.Height
      Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
      MyForm.Height = Me.Height ' Remember the current size
      MyForm.Width = Me.Width
End If
primeravez = 1
End Sub

Private Sub Form_Terminate()
On Error Resume Next
 base.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim file_name As String


    file_name = "c:\goals" 'App.Path
    
    If Right$(file_name, 1) <> "\" Then file_name = file_name & "\"
    file_name = file_name & "test.txt"
    SaveTreeViewIntoFile file_name, TreeView1
End Sub

Private Sub mnuPopupAddNode_Click()
Dim txt As String
Dim new_node As Node

    txt = InputBox("Text", "Add Node", "")
    If Len(txt) > 0 Then
        If TreeView1.SelectedItem Is Nothing Then
            Set new_node = TreeView1.Nodes.Add(, , , txt)
        Else
            Set new_node = TreeView1.Nodes.Add( _
                TreeView1.SelectedItem, tvwChild, , txt)
        End If
        new_node.EnsureVisible
    End If
End Sub





Private Sub mnuPopupDeleteNode_Click()
    TreeView1.Nodes.Remove TreeView1.SelectedItem.Index
End Sub






' Set a title Label and the values in a ListBox. Get the title from cell (row, col).
' Get the values from cell (row + 1, col) to the end of the column.
Private Sub SetTitleAndListValues(ByVal sheet As Excel.Worksheet, _
    ByVal row As Integer, ByVal col As Integer, ByVal lbl As label, ByVal lst As ListBox)
    
On Error Resume Next


Dim range As Excel.range
Dim last_cell As Excel.range
Dim first_cell As Excel.range
Dim value_range As Excel.range
    
    
Dim range_values() As Variant
Dim num_items As Integer
Dim i As Integer

    ' Set the title.
    Set range = sheet.Cells(row, col)
    lbl.Caption = CStr(range.Value2)
    lbl.ForeColor = range.Font.Color
    lbl.BackColor = range.Interior.Color

    ' Get the values.
    ' Find the last cell in the column.
    Set range = sheet.Columns(col)
    
    
     a = ActiveSheet.Columns("D").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
   
    Set last_cell = range.End(xlDown)
    
   
    total_filas_max = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
    ' total_filas_max = ActiveSheet.Columns("A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
   
    
    ' Get a Range holding the values.
    Set first_cell = sheet.Cells(row + 1, col)
    Set value_range = sheet.range(first_cell, last_cell)

    ' Get the values.
    range_values = value_range.Value



    ' Convert this into a 1-dimensional array.
    ' Note that the Range's array has lower bounds 1.
    'num_items = UBound(range_values, 1)
    num_items = UltimaFila - 1
    
    For i = 1 To num_items
        lst.AddItem range_values(i, 1)
    Next i
    
    ' poner los valores en una hoja de excel en el programa
    num_items = UBound(range_values, 1)
    
    
    num_items = UltimaFila - 1
    
    filas = num_items + 1
    If filas >= Grid1.Rows Then
       Grid1.Rows = filas
    End If
       
    
    
    
        
    Grid1.col = columna
    For i = 1 To num_items
        Grid1.row = i
        Grid1.Text = range_values(i, 1)
        
    Next i
    
    columna = columna + 1
    
    
   
    
End Sub


Private Sub SetTitleAndListValues1(ByVal sheet As Excel.Worksheet, _
    ByVal row As Integer, ByVal col As Integer, ByVal lbl As label, ByVal lst As ListBox)
    
On Error Resume Next


Dim range As Excel.range
Dim last_cell As Excel.range
Dim first_cell As Excel.range
Dim value_range As Excel.range
    
    
Dim range_values() As Variant
Dim num_items As Integer
Dim i As Integer

    ' Set the title.
    Set range = sheet.Cells(row, col)
    lbl.Caption = CStr(range.Value2)
    lbl.ForeColor = range.Font.Color
    lbl.BackColor = range.Interior.Color

    ' Get the values.
    ' Find the last cell in the column.
    Set range = sheet.Columns(col)
    
    
     a = ActiveSheet.Columns("D").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
   
    Set last_cell = range.End(xlDown)
    
   
    total_filas_max = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
    ' total_filas_max = ActiveSheet.Columns("A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
   
    
    ' Get a Range holding the values.
    Set first_cell = sheet.Cells(row + 1, col)
    Set value_range = sheet.range(first_cell, last_cell)

    ' Get the values.
    range_values = value_range.Value



    ' Convert this into a 1-dimensional array.
    ' Note that the Range's array has lower bounds 1.
    'num_items = UBound(range_values, 1)
    num_items = UltimaFila - 1
    
    For i = 1 To num_items
        lst.AddItem range_values(i, 1)
    Next i
    
    ' poner los valores en una hoja de excel en el programa
    num_items = UBound(range_values, 1)
    
    
    num_items = UltimaFila - 1
    
    filas = num_items + 1
    If filas >= Grid1.Rows Then
       Grid1.Rows = filas
    End If
       
    
    
    
        
    Grid1.col = columna
    For i = 1 To num_items
        Grid1.row = i
        Grid1.Text = range_values(i, 1)
        
    Next i
    
    columna = columna + 1
    
    
   
    
End Sub


Private Sub SetTitleAndListValues2(ByVal sheet As Excel.Worksheet, _
    ByVal row As Integer, ByVal col As Integer, ByVal lbl As label, ByVal lst As ListBox)
    
On Error Resume Next


Dim range As Excel.range
Dim last_cell As Excel.range
Dim first_cell As Excel.range
Dim value_range As Excel.range
    
    
Dim range_values() As Variant
Dim num_items As Integer
Dim i As Integer

    ' Set the title.
    Set range = sheet.Cells(row, col)
    lbl.Caption = CStr(range.Value2)
    lbl.ForeColor = range.Font.Color
    lbl.BackColor = range.Interior.Color

    ' Get the values.
    ' Find the last cell in the column.
    Set range = sheet.Columns(col)
    
    
     a = ActiveSheet.Columns("D").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
   
    Set last_cell = range.End(xlDown)
    
   
    total_filas_max = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).row
    ' total_filas_max = ActiveSheet.Columns("A").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
   
    
    ' Get a Range holding the values.
    Set first_cell = sheet.Cells(row + 1, col)
    Set value_range = sheet.range(first_cell, last_cell)

    ' Get the values.
    range_values = value_range.Value



    ' Convert this into a 1-dimensional array.
    ' Note that the Range's array has lower bounds 1.
    'num_items = UBound(range_values, 1)
    num_items = UltimaFila - 1
    
    For i = 1 To num_items
        lst.AddItem range_values(i, 1)
    Next i
    
    ' poner los valores en una hoja de excel en el programa
    num_items = UBound(range_values, 1)
    
    
    num_items = UltimaFila - 1
    
    filas = num_items + 1
    If filas >= Grid2.Rows Then
       Grid2.Rows = filas
    End If
       
    
    
    
        
    Grid2.col = columna
    For i = 1 To num_items
        Grid2.row = i
        Grid2.Text = range_values(i, 1)
        
    Next i
    
    columna = columna + 1
    
    
   
    
End Sub



Public Sub separa_campos_NB()
On Error Resume Next


grid.Clear
grid.cols = 8
grid.Rows = Grid1.Rows

Grid1.row = 1
grid.row = 0

   Grid1.col = 3   ' RECEIPT
   grid.col = 1
   grid.Text = Grid1.Text

   Grid1.col = 4  ' DATE
   grid.col = 2
   grid.Text = Grid1.Text
   
   Grid1.col = 7  ' INVOICE ITEM
   grid.col = 3
   grid.Text = Grid1.Text
   
   Grid1.col = 13  ' user
   grid.col = 4
   grid.Text = Grid1.Text
   
   Grid1.col = 14  ' CSR
   grid.col = 5
   grid.Text = Grid1.Text
      
   Grid1.col = 20  ' Amount
   grid.col = 6
   grid.Text = Grid1.Text
   
   Grid1.col = 16  ' office
   grid.col = 7
   Grid1.col = Grid1.Text


For t = 2 To Grid1.Rows - 1
   
   
   Grid1.row = t
   grid.row = t
   
   Grid1.col = 0
   grid.col = 0
   grid.Text = Grid1.Text
   
      
   
   Grid1.col = 3   ' RECEIPT
   grid.col = 1
   grid.Text = Grid1.Text
   
   Grid1.col = 4  ' DATE
   grid.col = 2
   grid.Text = Grid1.Text
   
   Grid1.col = 7  ' INVOICE ITEM
   grid.col = 3
   grid.Text = Grid1.Text
   
   Grid1.col = 13  ' user
   grid.col = 4
   grid.Text = Grid1.Text
   
   Grid1.col = 14  ' CSR
   grid.col = 5
   grid.Text = Grid1.Text
   
   
   Grid1.col = 20  ' Amount
   grid.col = 6
   grid.Text = Grid1.Text
   
   Grid1.col = 16  ' office
   grid.col = 7
   grid.Text = Grid1.Text
   
   
Next t


' transfiere datos necesarios solamente

Grid1.Clear
Grid1.cols = 8
Grid1.Rows = grid.Rows - 1

Grid1.row = 0
grid.row = 0
For t = 0 To 7
  Grid1.col = t
  grid.col = t
  Grid1.Text = grid.Text
  
Next t


Erase tabla_NB



For z = 2 To grid.Rows - 1
   Grid1.row = z - 1
   grid.row = z
   
   For Y = 0 To 7
     Grid1.col = Y
     grid.col = Y
     Grid1.Text = grid.Text
     
     tabla_NB(z - 1, Y) = grid.Text
   Next Y
Next z
   
   
   
   
   
setup_grid1





End Sub

Public Sub separa_campos_GI()
On Error Resume Next


grid.Clear
grid.cols = 8
grid.Rows = Grid2.Rows

Grid2.row = 1
grid.row = 0

   Grid2.col = 3   ' RECEIPT
   grid.col = 1
   grid.Text = Grid2.Text

   Grid2.col = 4  ' DATE
   grid.col = 2
   grid.Text = Grid2.Text
   
   Grid2.col = 7  ' INVOICE ITEM
   grid.col = 3
   grid.Text = Grid2.Text
   
   Grid2.col = 13  ' user
   grid.col = 4
   grid.Text = Grid2.Text
   
   Grid2.col = 14  ' CSR
   grid.col = 5
   grid.Text = Grid2.Text
      
   Grid2.col = 20  ' Amount
   grid.col = 6
   grid.Text = Grid2.Text
   
   Grid2.col = 16  ' office
   grid.col = 7
   grid.Text = Grid2.Text
   


For t = 2 To Grid2.Rows - 1
   
   
   Grid2.row = t
   grid.row = t
   
   Grid2.col = 0
   grid.col = 0
   grid.Text = Grid2.Text
   
      
   
   Grid2.col = 3   ' RECEIPT
   grid.col = 1
   grid.Text = Grid2.Text
   
   Grid2.col = 4  ' DATE
   grid.col = 2
   grid.Text = Grid2.Text
   
   Grid2.col = 7  ' INVOICE ITEM
   grid.col = 3
   grid.Text = Grid2.Text
   
   Grid2.col = 13  ' user
   grid.col = 4
   grid.Text = Grid2.Text
   
   Grid2.col = 14  ' CSR
   grid.col = 5
   grid.Text = Grid2.Text
   
   
   Grid2.col = 20  ' Amount
   grid.col = 6
   grid.Text = Grid2.Text
   
   Grid2.col = 16  ' Amount
   grid.col = 7
   grid.Text = Grid2.Text
   
   
Next t


' transfiere datos necesarios solamente

Grid2.Clear
Grid2.cols = 8
Grid2.Rows = grid.Rows - 1

Grid2.row = 0
grid.row = 0
For t = 0 To 7
  Grid2.col = t
  grid.col = t
  Grid2.Text = grid.Text
  
Next t

Erase tabla_GI



For z = 2 To grid.Rows - 1
   Grid2.row = z - 1
   grid.row = z
   
   For Y = 0 To 7
     Grid2.col = Y
     grid.col = Y
     Grid2.Text = grid.Text
     
     tabla_GI(z - 1, Y) = grid.Text
   Next Y
Next z
   
   
   
   
setup_grid2




End Sub

Public Sub CARGA_AGENTES()
On Error Resume Next
Dim sSelect As String
Dim Rs As ADODB.Recordset
    
Set Rs = New ADODB.Recordset


If Lista_empleados.ListCount > 0 Then

  lista_agentes.Clear

  For t = 0 To Lista_empleados.ListCount - 1
     If UCase(Left(Lista_empleados.List(t), 1)) = letra$ Or letra$ = "-" Then
           lista_agentes.AddItem Lista_empleados.List(t)
     End If
  Next t



Else

' carga la lista por primera vez


lista_agentes.Clear
For t = 1 To Grid1.Rows
   Grid1.col = 5
   Grid1.row = t
   agente1$ = UCase(Grid1.Text)
   existe = 0
   For Y = 0 To lista_agentes.ListCount - 1
     If UCase(Left(agente1$, 1)) = letra$ Or letra$ = "-" Then
      If UCase(agente1$) = UCase(lista_agentes.List(Y)) Then
         existe = 1
         Exit For
      End If
     End If
   Next Y
   
   If UCase(Left(agente1$, 1)) = letra$ Or letra$ = "-" Then
     If existe = 0 And agente1$ <> "" Then
       lista_agentes.AddItem agente1$
     End If
   End If

Next t
         
         
         
For t = 1 To Grid1.Rows
   Grid1.col = 4
   Grid1.row = t
   agente1$ = UCase(Grid1.Text)
   existe = 0
   For Y = 0 To lista_agentes.ListCount - 1
     If UCase(Left(agente1$, 1)) = letra$ Or letra$ = "-" Then
      If UCase(agente1$) = UCase(lista_agentes.List(Y)) Then
         existe = 1
         Exit For
      End If
     End If
   Next Y
   
   If UCase(Left(agente1$, 1)) = letra$ Or letra$ = "-" Then
     If existe = 0 And agente1$ <> "" Then
       lista_agentes.AddItem agente1$
     End If
   End If

Next t



' ***************************************************************
'  GRID2

For t = 1 To Grid2.Rows
   Grid2.col = 5
   Grid2.row = t
   agente1$ = UCase(Grid2.Text)
   existe = 0
   For Y = 0 To lista_agentes.ListCount - 1
     If UCase(Left(agente1$, 1)) = letra$ Or letra$ = "-" Then
      If UCase(agente1$) = UCase(lista_agentes.List(Y)) Then
         existe = 1
         Exit For
      End If
     End If
   Next Y
   
   If UCase(Left(agente1$, 1)) = letra$ Or letra$ = "-" Then
     If existe = 0 And agente1$ <> "" Then
       lista_agentes.AddItem agente1$
     End If
   End If

Next t
         
         
         
For t = 1 To Grid2.Rows
   Grid2.col = 4
   Grid2.row = t
   agente1$ = UCase(Grid2.Text)
   existe = 0
   For Y = 0 To lista_agentes.ListCount - 1
     If UCase(Left(agente1$, 1)) = letra$ Or letra$ = "-" Then
      If UCase(agente1$) = UCase(lista_agentes.List(Y)) Then
         existe = 1
         Exit For
      End If
     End If
   Next Y
   
   If UCase(Left(agente1$, 1)) = letra$ Or letra$ = "-" Then
     If existe = 0 And agente1$ <> "" Then
       lista_agentes.AddItem agente1$
     End If
   End If

Next t




For t = 0 To lista_agentes.ListCount - 1

   r$ = lista_agentes.List(t)

 ' verifica si existe
    sSelect = "Select idemployee from employeeinfo where username='" + r$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    id_emp$ = Rs(0)
                         
    Rs.Close
    
    
    id_empLAE$ = ""
    ' verifica si existe
    sSelect = "Select initials from payrollconfig where idemployeeLAE='" + id_emp$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    id_empLAE$ = Rs(0)
                         
    Rs.Close
    
    
     
    If id_empLAE$ = "" Then
       ' MsgBox "The user named " + UCase(r$) + " does not exist in the Payroll Config. Please, add it!", 64, "Attention"
      
    End If
    
Next t

Lista_empleados.Clear
For Y = 0 To lista_agentes.ListCount - 1
  Lista_empleados.AddItem lista_agentes.List(Y)
Next Y


End If

End Sub

Private Sub List3_Click()
On Error Resume Next
Dim sSelect As String
Dim Rs As ADODB.Recordset
    
Set Rs = New ADODB.Recordset


lbltotal_unidades.Caption = "0"
If marca = 1 Then
   marca = 0
   Exit Sub
End If
        

folio = Val(Left(List3.List(List3.ListIndex), 6))

encontrado = 0
For t = 1 To Grid2.Rows
  Grid2.row = t
  Grid2.col = 1
  num = Val(Grid2.Text)
  
  If num = folio Then
   
    fila = Grid2.RowSel
    Grid2.row = fila
    
    Grid2.col = 0
    num1$ = Grid2.Text
    
    Grid2.col = 1
    recibo$ = Grid2.Text
    
    Grid2.col = 2
    fecha$ = Grid2.Text
    
    Grid2.col = 3
    Item$ = Grid2.Text
    
    Grid2.col = 4
    user1$ = Grid2.Text
    
    Grid2.col = 5
    csr1$ = Grid2.Text
    
    Grid2.col = 6
    cant$ = Grid2.Text
    
    Grid2.col = 9
    cant_pagada$ = Grid2.Text
    
    Grid2.col = 13
    csr2$ = Grid2.Text
    
   
    Exit For
  End If
  
Next t



  agente_cobrador$ = ""
 
  sSelect = "select idemployeeUSR from [ReceiptsDTL] recdtl " & _
  "inner join  ReceiptsHDR  rechdr on rechdr.IDReceiptHDR =recdtl.IdReceiptHDR " & _
  "inner join InvoiceItemCatalog iitem on iitem.IdInvoiceItem=recdtl.IdInvoiceItem " & _
  "where rechdr.Active=1 and recdtl.[IdReceiptHDR] in (" + recibo$ + ") and iitem.InvoiceItemName='Late Fee'"
 
   Rs.Open sSelect, base, adOpenUnspecified
   agente_cobrador$ = Rs(0)
   Rs.Close
    





  sSelect = "select rechdr.Date, IDReceiptHDR, idemployeeUSR, idemployeeCSR1, balancedue from ReceiptsBalancePayments recbalpay " & _
         "inner join ReceiptsHDR rechdr on recbalpay.IdReceiptsHDRWBalance=rechdr.IDReceiptHDR " & _
         "Where IdReceiptsHDRPayBalance='" + Format(folio, "#000000") + "'"
         
          ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
          Rs.Open sSelect, base, adOpenUnspecified
    
          fecha_de_creacion_recibo$ = Format(Rs(0), "mm/dd/yyyy")
          reciboHDR$ = Rs(1)
          
          user2$ = Rs(2)
          csr2$ = Rs(3)
          
          balance_factura$ = Rs(4)
          
          Rs.Close
          
          
          
          
          sSelect = "select IdReceiptsHDRwBalance from ReceiptsBalancePayments where IdReceiptsHDRPayBalance='" + Format(folio, "#000000") + "'"
          Rs.Open sSelect, base, adOpenUnspecified
          ID_reciboHDRW$ = Rs(0)
          Rs.Close
              
    
    
          sSelect = "select idemployeeUSR, idemployeeCSR1 from ReceiptsHDR where IDReceiptHDR='" + ID_reciboHDRW$ + "'"
          Rs.Open sSelect, base, adOpenUnspecified
          user1$ = Rs(0)
          csr1$ = Rs(1)
          Rs.Close
          
          
                   
          
          'If user2$ <> csr2$ Then
          '  user1$ = user2$
          '  CSR1$ = csr2$
          'End If
          
          
          
          sSelect = "select username from EmployeeInfo where IDEmployee='" + user1$ + "'"
          Rs.Open sSelect, base, adOpenUnspecified
          user1$ = Rs(0)
          Rs.Close
           
           
          sSelect = "select username from EmployeeInfo where IDEmployee='" + csr1$ + "'"
               Rs.Open sSelect, base, adOpenUnspecified
               csr1$ = Rs(0)
               Rs.Close
         
         
          'sSelect = "select balancedue from ReceiptsHDR where IDReceiptHDR='" + reciboHDR$ + "'"
          'Rs.Open sSelect, base, adOpenUnspecified
    
          'balance_factura2$ = Rs(0)
          
          'Rs.Close
          
          
   If UCase(Item$) = "INVOICE" Then
          
      If Val(cant_pagada$) < Val(cant$) Then
           cant$ = cant_pagada$
      End If
       
   End If
   
   
          
If csr2$ <> "" Then
   a$ = "Receipt# " + recibo$ + Chr$(13) + "Date: " + fecha$ + Chr$(13) + "Concept: " + Item$ + Chr$(13) + Chr$(13) + "User: " + user1$ + Chr$(13) + "CSR: " + csr1$ + Chr$(13) + "CSR2: " + csr2$ + Chr$(13) + Chr$(13) + "Amount: " + Format(Val(cant$), "$###,##0.00") + Chr$(13) + Chr$(13) ' + "Amount paid: " + Format(Val(cant_pagada$), "$###,##0.00") + Chr$(13) + Chr$(13)
Else
   a$ = "Receipt# " + recibo$ + Chr$(13) + "Date: " + fecha$ + Chr$(13) + "Concept: " + Item$ + Chr$(13) + Chr$(13) + "User: " + user1$ + Chr$(13) + "CSR: " + csr1$ + Chr$(13) + Chr$(13) + "Amount: " + Format(Val(cant$), "$###,##0.00") + Chr$(13) + Chr$(13) ' + "Amount paid: " + Format(Val(cant_pagada$), "$###,##0.00") + Chr$(13) + Chr$(13)
End If


If agente_cobrador$ <> "" Then
   sSelect = "select username from EmployeeInfo where IDEmployee='" + agente_cobrador$ + "'"
   Rs.Open sSelect, base, adOpenUnspecified
   cobrador$ = Rs(0)
   Rs.Close

   a$ = a$ + "Collection Agent: " + cobrador$ + Chr$(13) + Chr$(13)
End If



If UCase(Item$) = "INVOICE" Then

   a$ = a$ + "----------------------" + Chr$(13) + "Creation date: " + fecha_de_creacion_recibo$

End If


MsgBox a$, 64, "Row# " + num1$

End Sub

Private Sub List3_DblClick()
On Error Resume Next
marca = 2
List3_Click

End Sub

Private Sub List3_ItemCheck(Item As Integer)
On Error Resume Next
marca = 1
End Sub


Private Sub List5_Click()
On Error Resume Next
List5.Selected(List5.ListIndex) = False
Text1.SetFocus
End Sub

Private Sub lista_agentes_Click()
On Error Resume Next


agente$ = lista_agentes.List(lista_agentes.ListIndex)

lblmsg.Caption = "Please, wait a moment... loading all data"
msg.Visible = True
msg.Refresh


'carga_tabla
'Exit Sub
calcula_NB
carga_inv
carga_GI

' ================  AQUI EMPIEZA LA RUTINA ORIGINAL   ==============================================================================================


lbltotal_unidades.Caption = "0"

lblidentificacion.Caption = ""
lbllae.Caption = ""
lblfull_name.Caption = ""
agente$ = lista_agentes.List(lista_agentes.ListIndex)
lblagent.Caption = lista_agentes.List(lista_agentes.ListIndex)

chklock1.Value = False
chklock2.Value = False
chklock3.Value = False
'chklock1.Enabled = False
'chklock2.Enabled = False
'chklock3.Enabled = False




txtnotes.Text = RTrim(notes(lista_agentes.ListIndex))





If chklock3.Value = 1 Then

   lblporcentaje.Caption = ""
Else
   lblporcentaje.Caption = RTrim(porcentaje(lista_agentes.ListIndex))

End If

calcula_CSR_and_USER
'asigna_limite


' detecta si es manager y lo marca
existe = 0
existe_monterrey = 0

For t = 0 To lista_managers.ListCount - 1
  n$ = UCase(RTrim(Left(lista_managers.List(t), 20)))
  If n$ = agente$ Then
    puesto$ = Right(lista_managers.List(t), Len(lista_managers.List(t)) - 20)
    If LTrim(UCase(RTrim(puesto$))) = "MANAGER" Then
          img_gerente.Visible = True
          existe = 1
          Exit For
    End If
  End If
  
  
  
  
Next t



If existe = 0 Then img_gerente.Visible = False
     
  
  
' detecta si es monterrey y lo marca
existe_monterrey = 0
For t = 0 To lista_managers.ListCount - 1
  n$ = UCase(RTrim(Left(lista_managers.List(t), 20)))
  If n$ = agente$ Then
    puesto$ = Right(lista_managers.List(t), Len(lista_managers.List(t)) - 20)
    If LTrim(UCase(RTrim(puesto$))) = "MONTERREY" Then
          
          existe_monterrey = 1
          Exit For
    End If
  End If
Next t
  

  
  
  

' detecta si hay un invoice
suma_invoice = 0
For t = 0 To List1.ListCount - 1
  If RTrim(UCase(Left(List1.List(t), 20))) = UCase(lblagent.Caption) Then
  
   n$ = Right(List1.List(t), Len(List1.List(t)) - 20)
   a$ = Left(n$, 20)
   If LTrim(RTrim(UCase$(a$))) = "INVOICE" Then
     cant = Val(Right(n$, Len(n$) - 20))
     suma_invoice = suma_invoice + cant
   End If
  
  End If
Next t

carga_iniciales



' detecta si hay errores para descontar
' monterrey

multa = 0
existe = 0
For t = 0 To List7.ListCount - 1
  n$ = Left(List7.List(t), 20)
  existe = 0
  If UCase(RTrim(n$)) = UCase(agente$) Then
    existe = 1
    multa = Val(Right(List7.List(t), Len(List7.List(t)) - 20))
    Exit For
  End If
Next t


 ' underwriting
existe2 = 0
For t = 0 To lista_PhoneSales.ListCount - 1
  a$ = Left(lista_PhoneSales.List(t), 20)
  If UCase(RTrim(a$)) = agente$ Then
     existe2 = 1
     Exit For
  End If
Next t
 
If existe2 = 1 Then GoTo brincado
 
 
 
NB_err = 0
 
For t = 0 To List8.ListCount - 1
  inicial1$ = Left(List8.List(t), 3)
  num_veces = Val(Mid(List8.List(t), 5, 3))
  cant = Val(Right(List8.List(t), 7))
  
  If UCase(RTrim(inicial1$)) = RTrim(UCase(lblinitials.Caption)) And lblinitials.Caption <> "" Then
     existe = 1
     multa = multa + cant
     NB_err = num_veces
  End If
  
Next t



brincado:



If existe = 1 Then
  lbldeduction.Caption = Format(multa, "$###,##0.00")
  lblnb_deduc.Caption = Format(NB_err, "#0")
  If lblnb_deduc.Caption = "" Then
      lblnb_deduc.Caption = "0"
  End If

Else
  lbldeduction.Caption = "$0.00"
  lblnb_deduc.Caption = "0"
End If





lblinvoice.Caption = Format(suma_invoice, "$###,##0.00")
lblnb.Caption = lbl_total_CSR_user.Caption
lblbf.Caption = lbltotal_invoices.Caption


bf = Val(Format(lblbf.Caption, "00000.00"))
d = Val(Format(lbldeduction.Caption, "00000.00"))
lbltotal_bf.Caption = Format(bf - d, "$###,##0.00")

lbltotal_final.Caption = lbltotal_bf.Caption


' actualiza el total de INVOICES
total_invoices = 0
For t = 0 To List9.ListCount - 1
   total_invoices = total_invoices + Val(List9.List(t))
Next t



If existe_monterrey <> 1 Then
  lblinvoice1.Caption = Format(total_facturas_propias, "$###,##0.00")
  lblinvoice2.Caption = Format(total_facturas_ajenas, "$###,##0.00")
  lblinvoice.Caption = Format(total_invoices, "$###,##0.00")

Else
  
  lblinvoice1.Caption = "$0.00"
  lblinvoice2.Caption = "$0.00"
  lblinvoice.Caption = "$0.00"
End If




' calcula los NB
nb = Val(lblnb.Caption)
NB_DED = Val(lblnb_deduc.Caption)
total_nb = nb - NB_DED

lbltotal_NB.Caption = Format(total_nb, "#0.0")
If Right(lbltotal_NB.Caption, 2) = ".0" Then
   lbltotal_NB.Caption = Left(lbltotal_NB.Caption, Len(lbltotal_NB.Caption) - 2)
End If

asigna_limite

carga_tier_del_agente


'porcentaje
If Left(lblporcentaje.Caption, 1) <> "$" And lblporcentaje.Caption <> "" Then
  comision = (bf - d)
  porciento = Val(Left(lblporcentaje.Caption, Len(lblporcentaje.Caption) - 1))
  Total = (comision * porciento) / 100
  lblcommission.Caption = Format(Total, "$###,##0.00")
ElseIf Left(lblcommission.Caption, 1) = "$" Then
  
Else
 lblcommission.Caption = ""
End If

lblid.Caption = lblidentificacion.Caption


num_fila_agente = lista_agentes.ListIndex



GoTo salta_dmv

' detecta si tiene multa del DMV
If grid6.Rows > 1 Then
 lbldmv.Caption = ""
 lblpenalty_dmv.Caption = ""
 
 For t = 1 To grid6.Rows - 1
   grid6.row = t
   grid6.col = 1
   agent_dmv$ = grid6.Text
   
   grid6.col = 2
   cant_dmv$ = grid6.Text
   
   If UCase(agent_dmv$) = UCase(lblinitials.Caption) And agent_dmv$ <> "" Then
      If Val(cant_dmv$) < 400 Then
        lbldmv.Caption = Format(cant_dmv$, "$###,##0.00")
        lblpenalty_dmv.Caption = "$40"
        
        comision_obtenida = Val(Format(lblcommission.Caption, "000000.00"))
        comision_final = comision_obtenida - 40
        lblcommission.Caption = Format(comision_final, "$###,##0.00")
 
        Exit For
      Else
        lbldmv.Caption = Format(cant_dmv$, "$###,##0.00")
        lblpenalty_dmv.Caption = "NONE"
        Exit For
      End If
   
   End If
   
 Next t
 
 
End If


salta_dmv:

' agrega las columnas faltantes
' ******************************************************************************************************
' ******************************************************************************************************



existe = 0
For Y = 1 To grid6.Rows - 1
  grid6.row = Y
  grid6.col = 1
  identifica = Val(grid6.Text)
  
  grid6.col = 4
  fecha_ini$ = grid6.Text
  
  
  
  
  If Val(lblidentificacion.Caption) = identifica And Format(txtdatefrom.Text, "mm/dd/yyyy") = Format(fecha_ini$, "mm/dd/yyyy") Then

' agrega Invoice
      
      grid6.row = Y
      grid6.col = 19
      grid6.Text = Format(lblinvoice.Caption, "####0.00")
      
      grid6.row = 0
      grid6.Text = "Invoice"

' agrega Deduction
      grid6.row = Y
      grid6.col = 20
      grid6.Text = Format(lbldeduction.Caption, "####0.00")
      x3 = Val(Format(lbldeduction.Caption, "####0.00"))

      grid6.row = 0
      grid6.Text = "Deduction"
      
      
' agrega NB Deduction
      grid6.row = Y
      grid6.col = 21
      grid6.Text = Format(lblnb_deduc.Caption, "####0.00")
     
      grid6.row = 0
      grid6.Text = "NB Deduction"
      
      
      
      
 ' Agrega BF
      grid6.row = Y
      grid6.col = 22
      grid6.Text = Format(lblbf.Caption, "####0.00")
      X1 = Val(Format(lblbf.Caption, "####0.00"))

       grid6.row = 0
      grid6.Text = "BF"
      
      
 ' Agrega NB
      grid6.row = Y
      grid6.col = 23
      grid6.Text = Format(lblnb.Caption, "####0.00")
      

       grid6.row = 0
      grid6.Text = "NB"
      
 

' Agrega porcentaje
      grid6.row = Y
      grid6.col = 24
      X2 = Val(Format(lblporcentaje.Caption, "####0.00")) * 100
      grid6.Text = Format(X2, "####0.00")
      
      grid6.row = 0
      grid6.Text = "%"
 
 
 ' Agrega Comm
      grid6.row = Y
      grid6.col = 25
      comm_total = (X1 - x3) * (X2) / 100
      grid6.Text = Format(comm_total, "####0.00")
      
      If Val(Format(lblcommission.Caption, "0000.00")) > 0 Then
          grid6.Text = Format(Val(Format(lblcommission.Caption, "0000.00")), "####0.00")
      End If

      grid6.row = 0
      grid6.Text = "Comm"
      
      
      ' Agrega BF
      grid6.row = Y
      grid6.col = 26
      grid6.Text = Format(lbltotal_bf.Caption, "####0.00")
     
       grid6.row = 0
      grid6.Text = "TOTAL BF"
      
      
 ' Agrega NB
      grid6.row = Y
      grid6.col = 27
      grid6.Text = Format(lbltotal_NB.Caption, "####0.00")
     
       grid6.row = 0
      grid6.Text = "TOTAL NB"
      
      
 ' Agrega LAE empl
      grid6.row = Y
      grid6.col = 28
      grid6.Text = Format(lbllae.Caption, "####0")
     
      grid6.row = 0
      grid6.Text = "LAE#"
      
      
       ' Agrega nota
      grid6.row = Y
      grid6.col = 29
      grid6.Text = txtnotes.Text
     
      grid6.row = 0
      grid6.Text = "Notes"


  End If
Next Y


msg.Visible = False


End Sub



Public Sub calcula_CSR_and_USER()

On Error Resume Next

' =================================================================================================================================
' RUTINA ORIGINAL

Dim sSelect As String
Dim Rs As ADODB.Recordset
    
Set Rs = New ADODB.Recordset

lista_users_shared.Clear
total_csr = 0
total_agent = 0
total_csr2 = 0
List11.Clear
List5.Clear
List6.Clear



For t = 1 To Grid1.Rows - 1
   Grid1.col = 5
   Grid1.row = t
   csr$ = UCase(Grid1.Text)  ' CSR
   
   Grid1.col = 4
   user$ = UCase(Grid1.Text)
   
   Grid1.col = 1
   recibo$ = Grid1.Text
   
   SI_COMERCIAL = 0
   
     Es_agente_comercial = False
     concepto2$ = ""
     
   '   If UCase(csr$) = "MFUENTES" Then Stop
   '   If UCase(user$) = "MFUENTES" Then Stop
      
     '  If recibo$ = "266938" Then Stop
   
   
   If agente$ = csr$ Or agente$ = user$ Then
   
    recibo2$ = ""
    concepto2$ = ""
   
   
     sSelect = "SELECT " & _
    "recdtl.[IdReceiptHDR] as [Receipt #] " & _
    ",iitem.InvoiceItemName as [Invoice Item] " & _
    "FROM  [ReceiptsDTL] recdtl " & _
    "inner join  ReceiptsHDR  rechdr on recdtl.IdReceiptHDR=rechdr.IDReceiptHDR " & _
    "inner join PoliciesHDR polhdr on polhdr.IdPoliciesHDR=rechdr.IdPoliciesHDR " & _
    "inner join EmployeeInfo emp on emp.IDEmployee=rechdr.IdEmployeeUSR " & _
    "inner join EmployeeInfo csr on csr.IDEmployee=rechdr.IdEmployeeCSR1 " & _
    "inner join OfficesCatalog ofc on ofc.IdOffice=rechdr.IdOffice " & _
    "inner join OfficesCatalog ofcpol on ofcpol.IdOffice=rechdr.IdOffice " & _
    "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
    "inner join Customers cus on cus.IdCustomer=polhdr.IdCustomer " & _
    "inner join InvoiceItemCatalog iitem on iitem.IdInvoiceItem=recdtl.IdInvoiceItem " & _
    "inner join BankAccountCatalog bank on bank.IdBankAccount=iitem.IdBankAccount " & _
    "inner join CompanysAGICatalog ja on ja.IdCompanyAGI=polhdr.IdCompanyAGI " & _
    "where iitem.IdInvoiceItem in (1,20,2,35,3,11,15,22,23,33,9,30,33,17,13,31,32, 36, 37,39,42,43) " & _
    "and cast(rechdr.Date as Date) >= '" + txtdatefrom.Text + "' AND cast( rechdr.DATE as Date) <= '" + txtdateto.Text + "' " & _
    "and rechdr.Active=1 and recdtl.IdReceiptHDR='" + recibo$ + "'   order by [Receipt #]"
    
    ' and iitem.InvoiceItemName='BF Commercial'   se elimino esta parte  2/21
   
    Rs.Open sSelect, base, adOpenUnspecified
    recibo2$ = Rs(0)
    concepto2$ = Rs(1)
    Rs.Close
  
  
  
  
     recibo2b$ = ""
    concepto2b$ = ""
   
   
     sSelect = "SELECT " & _
    "recdtl.[IdReceiptHDR] as [Receipt #] " & _
    ",iitem.InvoiceItemName as [Invoice Item] " & _
    "FROM  [ReceiptsDTL] recdtl " & _
    "inner join  ReceiptsHDR  rechdr on recdtl.IdReceiptHDR=rechdr.IDReceiptHDR " & _
    "inner join PoliciesHDR polhdr on polhdr.IdPoliciesHDR=rechdr.IdPoliciesHDR " & _
    "inner join EmployeeInfo emp on emp.IDEmployee=rechdr.IdEmployeeUSR " & _
    "inner join EmployeeInfo csr on csr.IDEmployee=rechdr.IdEmployeeCSR1 " & _
    "inner join OfficesCatalog ofc on ofc.IdOffice=rechdr.IdOffice " & _
    "inner join OfficesCatalog ofcpol on ofcpol.IdOffice=rechdr.IdOffice " & _
    "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
    "inner join Customers cus on cus.IdCustomer=polhdr.IdCustomer " & _
    "inner join InvoiceItemCatalog iitem on iitem.IdInvoiceItem=recdtl.IdInvoiceItem " & _
    "inner join BankAccountCatalog bank on bank.IdBankAccount=iitem.IdBankAccount " & _
    "inner join CompanysAGICatalog ja on ja.IdCompanyAGI=polhdr.IdCompanyAGI " & _
    "where iitem.IdInvoiceItem in (1,20,2,35,3,11,15,22,23,33,9,30,33,17,13,31,32, 36, 37,39,42,43) " & _
    "and cast(rechdr.Date as Date) >= '" + txtdatefrom.Text + "' AND cast( rechdr.DATE as Date) <= '" + txtdateto.Text + "' " & _
    "and rechdr.Active=1 and recdtl.IdReceiptHDR='" + recibo$ + "' and iitem.InvoiceItemName='BF Commercial'  order by [Receipt #]"
    
    '
   
    Rs.Open sSelect, base, adOpenUnspecified
    recibo2b$ = Rs(0)
    concepto2b$ = Rs(1)
    Rs.Close
    
    
    If recibo2$ = recibo2b$ Then
       ' es un BF COMMERCIAL
       
       SI_COMERCIAL = 1
    End If
    
  
  
  
   
    If recibo2$ = recibo$ Then
    
      recibo3$ = ""
      concepto3$ = ""
    
      sSelect = "SELECT " & _
      "recdtl.[IdReceiptHDR] as [Receipt #] " & _
      ",iitem.InvoiceItemName as [Invoice Item] " & _
      "FROM  [ReceiptsDTL] recdtl " & _
      "inner join  ReceiptsHDR  rechdr on recdtl.IdReceiptHDR=rechdr.IDReceiptHDR " & _
      "inner join PoliciesHDR polhdr on polhdr.IdPoliciesHDR=rechdr.IdPoliciesHDR " & _
      "inner join EmployeeInfo emp on emp.IDEmployee=rechdr.IdEmployeeUSR " & _
      "inner join EmployeeInfo csr on csr.IDEmployee=rechdr.IdEmployeeCSR1 " & _
      "inner join OfficesCatalog ofc on ofc.IdOffice=rechdr.IdOffice " & _
      "inner join OfficesCatalog ofcpol on ofcpol.IdOffice=rechdr.IdOffice " & _
      "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
      "inner join Customers cus on cus.IdCustomer=polhdr.IdCustomer " & _
      "inner join InvoiceItemCatalog iitem on iitem.IdInvoiceItem=recdtl.IdInvoiceItem " & _
      "inner join BankAccountCatalog bank on bank.IdBankAccount=iitem.IdBankAccount " & _
      "inner join CompanysAGICatalog ja on ja.IdCompanyAGI=polhdr.IdCompanyAGI " & _
      "where iitem.IdInvoiceItem in (1,20,2,35,3,11,15,22,23,33,9,30,33,17,13,31,32, 36, 37,39,42,43) " & _
      "and cast(rechdr.Date as Date) >= '" + txtdatefrom.Text + "' AND cast( rechdr.DATE as Date) <= '" + txtdateto.Text + "' " & _
      "and rechdr.Active=1 and recdtl.IdReceiptHDR='" + recibo$ + "' and (iitem.InvoiceItemName='BF' or iitem.InvoiceItemName='BF CALL CENTER')   order by [Receipt #]"
     
      Rs.Open sSelect, base, adOpenUnspecified
      recibo3$ = Rs(0)
      concepto3$ = Rs(1)
      Rs.Close
    
    
      
      
    
      If recibo3$ = recibo$ Or (UCase(concepto2$) = "BF COMMERCIAL" And recibo$ = recibo2$) Then
          ' verifica si ambos son manager commercial
          existe = 0
          manager_csr$ = ""
          manager_user$ = ""
          manager_agente$ = ""
          
         For z = 0 To lista_managers.ListCount - 1
            a$ = RTrim(UCase(Left$(lista_managers.List(z), 20)))
            b$ = LTrim(RTrim(Right(UCase(lista_managers.List(z)), Len(lista_managers.List(z)) - 20)))
        
           ' If a$ = "MFUENTES" Then Stop
        
            If a$ = csr$ Then '
               If b$ = "MANAGER" Or b$ = "MANAGER_COMMERCIAL" Or b$ = "BF COMMERCIAL" Then
                  manager_csr$ = a$
                  existe = 1
             
               End If
               
               If b$ = "COMMERCIAL" Then
                 
                  manager_csr$ = a$
                  manager_commercial$ = a$
                  existe = 2
                  
               End If
               
               
               
               If b$ = "MONTERREY" Then
                 
                  Es_Monterrey = 1
                  
               End If
               
               
          
            End If
        
        
           If a$ = user$ Then
             If b$ = "MANAGER" Or b$ = "MANAGER_COMMERCIAL" Or b$ = "BF COMMERCIAL" Then
                manager_user$ = a$
                existe = 1
             
             End If
             
             
             If b$ = "COMMERCIAL" Then
                manager_commercial$ = a$
                manager_user$ = a$
                existe = 2
                
             End If
             
             If b$ = "MONTERREY" Then
                 
                  Es_Monterrey = 1
                  
             End If
             
           End If
        
        
           If a$ = agente$ Then
              If b$ = "MANAGER" Or b$ = "MANAGER_COMMERCIAL" Or b$ = "BF COMMERCIAL" Then
                  manager_agente$ = a$
                  existe = 1
             
              End If
             
             
              If b$ = "COMMERCIAL" Then
                  manager_commercial$ = a$
                  manager_agente$ = a$
                  existe = 2
                  
              End If
             
             
               If b$ = "MONTERREY" Then
                 
                  Es_Monterrey = 1
                  
               End If
               
               
           End If
        
        
                  
        
                  
        
           If a$ = agente$ Or a$ = csr$ Or a$ = user$ Then
              If b$ = "AGENT_COMMERCIAL" Then
                 Agent_Oficinacommercial$ = a$
                 If SI_COMERCIAL = 1 Then
                       manager_commercial$ = a$  'se elimino 4/1/2024
                 End If
                 Es_agente_comercial = True
                 existe = 1
               
              End If
              
              If b$ = "MONTERREY" Then
                 
                  Es_Monterrey = 1
                  
               End If
             
           End If
        
        
           
        
         ' If existe = 1 Then
         '    Exit For
         ' End If
        
        
        
     Next z
     
     
         If manager_commercial$ = Agent_Oficinacommercial$ And manager_commercial$ <> "" Then
           GoTo BF_valido
         End If
     
     
         
         
         
         
         If recibo3$ = "" And (UCase(concepto2$) = "BF COMMERCIAL" And recibo$ = recibo2$) Then
          GoTo BF_valido
         End If
         
         
         If (recibo3$ = recibo$ And concepto3$ = "BF") And (UCase(concepto2$) = "BF COMMERCIAL" And recibo$ = recibo2$) Then
          GoTo BF_valido
         End If
         
         
         If manager_agente$ = manager_csr$ And manager_agente$ <> manager_user$ And manager_user$ <> "" And manager_commercial$ = manager_agente$ And manager_agente$ <> "" Then
           GoTo brinca
         End If
         
         
         
         If existe = 2 And manager_agente$ <> "" Then
            ' GoTo brinca
         End If
      
      
         GoTo BF_valido
      Else
       
         GoTo brinca
     
      End If
      
    End If
   
   End If
      
BF_valido:
   
   manager_csr$ = ""
   manager_user$ = ""
   manager_agente$ = ""
   ' manager_commercial$ = ""
   Agent_Oficinacommercial$ = ""
   
      
    ' verifica si es manager
     existe = 0
     Es_Monterrey = 0
     
     For z = 0 To lista_managers.ListCount - 1
       a$ = RTrim(UCase(Left$(lista_managers.List(z), 20)))
       b$ = LTrim(RTrim(Right(UCase(lista_managers.List(z)), Len(lista_managers.List(z)) - 20)))
        
        If a$ = csr$ Then '
          If b$ = "MANAGER" Or b$ = "MANAGER_COMMERCIAL" Then
             manager_csr$ = a$
             existe = 1
             'Exit For
          End If
          
        End If
        
        
        If a$ = user$ Then
          If b$ = "MANAGER" Or b$ = "MANAGER_COMMERCIAL" Then
             manager_user$ = a$
             existe = 1
             'Exit For
          End If
             
        End If
        
        
        If a$ = agente$ Then
          If b$ = "MANAGER" Or b$ = "MANAGER_COMMERCIAL" Then
             manager_agente$ = a$
             existe = 1
             'Exit For
          End If
             
        End If
        
        
        If a$ = agente$ Or a$ = csr$ Or a$ = user$ Then
          If b$ = "MANAGER_COMMERCIAL" Then
             manager_commercial$ = a$
             existe = 1
             'Exit For
          End If
             
        End If
        
        
        If a$ = agente$ Or a$ = csr$ Or a$ = user$ Then
          If b$ = "COMMERCIAL" Then
             manager_commercial$ = a$
             existe = 1
             'Exit For
          End If
             
        End If
        
        
        
        If a$ = agente$ Or a$ = csr$ Or a$ = user$ Then
          If b$ = "AGENT_COMMERCIAL" Then
             Agent_Oficinacommercial$ = a$
             existe = 1
             'Exit For
          End If
             
        End If
        
        
        If a$ = agente$ Or a$ = csr$ Or a$ = user$ Then
            If b$ = "MONTERREY" Then
               Es_Monterrey = 1
               'Exit For
            End If
        End If
        
        
        If existe = 1 Or Es_Monterrey = 1 Then
            ' Exit For
        End If
        
        
        
     Next z
     
     
   
   
     
     
     If csr$ = manager_commercial$ And csr$ <> "" Then
       manager_csr$ = manager_commercial$
     End If
     
     If agente$ = manager_commercial$ And agente$ <> csr$ Then
       manager_agente$ = manager_commercial$
     End If
     
     If user$ = manager_commercial$ And user$ <> "" Then
       manager_user$ = manager_commercial$
     End If
     
     
 
     
 
   
   
   
  ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
   
   If agente$ = user$ And agente$ <> csr$ Then
   
  
        If Es_Monterrey = 1 Then
            total_user = total_user + 1
            If agente$ = csr$ Then
            
            Else
                  existe = 0
                  cont = 0
                  For Y = 0 To lista_users_shared.ListCount - 1
                         If csr$ = Left(lista_users_shared.List(Y), Len(lista_users_shared.List(Y)) - 3) Then   ' aqui habia user
                                  cont = Right(lista_users_shared.List(Y), 2) + 1
                                  lista_users_shared.RemoveItem Y
                                  lista_users_shared.AddItem csr$ + " " + Format(cont, "00")
                                  existe = 1
                                   Exit For
                         End If
                  Next Y
       
                  If existe = 0 Then
                         lista_users_shared.AddItem csr$ + " " + "01"
                  End If
                                             
            End If
            GoTo brinca
        End If
        
        
  
       ' verifica oficinas a que pertenecen
                  For w = 1 To Val(ubicacion(0, 0))
                      If UCase(RTrim(ubicacion(w, 0))) = UCase(user$) Then
                          oficina_user$ = RTrim(ubicacion(w, 1))
                          If RTrim(ubicacion(w, 2)) = "" Then ubicacion(w, 2) = "None"
                          oficina_user2$ = RTrim(ubicacion(w, 2))
                      End If
                  
                      If UCase(RTrim(ubicacion(w, 0))) = UCase(csr$) Then
                          oficina_csr$ = RTrim(ubicacion(w, 1))
                          If RTrim(ubicacion(w, 2)) = "" Then ubicacion(w, 2) = "Nada"
                          oficina_csr2$ = RTrim(ubicacion(w, 2))
                      End If
                  Next w
                  
                  
                  If oficina_user2$ = "None" And oficina_csr2$ = "None" Then
                     oficina_user2$ = "None2"
                  End If
                  
                  
                  If oficina_user2$ = "Nada" And oficina_csr2$ = "Nada" Then
                     oficina_user2$ = "Nada2"
                  End If
                  
                   
                  If oficina_user$ = oficina_csr$ Or oficina_user2$ = oficina_csr2$ Or oficina_user$ = oficina_csr2$ Or oficina_user2$ = oficina_csr$ Then
                         
                         If (user$ = manager_user$ And csr$ <> manager_csr$) Then
                          
                           If Es_agente_comercial = True Then
                           
                                        sSelect = "SELECT " & _
                                        "recdtl.[IdReceiptHDR] as [Receipt #] " & _
                                        ",iitem.InvoiceItemName as [Invoice Item] " & _
                                        "FROM  [ReceiptsDTL] recdtl " & _
                                        "inner join  ReceiptsHDR  rechdr on recdtl.IdReceiptHDR=rechdr.IDReceiptHDR " & _
                                        "inner join PoliciesHDR polhdr on polhdr.IdPoliciesHDR=rechdr.IdPoliciesHDR " & _
                                        "inner join EmployeeInfo emp on emp.IDEmployee=rechdr.IdEmployeeUSR " & _
                                        "inner join EmployeeInfo csr on csr.IDEmployee=rechdr.IdEmployeeCSR1 " & _
                                        "inner join OfficesCatalog ofc on ofc.IdOffice=rechdr.IdOffice " & _
                                        "inner join OfficesCatalog ofcpol on ofcpol.IdOffice=rechdr.IdOffice " & _
                                        "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
                                        "inner join Customers cus on cus.IdCustomer=polhdr.IdCustomer " & _
                                        "inner join InvoiceItemCatalog iitem on iitem.IdInvoiceItem=recdtl.IdInvoiceItem " & _
                                        "inner join BankAccountCatalog bank on bank.IdBankAccount=iitem.IdBankAccount " & _
                                        "inner join CompanysAGICatalog ja on ja.IdCompanyAGI=polhdr.IdCompanyAGI " & _
                                        "where iitem.IdInvoiceItem in (2,35,3,11,15,22,23,33,9,30,33,17,13,31,32, 36, 37,39,42,43) " & _
                                        "and cast(rechdr.Date as Date) >= '" + txtdatefrom.Text + "' AND cast( rechdr.DATE as Date) <= '" + txtdateto.Text + "' " & _
                                        "and rechdr.Active=1 and recdtl.IdReceiptHDR='" + recibo$ + "' and (iitem.InvoiceItemName='BF COMMERCIAL')   order by [Receipt #]"
     
                                        Rs.Open sSelect, base, adOpenUnspecified
                                        recibo4$ = Rs(0)
                                        concepto4$ = Rs(1)
                                        Rs.Close
                                        
                                        
                                        If concepto4$ <> "" Then
                                        
                                              existe = 0
                                              cont = 0
                                              For Y = 0 To List11.ListCount - 1
                                                    If csr$ = Left(List11.List(Y), Len(List11.List(Y)) - 3) Then   ' aqui habia user
                                                        cont = Right(List11.List(Y), 2) + 1
                                                        List11.RemoveItem Y
                                                        List11.AddItem csr$ + " " + Format(cont, "00")
                                                        existe = 1
                                                        Exit For
                                                    End If
                                              Next Y
       
                                              If existe = 0 Then
                                                    List11.AddItem csr$ + " " + "01"
        
                                              End If
                                                           
                                              GoTo brinca
                                        
                                        Else
                                        
                                             existe = 0
                                             cont = 0
                                
                                             total_user = total_user + 0.5   ' Se agrego esta linea
                                
                                             For Y = 0 To lista_users_shared.ListCount - 1 '
                                                If user$ = Left(lista_users_shared.List(Y), Len(lista_users_shared.List(Y)) - 3) Then   '
                                                    cont = Right(lista_users_shared.List(Y), 2) + 1
                                                    lista_users_shared.RemoveItem Y
                                                    lista_users_shared.AddItem csr$ + " " + Format(cont, "00")
                                                    existe = 1
                                                    Exit For
                                                End If
                                             Next Y
       
                                             If existe = 0 Then
                                                 lista_users_shared.AddItem csr$ + " " + "01"
        
                                             End If
                                             
                                             GoTo brinca
                                        
                                           
                                        End If
 
                           
                           
                           
                           Else
                          
                               existe = 0
                                cont = 0
                                For Y = 0 To List11.ListCount - 1
                                        If csr$ = Left(List11.List(Y), Len(List11.List(Y)) - 3) Then   ' aqui habia user
                                                    cont = Right(List11.List(Y), 2) + 1
                                                    List11.RemoveItem Y
                                                    List11.AddItem csr$ + " " + Format(cont, "00")
                                                    existe = 1
                                                    Exit For
                                        End If
                                Next Y
       
                                If existe = 0 Then
                                        List11.AddItem csr$ + " " + "01"
        
                                End If
                                                           
                                GoTo brinca
                                
                           End If
                             
                         End If
                         
                         
                         
                         
                         
                         
                         If user$ = manager_user$ Then
                             total_csr2 = total_csr2 + 1
                             existe = 0
                                cont = 0
                                For Y = 0 To List5.ListCount - 1
                                        If user$ = Left(List5.List(Y), Len(List5.List(Y)) - 3) Then   ' aqui habia user
                                                    cont = Right(List5.List(Y), 2) + 1
                                                    List5.RemoveItem Y
                                                    List5.AddItem user$ + " " + Format(cont, "00")
                                                    existe = 1
                                                    Exit For
                                        End If
                                Next Y
       
                                If existe = 0 Then
                                        List5.AddItem user$ + " " + "01"
        
                                End If
                                                           
                                GoTo brinca
                         End If
                         
                         If csr$ = manager_csr$ Then
                             total_csr2 = total_csr2 + 1
                             existe = 0
                                cont = 0
                                For Y = 0 To List5.ListCount - 1
                                        If csr$ = Left(List5.List(Y), Len(List5.List(Y)) - 3) Then   ' aqui habia user
                                                    cont = Right(List5.List(Y), 2) + 1
                                                    List5.RemoveItem Y
                                                    List5.AddItem csr$ + " " + Format(cont, "00")   'user$ + " " + Format(cont, "00")
                                                    existe = 1
                                                    Exit For
                                        End If
                                Next Y
       
                                If existe = 0 Then
                                        List5.AddItem csr$ + " " + "01"    ' user$ + " " + "01"
        
                                End If
                                                           
                                GoTo brinca
                         End If
                         
                         If (user$ <> manager_user$) And (csr$ <> manager_csr$) And (user$ <> csr$) Then
                             total_user = total_user + 0.5
                             'total_csr2 = total_csr2 + 0.5
                         End If
                         
                         If user$ = manager_commercial$ Or csr$ = manager_commercial$ Then
                               total_csr2 = total_csr2 + 1
                               existe = 0
                                cont = 0
                                For Y = 0 To List5.ListCount - 1
                                        If user$ = Left(List5.List(Y), Len(List5.List(Y)) - 3) Then   ' aqui habia user
                                                    cont = Right(List5.List(Y), 2) + 1
                                                    List5.RemoveItem Y
                                                    List5.AddItem user$ + " " + Format(cont, "00")
                                                    existe = 1
                                                    Exit For
                                        End If
                                Next Y
       
                                If existe = 0 Then
                                      '  List5.AddItem user$ + " " + "01"
        
                                End If
                                                           
                                GoTo brinca
                                                           
                               
                         End If
                         
                         
                         
                  Else
                         
                         If ((user$ = manager_commercial$) Or (csr$ = manager_commercial$)) And agente$ <> manager_commercial$ And agente <> manager_agente$ Then
                              
                               total_csr2 = total_csr2 + 1
                             existe = 0
                                cont = 0
                                For Y = 0 To List5.ListCount - 1
                                        If csr$ = Left(List5.List(Y), Len(List5.List(Y)) - 3) Then   ' aqui habia user
                                                    cont = Right(List5.List(Y), 2) + 1
                                                    List5.RemoveItem Y
                                                    List5.AddItem csr$ + " " + Format(cont, "00")
                                                    existe = 1
                                                    Exit For
                                        End If
                                Next Y
       
                                If existe = 0 Then
                                   
                                        List5.AddItem csr$ + " " + "01"
                                   
        
                                End If
                                                           
                                GoTo brinca
                                                           
                              
                         End If
                  
                         If user$ <> manager_user$ And csr$ <> manager_csr$ Then
                             total_user = total_user + 0.5
                             'total_csr2 = total_csr2 + 0.5
                         End If
                         
                         If (user$ = manager_user$) And (csr$ = manager_csr$) And (csr$ <> manager_commercial$) Then
                             total_user = total_user + 0.5
                             'total_csr2 = total_csr2 + 0.5
                         End If
                         
                         If ((user$ = manager_user$ And csr$ <> manager_csr$) Or (user$ <> manager_user$ And csr$ = manager_csr$)) And manager_commercial$ = "" Then
                           
                             total_user = total_user + 0.5
                             'total_csr2 = total_csr2 + 0.5
                         End If
                         
                         
                          If (user$ = manager_user$) And (csr$ = manager_commercial$) Then
                             total_user = total_user + 1
                              existe = 0
                              cont = 0
                              For Y = 0 To List5.ListCount - 1
                                        If csr$ = Left(List5.List(Y), Len(List5.List(Y)) - 3) Then   ' aqui habia user
                                                    cont = Right(List5.List(Y), 2) + 1
                                                    List5.RemoveItem Y
                                                    List5.AddItem csr$ + " " + Format(cont, "00")
                                                    existe = 1
                                                    Exit For
                                        End If
                              Next Y
       
                              If existe = 0 Then
                                        List5.AddItem csr$ + " " + "01"
        
                              End If
                                                           
                              GoTo brinca
                         End If
                        
                        
                        
                          If (user$ = manager_commercial$) And (csr$ <> user$) Then
                             ' GoTo brinca
                             
                            
                              existe = 0
                              cont = 0
                              For Y = 0 To List11.ListCount - 1
                                        If csr$ = Left(List11.List(Y), Len(List11.List(Y)) - 3) Then   ' aqui habia user
                                                    cont = Right(List11.List(Y), 2) + 1
                                                    List11.RemoveItem Y
                                                    List11.AddItem csr$ + " " + Format(cont, "00")
                                                    existe = 1
                                                    Exit For
                                        End If
                              Next Y
       
                              If existe = 0 Then
                                        List11.AddItem csr$ + " " + "01"
        
                              End If
                                                           
                              GoTo brinca
                         End If
                                                      
                  End If
                  
                  
                    
       
       existe = 0
       cont = 0
       For Y = 0 To lista_users_shared.ListCount - 1
         If csr$ = Left(lista_users_shared.List(Y), Len(lista_users_shared.List(Y)) - 3) Then   ' aqui habia user
             cont = Right(lista_users_shared.List(Y), 2) + 1
             lista_users_shared.RemoveItem Y
             lista_users_shared.AddItem csr$ + " " + Format(cont, "00")
             existe = 1
             Exit For
          End If
       Next Y
       
       If existe = 0 Then
          
          lista_users_shared.AddItem csr$ + " " + "01"
        
       End If
       
    
     GoTo brinca
   End If
  
  
  
  
  
 
  
  
  If agente$ = csr$ And agente$ <> user$ Then
     
        If Es_Monterrey = 1 Then
            total_csr2 = total_csr2 + 1
            If agente$ = csr$ Then
              total_csr2 = total_csr2 - 1
            End If
            GoTo brinca
        End If
        
        
        ' verifica oficinas a que pertenecen
                  For w = 1 To Val(ubicacion(0, 0))
                      If UCase(RTrim(ubicacion(w, 0))) = UCase(user$) Then
                          oficina_user$ = RTrim(ubicacion(w, 1))
                          
                          If RTrim(ubicacion(w, 2)) = "" Then ubicacion(w, 2) = "None"
                          oficina_user2$ = RTrim(ubicacion(w, 2))
                          
                      End If
                  
                      If UCase(RTrim(ubicacion(w, 0))) = UCase(csr$) Then
                          oficina_csr$ = RTrim(ubicacion(w, 1))
                          
                          If RTrim(ubicacion(w, 2)) = "" Then ubicacion(w, 2) = "Nada"
                          oficina_csr2$ = RTrim(ubicacion(w, 2))
                      End If
                  Next w
                  
                   
                  If oficina_user2$ = "None" And oficina_csr2$ = "None" Then
                     oficina_user2$ = "None2"
                  End If
                  
                  If oficina_user2$ = "Nada" And oficina_csr2$ = "Nada" Then
                     oficina_user2$ = "Nada2"
                  End If
                   
                   
                   ' SI SON DE LA MISMA OFICINA
                  If oficina_user$ = oficina_csr$ Or oficina_user2$ = oficina_csr2$ Or oficina_user$ = oficina_csr2$ Or oficina_user2$ = oficina_csr$ Then
                         
                         If (user$ = manager_user$ And csr$ <> user$) Then
                         
                          
                               If Es_agente_comercial = True And ((agent_commercial = csr$ And manager_commercial <> user$ And manager_commercial <> "") Or (agent_commercial = user$ And manager_commercial <> csr$ And manager_commercial <> "")) Then
                               
                                         sSelect = "SELECT " & _
                                        "recdtl.[IdReceiptHDR] as [Receipt #] " & _
                                        ",iitem.InvoiceItemName as [Invoice Item] " & _
                                        "FROM  [ReceiptsDTL] recdtl " & _
                                        "inner join  ReceiptsHDR  rechdr on recdtl.IdReceiptHDR=rechdr.IDReceiptHDR " & _
                                        "inner join PoliciesHDR polhdr on polhdr.IdPoliciesHDR=rechdr.IdPoliciesHDR " & _
                                        "inner join EmployeeInfo emp on emp.IDEmployee=rechdr.IdEmployeeUSR " & _
                                        "inner join EmployeeInfo csr on csr.IDEmployee=rechdr.IdEmployeeCSR1 " & _
                                        "inner join OfficesCatalog ofc on ofc.IdOffice=rechdr.IdOffice " & _
                                        "inner join OfficesCatalog ofcpol on ofcpol.IdOffice=rechdr.IdOffice " & _
                                        "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
                                        "inner join Customers cus on cus.IdCustomer=polhdr.IdCustomer " & _
                                        "inner join InvoiceItemCatalog iitem on iitem.IdInvoiceItem=recdtl.IdInvoiceItem " & _
                                        "inner join BankAccountCatalog bank on bank.IdBankAccount=iitem.IdBankAccount " & _
                                        "inner join CompanysAGICatalog ja on ja.IdCompanyAGI=polhdr.IdCompanyAGI " & _
                                        "where iitem.IdInvoiceItem in (2,35,3,11,15,22,23,33,9,30,33,17,13,31,32, 36, 37,39,42,43) " & _
                                        "and cast(rechdr.Date as Date) >= '" + txtdatefrom.Text + "' AND cast( rechdr.DATE as Date) <= '" + txtdateto.Text + "' " & _
                                        "and rechdr.Active=1 and recdtl.IdReceiptHDR='" + recibo$ + "' and (iitem.InvoiceItemName='BF COMMERCIAL')   order by [Receipt #]"
     
                                        Rs.Open sSelect, base, adOpenUnspecified
                                        recibo4$ = Rs(0)
                                        concepto4$ = Rs(1)
                                        Rs.Close
                                        
                                        
                                        If concepto4$ = "" Then
                                             existe = 0
                                             cont = 0
                                
                                             total_user = total_user + 0.5   ' Se agrego esta linea
                                
                                             For Y = 0 To lista_users_shared.ListCount - 1 ' List11.ListCount - 1
                                                If user$ = Left(lista_users_shared.List(Y), Len(lista_users_shared.List(Y)) - 3) Then   ' aqui habia list11
                                                    cont = Right(lista_users_shared.List(Y), 2) + 1
                                                    lista_users_shared.RemoveItem Y
                                                    lista_users_shared.AddItem user$ + " " + Format(cont, "00")
                                                    existe = 1
                                                    Exit For
                                                End If
                                             Next Y
       
                                             If existe = 0 Then
                                                 lista_users_shared.AddItem user$ + " " + "01"
        
                                             End If
                                             
                                             GoTo brinca
                                        Else
                                           
                                        End If
                         
                         
                         End If
                         
                         
                                                                              
                                            total_csr2 = total_csr2 + 1
                                            existe = 0
                                            cont = 0
                                            For Y = 0 To List5.ListCount - 1
                                               If user$ = Left(List5.List(Y), Len(List5.List(Y)) - 3) Then   ' aqui habia user
                                                    cont = Right(List5.List(Y), 2) + 1
                                                    List5.RemoveItem Y
                                                    List5.AddItem user$ + " " + Format(cont, "00")
                                                    existe = 1
                                                    Exit For
                                               End If
                                            Next Y
       
                                            If existe = 0 Then
                                        
                                        
                                                 List5.AddItem user$ + " " + "01"
                                        
                                            End If
                                                           
                                            GoTo brinca
                                            
                                        
                               
                               
                         End If
                         
                         
                         
                         If user$ = manager_user$ Then
                             total_csr2 = total_csr2 + 1
                             existe = 0
                                cont = 0
                                For Y = 0 To List5.ListCount - 1
                                        If user$ = Left(List5.List(Y), Len(List5.List(Y)) - 3) Then   ' aqui habia user
                                                    cont = Right(List5.List(Y), 2) + 1
                                                    List5.RemoveItem Y
                                                    List5.AddItem user$ + " " + Format(cont, "00")
                                                    existe = 1
                                                    Exit For
                                        End If
                                Next Y
       
                                If existe = 0 Then
                                        List5.AddItem user$ + " " + "01"
        
                                End If
                                                           
                                GoTo brinca
                         End If
                         
                         
                                                  
                         
                         
                         
                         If csr$ = manager_csr$ Then
                              
                              If (UCase(concepto2$) = "BF COMMERCIAL" And UCase(concepto3$) = "BF") Then
                                 GoTo brinca
                              End If
                              
                              
                              If (UCase(concepto2b$) = "BF COMMERCIAL" And UCase(concepto3$) = "BF") Then
                                 GoTo lista_de_excepciones   '  Se agrego 8/26
                              End If
                              
                              
                              If (UCase(concepto2$) <> "BF COMMERCIAL" And UCase(concepto3$) = "BF") And Agent_Oficinacommercial$ = agente$ Then
                                 
                                      existe = 0
                                      cont = 0
                                
                                      total_user = total_user + 0.5   ' Se agrego esta linea
                                
                                      For Y = 0 To lista_users_shared.ListCount - 1 ' List11.ListCount - 1
                                        If user$ = Left(lista_users_shared.List(Y), Len(lista_users_shared.List(Y)) - 3) Then   ' aqui habia list11
                                                    cont = Right(lista_users_shared.List(Y), 2) + 1
                                                    lista_users_shared.RemoveItem Y
                                                    lista_users_shared.AddItem user$ + " " + Format(cont, "00")
                                                    existe = 1
                                                    Exit For
                                        End If
                                      Next Y
       
                                      If existe = 0 Then
                                        lista_users_shared.AddItem user$ + " " + "01"
        
                                      End If
                                   
                                 GoTo brinca
                              End If
                              
lista_de_excepciones:
                             
                             'total_csr2 = total_csr2 + 1
                             existe = 0
                                cont = 0
                                For Y = 0 To List11.ListCount - 1
                                        If user$ = Left(List11.List(Y), Len(List11.List(Y)) - 3) Then   ' aqui habia user
                                                    cont = Right(List11.List(Y), 2) + 1
                                                    List11.RemoveItem Y
                                                    List11.AddItem user$ + " " + Format(cont, "00")
                                                    existe = 1
                                                    Exit For
                                        End If
                                Next Y
       
                                If existe = 0 Then
                                   If Es_agente_comercial = True Then   ' se cambio a true   4/30
                                        List11.AddItem user$ + " " + "01"
                                        
                                   ElseIf (oficina_csr$ = oficina_user$) Or (oficina_csr2$ = oficina_user$) Then
                                        List11.AddItem user$ + " " + "01"
                                        
                                   Else
                                   
                                   
                                      existe = 0
                                      cont = 0
                                
                                      total_user = total_user + 0.5   ' Se agrego esta linea
                                
                                      For Y = 0 To lista_users_shared.ListCount - 1 ' List11.ListCount - 1
                                        If user$ = Left(lista_users_shared.List(Y), Len(lista_users_shared.List(Y)) - 3) Then   ' aqui habia list11
                                                    cont = Right(lista_users_shared.List(Y), 2) + 1
                                                    lista_users_shared.RemoveItem Y
                                                    lista_users_shared.AddItem user$ + " " + Format(cont, "00")
                                                    existe = 1
                                                    Exit For
                                        End If
                                      Next Y
       
                                      If existe = 0 Then
                                        lista_users_shared.AddItem user$ + " " + "01"
        
                                      End If
                                   
                                   
                                   
                                       Es_agente_comercial = False
                                   End If
                                End If
                                                           
                                GoTo brinca
                         End If
                         
                         If (user$ <> manager_user$) And (csr$ <> manager_csr$) And (user$ <> csr$) And user$ <> manager_commercial$ Then
                             total_user = total_user + 0.5
                             'total_csr2 = total_csr2 + 0.5
                         End If
                         
                                                  
                         
                         If user$ = manager_commercial$ Then  ' Or csr$ = manager_commercial$ Then
                             total_csr2 = total_csr2 + 1
                             existe = 0
                                cont = 0
                                For Y = 0 To List5.ListCount - 1
                                        If user$ = Left(List5.List(Y), Len(List5.List(Y)) - 3) Then   ' aqui habia user
                                                    cont = Right(List5.List(Y), 2) + 1
                                                    List5.RemoveItem Y
                                                    List5.AddItem user$ + " " + Format(cont, "00")
                                                    existe = 1
                                                    Exit For
                                        End If
                                Next Y
       
                                If existe = 0 Then
                                        List5.AddItem user$ + " " + "01"
        
                                End If
                                                           
                                GoTo brinca
                         
                         End If
                         
                         
                         
                  Else
                         ' SI NO SON DE LA MISMA OFICINA
                         If (user$ = manager_commercial$) And (csr$ = manager_commercial$) Then
                           If agente$ <> manager_commercial$ Then
                             total_csr2 = total_csr2 + 1
                             existe = 0
                                cont = 0
                                For Y = 0 To List5.ListCount - 1
                                        If user$ = Left(List5.List(Y), Len(List5.List(Y)) - 3) Then   ' aqui habia user
                                                    cont = Right(List5.List(Y), 2) + 1
                                                    List5.RemoveItem Y
                                                    List5.AddItem user$ + " " + Format(cont, "00")
                                                    existe = 1
                                                    Exit For
                                        End If
                                Next Y
       
                                If existe = 0 Then
                                        List5.AddItem user$ + " " + "01"
        
                                End If
                           Else
                              ' AGENTE es MANAGER-COMMERCIAL
                              existe = 0
                                cont = 0
                                
                                total_user = total_user + 0.5   ' Se agrego esta linea
                                
                                For Y = 0 To lista_users_shared.ListCount - 1 ' List11.ListCount - 1
                                        If user$ = Left(lista_users_shared.List(Y), Len(lista_users_shared.List(Y)) - 3) Then   ' aqui habia list11
                                                    cont = Right(lista_users_shared.List(Y), 2) + 1
                                                    lista_users_shared.RemoveItem Y
                                                    lista_users_shared.AddItem user$ + " " + Format(cont, "00")
                                                    existe = 1
                                                    Exit For
                                        End If
                                Next Y
       
                                If existe = 0 Then
                                        lista_users_shared.AddItem user$ + " " + "01"
        
                                End If
                              
                           End If
                                GoTo brinca
                         End If
                  
                         If user$ <> manager_user$ And csr$ <> manager_csr$ Then
                             total_user = total_user + 0.5
                             'total_csr2 = total_csr2 + 0.5
                         End If
                         
                         If (user$ = manager_user$) And (csr$ = manager_csr$) And csr$ <> manager_commercial$ Then
                             total_user = total_user + 0.5
                           '  total_csr2 = total_csr2 + 0.5
                         End If
                         
                         
                         If (user$ = manager_user$) And (csr$ = manager_csr$) And csr$ = manager_commercial$ Then
                             
                             List11.AddItem user$ + " 01"
                             GoTo brinca
                           
                         End If
                         
                         
                         If (user$ <> manager_user$) And (csr$ = manager_csr$) And csr$ = manager_commercial$ And SI_COMERCIAL = 1 Then
                             
                             List11.AddItem user$ + " 01"
                             GoTo brinca
                         
                         ElseIf (user$ <> manager_user$) And (csr$ = manager_csr$) And csr$ = manager_commercial$ And SI_COMERCIAL = 0 Then
                              ' se agrego 4/30
                                total_user = total_user + 0.5   ' Se agrego esta linea
                                
                                
                                existe = 0
                                For Y = 0 To lista_users_shared.ListCount - 1 '
                                        If user$ = Left(lista_users_shared.List(Y), Len(lista_users_shared.List(Y)) - 3) Then   ' aqui habia list11
                                                    cont = Right(lista_users_shared.List(Y), 2) + 1
                                                    lista_users_shared.RemoveItem Y
                                                    lista_users_shared.AddItem user$ + " " + Format(cont, "00")
                                                    existe = 1
                                                    Exit For
                                        End If
                                Next Y
       
                                If existe = 0 Then
                                        lista_users_shared.AddItem user$ + " " + "01"
        
                                End If
                                
                                
                                
                                
                                
                              
                           
                         End If
                         
    ' ************************************************************************************************************************************************************************
    ' ************************************************************************************************************************************************************************
    ' ************************************************************************************************************************************************************************
    
    
                       
                         
                         
                         
                         If ((user$ = manager_user$ And csr$ <> manager_csr$) Or (user$ <> manager_user$ And csr$ = manager_csr$)) Then ' And csr$ <> manager_commercial$ Then
                             If csr$ = manager_commercial$ Then
                               GoTo brinca
                             ElseIf user$ = manager_commercial$ Then
                               total_user = total_user + 1
                               
                                cont = 0
                                existe = 0
                                For Y = 0 To List5.ListCount - 1
                                        If user$ = Left(List5.List(Y), Len(List5.List(Y)) - 3) Then   ' aqui habia user
                                                    cont = Right(List5.List(Y), 2) + 1
                                                    List5.RemoveItem Y
                                                    List5.AddItem user$ + " " + Format(cont, "00")
                                                    existe = 1
                                                    Exit For
                                        End If
                                Next Y
       
                                If existe = 0 Then
                                        List5.AddItem user$ + " " + "01"
        
                                End If
                               
                                GoTo brinca
                               
                             Else
                               total_user = total_user + 0.5
                             End If
                           '  total_csr2 = total_csr2 + 0.5
                         Else
                            ' GoTo brinca
                         
                         End If
                         
    ' *************************************************************************************************************************************************************************
                        
                                                      
                  End If
                  
                  
       
       existe = 0
       cont = 0
       For Y = 0 To lista_users_shared.ListCount - 1
         If user$ = Left(lista_users_shared.List(Y), Len(lista_users_shared.List(Y)) - 3) Then   ' aqui habia user
             cont = Right(lista_users_shared.List(Y), 2) + 1
             lista_users_shared.RemoveItem Y
             lista_users_shared.AddItem user$ + " " + Format(cont, "00")
             existe = 1
             Exit For
          End If
       Next Y
       
       If existe = 0 Then
          
          lista_users_shared.AddItem user$ + " " + "01"
        
       End If
       
    
     GoTo brinca
   End If
  
  
  
  
   If agente$ = user$ And agente$ = csr$ Then
     
     
     ' a$ = recibo$
    '  r$ = recibo$
     
     If Es_Monterrey = 1 Then
            total_csr2 = total_csr2 + 1
            If agente$ = csr$ Then
              total_csr2 = total_csr2 - 1
            End If
            GoTo brinca
     End If
        
       
     If UCase(concepto2$) = "BF COMMERCIAL" Then
     
       total_csr = total_csr + 1
       GoTo brinca2
      
     End If
     
     
     If UCase(concepto2$) = "BF" Or UCase(concepto2$) = "BF CALL CENTER" Or UCase(concepto2$) = UCase("NewB - EFT To Company") Then
     
       total_csr = total_csr + 1
       GoTo brinca2
      
     End If
     
     
     'If UCase(concepto2$) = "BF COMMERCIAL" And Agent_Oficinacommercial$ = agente$ And agente$ <> csr$ Then  ' se anulo esta condicion 2/21
        
    '      total_csr = total_csr + 1
      
    ' End If
     
brinca2:

     
   End If
   
   
   
   
   
  
   
   
   
   
   
brinca:
   
Next t



conta_man = 0
For z = 0 To List5.ListCount - 1
    c = Right(List5.List(z), 3)
    conta_man = conta_man + c
Next z

lbltotal_users_managers.Caption = Format(conta_man, "#0")


GoTo salida



 Es_Monterrey = 0
 For Y = 0 To lista_managers.ListCount - 1
        a$ = RTrim(UCase(Left$(lista_managers.List(Y), 20)))
        b$ = LTrim(RTrim(Right(UCase(lista_managers.List(Y)), Len(lista_managers.List(Y)) - 20)))
                
        If a$ = agente$ Then
            If b$ = "MONTERREY" Then
               Es_Monterrey = 1
               Exit For
            End If
        End If
 Next Y







'List5.Clear
List6.Clear

' copia lista de ususarios compartidos a lista 6
For t = 0 To lista_users_shared.ListCount - 1
   List6.AddItem lista_users_shared.List(t)
Next t






' verifica que los usuarios no sean de  Monterrey y/o managers
check_again:

existe = 0
For t = 0 To lista_users_shared.ListCount - 1

  n$ = RTrim(Left(lista_users_shared.List(t), Len(lista_users_shared.List(t)) - 2))
  SI_MONTERREY = 0
  SI_COMMERCIAL = 0
  
' verifica si monterrey
     For Y = 0 To lista_managers.ListCount - 1
        a$ = RTrim(UCase(Left$(lista_managers.List(Y), 20)))
        b$ = LTrim(RTrim(Right(UCase(lista_managers.List(Y)), Len(lista_managers.List(Y)) - 20)))
                
        If a$ = n$ Then
        
          If b$ = "MONTERREY" Then
            SI_MONTERREY = 1
          End If
          
          If b$ = "COMMERCIAL" Then
            SI_COMMERCIAL = 1 'GoTo aqui_sigue
          End If
          
                    
           
             ' verifica oficinas a que pertenecen
                  For w = 1 To Val(ubicacion(0, 0))
                      If UCase(RTrim(ubicacion(w, 0))) = UCase(a$) Then
                          oficina_user$ = RTrim(ubicacion(w, 1))
                          If RTrim(ubicacion(w, 2)) = "" Then ubicacion(w, 2) = "None"
                          oficina_user2$ = RTrim(ubicacion(w, 2))
                      End If
                  
                      If UCase(RTrim(ubicacion(w, 0))) = UCase(n$) Then
                          oficina_csr$ = RTrim(ubicacion(w, 1))
                          If RTrim(ubicacion(w, 2)) = "" Then ubicacion(w, 2) = "Nada"
                          oficina_csr2$ = RTrim(ubicacion(w, 2))
                      End If
                  Next w
                  
                  
                  If oficina_user$ = "JA - COMMERCIAL" Then
                     GoTo continua_como
                  End If
                  
                  If oficina_user2$ = "None" And oficina_csr2$ = "None" Then
                     oficina_user2$ = "None2"
                  End If
                  
                   If oficina_user2$ = "Nada" And oficina_csr2$ = "Nada" Then
                     oficina_user2$ = "Nada2"
                  End If
                  
                  If oficina_user$ = oficina_csr$ Or oficina_user2$ = oficina_csr2$ Or oficina_user$ = oficina_csr2$ Or oficina_user2$ = oficina_csr$ Then
                  
                  Else
                     GoTo aqui_sigue
                  End If
                  
                  
continua_como:
                  
           
                ' verifica si existe en la lista 11
                existe_man = 0
                For z = 0 To List11.ListCount - 1
                  If lista_users_shared.List(t) = List11.List(z) Then
                    existe_man = 1
                    Exit For
                  End If
                Next z
                
                If existe_man = 0 Then
                   List11.AddItem lista_users_shared.List(t)
                End If
                
           'Else
           '     total_agent = total_agent - Val(Right(lista_users_shared.List(t), 2)) ' tenia 1
           'End If
            
            
           For z = 0 To List6.ListCount - 1
               If Left(List6.List(z), Len(a$)) = a$ Then ' lista_users_shared.List(t) Then
                   List6.RemoveItem z
                   existe = 1
                   Exit For
               End If
           Next z
            
             'existe = 1
            If List6.ListCount = 0 Then Exit For
              
          'End If
         
        End If
        
        
     Next Y
     

     If existe = 1 Then Exit For
aqui_sigue:
     
Next t


If existe = 1 Then
   ' remuevelo
  GoTo check_again
End If







' actualiza lista_users_shared

lista_users_shared.Clear
For t = 0 To List6.ListCount - 1
  lista_users_shared.AddItem List6.List(t)
Next t





salida:

lbltotal_csr.Caption = Format(total_csr, "00")


If lista_users_shared.ListCount = 0 Then
  lbltotal_user.Caption = "0"
Else
  lbltotal_user.Caption = Format(total_user, "#0.##")
End If

If Right(lbltotal_user.Caption, 1) = "." Then
   lbltotal_user.Caption = Left(lbltotal_user.Caption, Len(lbltotal_user.Caption) - 1)
End If






lblsuma_total_CSR.Caption = Format(total_csr, "#0.##")

If Right(lblsuma_total_CSR.Caption, 1) = "." Then
   lblsuma_total_CSR.Caption = Left(lblsuma_total_CSR.Caption, Len(lblsuma_total_CSR.Caption) - 1)
End If



  lblsuma_total_user.Caption = Format(total_user + total_csr2, "#0.##")
  
  ' se agrego esta linea para eliminar el error
If Val(lbltotal_user.Caption) = 0 Then
   lblsuma_total_user.Caption = "0"
End If


If Right(lblsuma_total_user.Caption, 1) = "." Then
   lblsuma_total_user.Caption = Left(lblsuma_total_user.Caption, Len(lblsuma_total_user.Caption) - 1)
End If


If List5.ListCount = 0 And lista_users_shared.ListCount = 0 Then
     total_user = 0
     total_csr2 = 0
Else
   total_puntos_de_managers = 0
   For Y = 0 To List5.ListCount - 1
      cantidad1 = Val(Right(List5.List(Y), 3))
      total_puntos_de_managers = total_puntos_de_managers + cantidad1
   Next Y
   
   If Val(lblsuma_total_user.Caption) = 0 Then
      lblsuma_total_user.Caption = Format(total_puntos_de_managers, "#0")
   End If
      
End If
   









If Right(lblsuma_total_user.Caption, 1) = "." Then
  lblsuma_total_user.Caption = Left(lblsuma_total_user.Caption, Len(lblsuma_total_user.Caption) - 1)
End If

t = total_csr + (total_user) + total_csr2
lbl_total_CSR_user.Caption = Str(t)


calcula_cantidades
End Sub

Public Sub calcula_cantidades()
On Error Resume Next

   
    

Dim gtotal As Double
List1.Clear
List2.Clear
List3.Clear
List9.Clear
List10.Clear

total_facturas_propias = 0
total_facturas_ajenas = 0

'List4.Clear
'List5.Clear

Dim Commercial$(50, 3)
Dim sSelect As String
Dim Rs As ADODB.Recordset
    
Set Rs = New ADODB.Recordset

contador_comercial = 0

Alta_Comercial = 0
par_encontrado = 0


For t = 1 To Grid2.Rows - 1
   
   
   NO_CAMBIAR = False
   agente_en_factura = 0
   csr$ = ""
   Grid2.col = 5
   Grid2.row = t
   csr$ = UCase(Grid2.Text)
  
   Grid2.col = 4
   user$ = UCase(Grid2.Text)
      
   invoice_3_users = 0
   
   ' == S T A R T ======================================================================
   
   csr2$ = ""
   Grid2.col = 10
   csr2$ = UCase(Grid2.Text)
   
   If csr2$ <> "" Then
       sSelect = "select username from EmployeeInfo where IDEmployee='" + csr2$ + "'"
       Rs.Open sSelect, base, adOpenUnspecified
       csr2$ = UCase(Rs(0))
       Rs.Close
   End If
      
      
   Grid2.col = 1
   recibo$ = Grid2.Text
     
   Grid2.col = 3
   concepto$ = Grid2.Text
   
   Grid2.col = 6
   cantidad$ = Grid2.Text
     
   Grid2.col = 12
   ID_Cliente$ = Grid2.Text
   
   
      
   If recibo$ = "267745" Then
          ' Stop
   End If
   
   
   
   
   If UCase(concepto$) = "BF" Or UCase(concepto$) = "BF COMMERCIAL" Then
      GoSub check_commercial
      
   End If
   
   
   
     
     
     Grid2.col = 9
     cantidad_pagada$ = Grid2.Text
     
     Grid2.col = 11
     balance_que_se_debe$ = Grid2.Text
     
     
   
    ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
     
    GoSub agentes_de_factura
   
   ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
   
    If (UCase$(LTrim(concepto$)) = "INVOICE") Then
    
          ' asigna agentes originales
           
          If agente$ <> csr_original$ And agente$ <> user_original$ Then
             agente_en_factura = 0
             GoTo brincalo
          Else
             agente_en_factura = 1
          End If
          
          
          If Val(cantidad$) > Val(cantidad_pagada$) And Val(cantidad_pagada$) > 0 And (UCase$(LTrim(concepto$)) = "INVOICE") And Val(Format(balance_que_se_debe$, "00000.00")) > 0 Then
             cantidad$ = cantidad_pagada$
             GoTo condic
          End If
          
          
          If Val(cantidad$) < Val(cantidad_pagada$) And Val(cantidad_pagada$) > 0 And (UCase$(LTrim(concepto$)) = "INVOICE") Then
            ' cantidad$ = cantidad_pagada$
             GoTo condic
          End If
          
          
          
          If Val(Format(cantidad$, "00000.00")) <> Val(Format(balance_que_se_debe$, "00000.00")) And balance_que_se_debe$ <> "" And (UCase$(LTrim(concepto$)) = "INVOICE") And Val(balance_que_se_debe$) > 0 Then
             cantidad$ = balance_que_se_debe$
          End If
           
             
condic:
           existe_invoice_commercial = False
           If Val(Format(balance_que_se_debe$, "00000.00")) = 0 And Val(cantidad_pagada$) >= Val(cantidad$) And Val(cantidad_pagada$) > 0 And (UCase$(LTrim(concepto$)) = "INVOICE") Then
             ' verifica si hay un late fee
                  pago_late_fee = 0
                  For p = 1 To Grid2.Rows - 1
                     Grid2.row = p
                     Grid2.col = 1
                     recibo2$ = Grid2.Text
     
                     Grid2.col = 3
                     concepto2$ = Grid2.Text
   
                     Grid2.col = 6
                     cantidad2$ = Grid2.Text
                     
                     If recibo2$ = recibo$ Then
                         If UCase$(concepto2$) = "LATE FEE" Then
                            pago_late_fee = Val(cantidad2$)
                            Exit For
                         End If
                         
                         
                         If UCase$(concepto2$) = "INVOICE COMMERCIAL" Then
                            existe_invoice_commercial = True
                            Exit For
                         End If
                     End If
                                                               
                  Next p
                  
                  s = Val(cantidad_pagada$) - pago_late_fee
                  
                                  
                                  
                  If existe_invoice_commercial = True Then
                     c = Val(cantidad$)
                     linea$ = Format(concepto$, "!@@@@@@@@@@@@@@@@@@@@") + Format(recibo$, "000000") + " " + Format(r$, "!@@@@@@@@@@@@@@@@@@@@") + " " + Format(Val(cantidad$), "000000.00")
                     List2.AddItem linea$
                     existe = 1
                     GoTo fin_bucle   ' comercial_saltado
                  
                  End If
                  
                  
                  
                  
                  If s <> Val(cantidad$) And existe_invoice_commercial = False Then
                           cantidad$ = cantidad_pagada$
                  End If
          End If
            
            
    End If
    
    
brincalo:
    
    ' }}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}}
    
           If (UCase$(LTrim(concepto$)) = "LATE FEE") And Val(balance_que_se_debe$) > 0 Then
    
         Total_factura_que_se_debia_antes = (Val(cantidad_pagada$) + Val(balance_que_se_debe$)) - Val(cantidad$)
         
         saldo = Total_factura_que_se_debia_antes - Val(cantidad_pagada$)
         
       
         If Val(cantidad_pagada$) >= Total_factura_que_se_debia_antes Then
             cantidad$ = Val(cantidad_pagada$) - Total_factura_que_se_debia_antes
         Else
             cantidad$ = "0"
         End If

    
    End If
            
            
            
            
            
         ' ---------------------------  DETECTA SI EL AGENTE FUE EL COBRADOR DEL CHEQUE ----------------------------------------
    
   If RTrim(LTrim(UCase(concepto$))) = "LATE FEE" Or RTrim(LTrim(UCase(concepto$))) = "INVOICE" Then
      
     agente_cobrador$ = ""
     cobrador$ = ""
     agente_csr0$ = ""
     agente_csr1$ = ""
     
 
     sSelect = "select idemployeeUSR, IdEmployeeCSR1, IdEmployeeCSR2  from [ReceiptsDTL] recdtl " & _
     "inner join  ReceiptsHDR  rechdr on rechdr.IDReceiptHDR =recdtl.IdReceiptHDR " & _
     "inner join InvoiceItemCatalog iitem on iitem.IdInvoiceItem=recdtl.IdInvoiceItem " & _
     "where rechdr.Active=1 and recdtl.[IdReceiptHDR] in (" + recibo$ + ") "
     
     ' and iitem.InvoiceItemName='Late Fee'"
    
     Rs.Open sSelect, base, adOpenUnspecified
     agente_cobrador$ = Rs(0)
     agente_csr0$ = Rs(1)
     agente_csr1$ = Rs(2)
     Rs.Close
     
     
     
    
    
     If agente_csr1$ = "" And agente_csr0$ = "" And agente_cobrador$ = "" Then
        r$ = user$
        GoTo go_ahead
     End If
     
     
        
     If agente_csr0$ <> "" And agente_cobrador$ <> "" And RTrim(LTrim(UCase(concepto$))) = "LATE FEE" Then
          
        sSelect = "select username from EmployeeInfo where IDEmployee='" + agente_cobrador$ + "'"
        Rs.Open sSelect, base, adOpenUnspecified
        cobrador$ = Rs(0)
        Rs.Close
            
       ' sSelect = "select username from EmployeeInfo where IDEmployee='" + agente_csr1$ + "'"
       ' Rs.Open sSelect, base, adOpenUnspecified
       ' csr$ = Rs(0)
       ' Rs.Close
                  
       ' user$ = cobrador$
        'csr$ = cobrador$
        
     ElseIf agente_csr1$ <> "" And agente_csr0$ <> "" And agente_cobrador$ <> "" And RTrim(LTrim(UCase(concepto$))) = "INVOICE" Then
   
        sSelect = "select username from EmployeeInfo where IDEmployee='" + agente_cobrador$ + "'"
        Rs.Open sSelect, base, adOpenUnspecified
        cobrador$ = Rs(0)
        Rs.Close
            
       ' sSelect = "select username from EmployeeInfo where IDEmployee='" + agente_csr0$ + "'"
       ' Rs.Open sSelect, base, adOpenUnspecified
       ' csr$ = Rs(0)
       ' Rs.Close
                  
        'user$ = cobrador$
               
      ElseIf agente_csr1$ <> "" And agente_csr0$ <> "" And agente_cobrador$ = "" And RTrim(LTrim(UCase(concepto$))) = "INVOICE" Then
   
        'sSelect = "select username from EmployeeInfo where IDEmployee='" + agente_csr0$ + "'"
        'Rs.Open sSelect, base, adOpenUnspecified
        'user$ = Rs(0)
        'Rs.Close
            
        'sSelect = "select username from EmployeeInfo where IDEmployee='" + agente_csr1$ + "'"
        'Rs.Open sSelect, base, adOpenUnspecified
        'csr$ = Rs(0)
        'Rs.Close
                  
      ElseIf agente_csr1$ = "" And agente_csr0$ <> "" And agente_cobrador$ <> "" And RTrim(LTrim(UCase(concepto$))) = "INVOICE" Then
   
        'sSelect = "select username from EmployeeInfo where IDEmployee='" + agente_cobrador$ + "'"
        'Rs.Open sSelect, base, adOpenUnspecified
        'user$ = Rs(0)
        'Rs.Close
            
        'sSelect = "select username from EmployeeInfo where IDEmployee='" + agente_csr0$ + "'"
        'Rs.Open sSelect, base, adOpenUnspecified
        'csr$ = Rs(0)
        'Rs.Close
        
        
     End If
   
   
   End If
  ' -------------------------------------------------------------------------------------------------------------------
   
   
   
            
     
     
    ' oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
   
     
     
     
   ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>
   

   
   
     If (UCase$(LTrim(concepto$)) = "INVOICE") Or (UCase$(LTrim(concepto$)) = "LATE FEE") Then
     
     
        ' r$ = user$
        ' GoTo go_ahead
     End If
   
   

   
        r$ = Left(user$, 2) + LCase(Right(user$, Len(user$) - 2))
                            
                           If UCase(RTrim(user$)) = agente$ And csr$ <> agente$ Then
                                r$ = Left(csr$, 2) + LCase(Right(csr$, Len(csr$) - 2))
                           ElseIf UCase(RTrim(user$)) = agente$ And csr$ = agente$ Then
                                r$ = Left(user$, 2) + LCase(Right(user$, Len(user$) - 2))
                           End If

go_ahead:
                            
        If Val(cantidad$) = 0 Then GoTo fin_bucle
                            
        If agente_en_factura = 1 And (UCase$(LTrim(concepto$)) = "INVOICE") Then   ' si agente esta en factura
        
           linea$ = Format(concepto$, "!@@@@@@@@@@@@@@@@@@@@") + Format(recibo$, "000000") + " " + Format(r$, "!@@@@@@@@@@@@@@@@@@@@") + " " + Format(cantidad$, "000000.00")
           List2.AddItem linea$
           
        ElseIf agente_en_factura = 0 And (UCase$(LTrim(concepto$)) = "INVOICE") Then
        
        Else
           If CSR2_Es_manager_commercial = 0 Then
               linea$ = Format(concepto$, "!@@@@@@@@@@@@@@@@@@@@") + Format(recibo$, "000000") + " " + Format(r$, "!@@@@@@@@@@@@@@@@@@@@") + " " + Format(cantidad$, "000000.00")
           Else
               If (UCase$(LTrim(concepto$)) = "BF COMMERCIAL") Then
                    concepto$ = concepto$ + "^"
                    linea$ = Format(concepto$, "!@@@@@@@@@@@@@@@@@@@@") + Format(recibo$, "000000") + " " + Format(r$, "!@@@@@@@@@@@@@@@@@@@@") + " " + Format(Val(cantidad$) / 2, "000000.00")
               ElseIf (UCase$(LTrim(concepto$)) = "BF") Then
                    linea$ = Format(concepto$, "!@@@@@@@@@@@@@@@@@@@@") + Format(recibo$, "000000") + " " + Format(r$, "!@@@@@@@@@@@@@@@@@@@@") + " " + Format(Val(cantidad$), "000000.00")
               Else
                    linea$ = Format(concepto$, "!@@@@@@@@@@@@@@@@@@@@") + Format(recibo$, "000000") + " " + Format(r$, "!@@@@@@@@@@@@@@@@@@@@") + " " + Format(Val(cantidad$), "000000.00")
               End If
           End If
           
           List2.AddItem linea$
        
        End If
     
   
   ' oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
   
  
fin_bucle:


Next t







grand_total = 0

For t = 0 To List2.ListCount - 1
  concepto$ = (RTrim(Left(List2.List(t), 20)))
  cantidad$ = Right(List2.List(t), 9)
  recibo$ = Mid$(List2.List(t), 21, 6)
  user$ = RTrim(Mid$(List2.List(t), 28, 20))
  cobrador$ = ""
  
  'If Val(cantidad$) = 118 Then Stop
  'If recibo$ = "267745" Then Stop
  
  If UCase(LTrim(RTrim(concepto$))) = "INVOICE" Then
  
   
     ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
     
    GoSub agentes_de_factura
   
   ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
   
   
       
  End If
        
       
    
  
  
  If List1.ListCount = 0 Then
           If RTrim(LTrim(UCase(concepto$))) = "INVOICE" Or RTrim(LTrim(UCase(concepto$))) = "LATE FEE" Then
 
                            r$ = UCase(Left(user$, 2)) + LCase(Right(user$, Len(user$) - 2))
                            
                            If UCase(RTrim(user$)) = agente$ And csr$ <> agente$ Then
                                r$ = UCase(Left(csr$, 2)) + LCase(Right(csr$, Len(csr$) - 2))
                            ElseIf UCase(RTrim(user$)) = agente$ And csr$ = agente$ Then
                                r$ = UCase(Left(user$, 2)) + LCase(Right(user$, Len(user$) - 2))
                            End If
                            
           Else
                            r$ = RTrim(UCase(Left(user$, 2)) + LCase(Right(user$, Len(user$) - 2)))
           End If
                        
  
         
  
        List1.AddItem Format(r$, "!@@@@@@@@@@@@@@@@@@@@") + " " + Format(concepto$, "!@@@@@@@@@@@@@@@@@@@@") + Space(1) + cantidad$
       
    
  Else
  
        encontrado = 0
        For Y = 0 To List1.ListCount - 1
               concep$ = RTrim(Mid$(List1.List(Y), 22, 20))
               user2$ = RTrim(Left(List1.List(Y), 20))
                           
                   
                If UCase(RTrim(concep$)) = UCase(RTrim(concepto$)) And LTrim(UCase(RTrim(user$))) = LTrim(UCase(RTrim(user2$))) Then
                                     
                        
                        If RTrim(LTrim(UCase(concepto$))) = "INVOICE" Or RTrim(LTrim(UCase(concepto$))) = "LATE FEE" Then
                               
                            If UCase(RTrim(user$)) = UCase(RTrim(agente$)) And UCase(RTrim(csr$)) <> UCase(RTrim(agente$)) Then
                               If UCase(RTrim(user2$)) = UCase(RTrim(csr$)) Then
                                   cant$ = Right(List1.List(Y), 9)
                                   Total = Val(cantidad$) + Val(cant$)
                                   List1.RemoveItem Y
                                   GoTo r1
                               
                               End If
                               
                             cant$ = cantidad$
                             Total = Val(cant$)
                             user$ = csr$
                             r$ = UCase(Left(user$, 2)) + LCase(Right(user$, Len(user$) - 2))
                             GoTo sigue_aqui
                             
                             
                            ElseIf UCase(RTrim(user$)) = UCase(RTrim(agente$)) And UCase(RTrim(csr$)) = UCase(RTrim(agente$)) Then
                             cant$ = Right(List1.List(Y), 9)
                             Total = Val(cantidad$) + Val(cant$)
                             List1.RemoveItem Y
                             
                            ElseIf UCase(RTrim(user$)) <> UCase(RTrim(agente$)) And UCase(RTrim(csr$)) = UCase(RTrim(agente$)) Then
                             cant$ = Right(List1.List(Y), 9)
                             Total = Val(cantidad$) + Val(cant$)
                             List1.RemoveItem Y
                             
                            ElseIf UCase(RTrim(user$)) <> UCase(RTrim(agente$)) And UCase(RTrim(csr$)) <> UCase(RTrim(agente$)) Then  ' se agrego 2/12
                             cant$ = Right(List1.List(Y), 9)
                             Total = Val(cantidad$) + Val(cant$)
                             List1.RemoveItem Y
                             
                            End If
 
r1:
                            r$ = UCase(Left(user$, 2)) + LCase(Right(user$, Len(user$) - 2))
                            
                            If UCase(RTrim(user$)) = UCase(RTrim(agente$)) And UCase(RTrim(csr$)) <> UCase(RTrim(agente$)) Then
                                r$ = UCase(Left(csr$, 2)) + LCase(Right(csr$, Len(csr$) - 2))
                            ElseIf UCase(RTrim(user$)) = UCase(RTrim(agente$)) And UCase(RTrim(csr$)) = UCase(RTrim(agente$)) Then
                                r$ = UCase(Left(user$, 2)) + LCase(Right(user$, Len(user$) - 2))
                            End If
                            
                        Else
                            cant$ = Right(List1.List(Y), 9)
                            Total = Val(cantidad$) + Val(cant$)
                            List1.RemoveItem Y
                        
                            r$ = UCase(Left(user$, 2)) + LCase(Right(user$, Len(user$) - 2))
                        End If
                        
                        
sigue_aqui:
                        
                        
                        
                        List1.AddItem Format(r$, "!@@@@@@@@@@@@@@@@@@@@") + " " + Format(concepto$, "!@@@@@@@@@@@@@@@@@@@@") + Space(1) + Format(Format(Total, "#####0.00"), "@@@@@@@@@")
                        
                        encontrado = 1
                        Exit For
               End If
         
        Next Y
   
   
        If encontrado = 0 Then
        
               r$ = UCase(Left(user$, 2)) + LCase(Right(user$, Len(user$) - 2))
               
               If RTrim(LTrim(UCase(concepto$))) = "INVOICE" Or RTrim(LTrim(UCase(concepto$))) = "LATE FEE" Then
 
 
                            If UCase(RTrim(user$)) = agente$ And csr$ <> agente$ Then
                                r$ = UCase(Left(csr$, 2)) + LCase(Right(csr$, Len(csr$) - 2))
                            ElseIf UCase(RTrim(user$)) = agente$ And csr$ = agente$ Then
                                r$ = UCase(Left(user$, 2)) + LCase(Right(user$, Len(user$) - 2))
                            End If
                            
               Else
                            r$ = RTrim(UCase(Left(user$, 2)) + LCase(Right(user$, Len(user$) - 2)))
                            
               End If
               
brinca_aqui:
                        
                         
              List1.AddItem Format(r$, "!@@@@@@@@@@@@@@@@@@@@") + " " + Format(concepto$, "!@@@@@@@@@@@@@@@@@@@@") + Space(1) + Format(Format(cantidad$, "#####0.00"), "@@@@@@@@@")
                       
        End If
  
  End If
  
saltado:
Next t







' ================================================================================================================================================================
' ================================================================================================================================================================
' ================================================================================================================================================================
' ================================================================================================================================================================
' ================================================================================================================================================================
' ================================================================================================================================================================
' ==================================  AQUI EMPIEZA EL CHEQUEO DE CADA CANTIDAD DEL AGENTE ========================================================================



If manager_commercial$ <> "" Then
   manager_Oficinacommercial$ = manager_commercial$
End If




For t = 0 To List1.ListCount - 1
    csr$ = RTrim(UCase(Left$(List1.List(t), 20)))
    concepto$ = RTrim(UCase(Mid$(List1.List(t), 22, 20)))
    cant_del_concepto = Val(Right(List1.List(t), 9))
    
    
    GoSub check_manager
    
    GoSub Asigna_oficinas
    
    
    
    existe = -1
    
     'If recibo$ = "267103" Then Stop
    ' If cant_del_concepto = 56.16 Then Stop
   '  If cant_del_concepto = 315 Then Stop
     
    'If UCase(concepto$) = "INVOICE" Then Stop
    
    If UCase(LTrim(csr$)) <> UCase(LTrim(manager_csr$)) Then  ' se agrego 12/12
       manager_csr$ = ""
    End If
    
    
    
    
    '==========================================================================================================================
    
     If RTrim(LTrim(UCase(concepto$))) = "LATE FEE" Then
    
          agente_cobrador$ = ""
          cobrador$ = ""
 
          sSelect = "select idemployeeUSR, IdEmployeeCSR1, IdEmployeeCSR2  from [ReceiptsDTL] recdtl " & _
          "inner join  ReceiptsHDR  rechdr on rechdr.IDReceiptHDR =recdtl.IdReceiptHDR " & _
          "inner join InvoiceItemCatalog iitem on iitem.IdInvoiceItem=recdtl.IdInvoiceItem " & _
          "where rechdr.Active=1 and recdtl.[IdReceiptHDR] in (" + recibo$ + ") and iitem.InvoiceItemName='Late Fee'"
    
          Rs.Open sSelect, base, adOpenUnspecified
          agente_cobrador$ = Rs(0)
          agente_csr0$ = Rs(1)
          agente_csr1$ = Rs(2)
          Rs.Close
        
          If agente_csr0$ <> "" And agente_cobrador$ <> "" Then   ' and agente_csr1$ <> ""
          
                 sSelect = "select username from EmployeeInfo where IDEmployee='" + agente_cobrador$ + "'"
                 Rs.Open sSelect, base, adOpenUnspecified
                 cobrador$ = UCase(Rs(0))
                 Rs.Close
            
                 sSelect = "select username from EmployeeInfo where IDEmployee='" + agente_csr1$ + "'"
                 Rs.Open sSelect, base, adOpenUnspecified
                 csr$ = UCase(Rs(0))
                 Rs.Close
                  
                 user$ = cobrador$
      
                 If UCase(cobrador$) <> UCase(agente$) Then
                     existe = 3
                     GoTo brinca
                 End If
   
          End If
      
     End If
    
    '==========================================================================================================================
    
    
    'If cant_del_concepto = 19 Then Stop
    
      
    
    If UCase$(LTrim(concepto$)) = "BF COMMERCIAL" Then
      
      
        ' verifica si CSR2 es manager comercial
        
        
      If manager_csr2$ = "" And manager_Oficinacommercial$ = "" Then
      
           encontrado = 0
           For Y = 0 To contador_comercial
               
             If Val(Commercial$(Y, 1)) = cant_del_concepto Then   ' and Commercial$(Y, 0) = csr$
               manager_csr2$ = Commercial$(Y, 2)
               Exit For
             End If
           
           Next Y
           
      End If
            
            
      If manager_csr2$ <> csr$ And manager_csr2$ <> agente$ And manager_csr2$ <> "" Then
         existe = 3
         GoTo brinca
         
      ElseIf manager_csr2$ <> csr$ And manager_csr2$ <> agente$ And manager_csr2$ = "" And manager_Oficinacommercial$ = "" Then
         'existe = 9
         'GoTo brinca
         
      ElseIf manager_Oficinacommercial$ = agente$ Then
         existe = 1
         GoTo brinca
         
      End If
     
     
    End If
    
    
    
' si es un invoice commercial

    If UCase$(LTrim(concepto$)) = "INVOICE COMMERCIAL" Then

      If Agent_Oficinacommercial$ = UCase(agente$) Then
            c = Val(cantidad$)
            existe = 1
            GoTo revisa
         
      ElseIf manager_commercial$ = UCase(agente$) Then
            c = Val(cantidad$)
            existe = 1
            GoTo revisa
      ElseIf manager_Oficinacommercial$ = UCase(agente$) Then
            c = Val(cantidad$)
            existe = 1
            GoTo revisa
            
            
      ElseIf Agent_Oficinacommercial$ = UCase(csr$) Then
            c = Val(cantidad$)
            existe = 3
            GoTo revisa
         
      ElseIf manager_commercial$ = UCase(csr$) Then
            c = Val(cantidad$)
            existe = 3
            GoTo revisa
      ElseIf manager_Oficinacommercial$ = UCase(csr$) Then
            c = Val(cantidad$)
            existe = 3
            GoTo revisa
            
            
      End If
      
      
      

    
    
    End If
    
    
    
    
    
    
    
    If UCase(csr$) = UCase(agente$) Then
             c = Val(Right(List1.List(t), 9))
    Else
  
    ' verifica si monterrey el AGENTE
             existe = -1
             For Y = 0 To lista_managers.ListCount - 1
                       a$ = RTrim(UCase(Left$(lista_managers.List(Y), 20)))
                       b$ = LTrim(RTrim(Right(UCase(lista_managers.List(Y)), Len(lista_managers.List(Y)) - 20)))
                
                       If a$ = agente$ Then
                               If b$ = "MONTERREY" Then
                                          existe = 5
                                          Exit For
                               End If
                       End If
        
        
             Next Y
  
             If existe = 5 Then GoTo brinca
  
             ' verifica si monterrey el CSR
             existe = -1
             For Y = 0 To lista_managers.ListCount - 1
                       a$ = RTrim(UCase(Left$(lista_managers.List(Y), 20)))
                       b$ = LTrim(RTrim(Right(UCase(lista_managers.List(Y)), Len(lista_managers.List(Y)) - 20)))
                
                       If a$ = csr$ Then
                                If b$ = "MONTERREY" Then
                                             existe = 4
                                             Exit For
                                End If
                       End If
        
        
             Next Y
     
             If existe = 4 Then GoTo brinca
  
  
  ' verifica si commercial el CSR
     
     For Y = 0 To lista_managers.ListCount - 1
        a$ = RTrim(UCase(Left$(lista_managers.List(Y), 20)))
        b$ = LTrim(RTrim(Right(UCase(lista_managers.List(Y)), Len(lista_managers.List(Y)) - 20)))
        existe = -1
        
        If b$ = "MANAGER_COMMERCIAL" Then
           If manager_Oficinacommercial$ <> a$ Then
              b$ = "MANAGER"
           End If
        End If
        
        If a$ = csr$ Then
          If b$ = "COMMERCIAL" Or b$ = "MANAGER_COMMERCIAL" Or b$ = "AGENT_COMMERCIAL" Then
          
          
                                  GoSub check_manager
                                  
          
                                  csr$ = RTrim(UCase(Left$(List1.List(t), 20)))
                                  csr_despues$ = RTrim(UCase(Left$(List1.List(t + 1), 20)))
                                  csr_2despues$ = RTrim(UCase(Left$(List1.List(t + 2), 20)))
                                  csr_antes$ = RTrim(UCase(Left$(List1.List(t - 1), 20)))
                                  csr_mas_antes$ = RTrim(UCase(Left$(List1.List(t - 2), 20)))
                                  
                                 
                                  
                                  a$ = RTrim(UCase(Left$(lista_managers.List(Y), 20)))
                                  b$ = LTrim(RTrim(Right(UCase(lista_managers.List(Y)), Len(lista_managers.List(Y)) - 20)))
                                  
                                  
                                   GoSub Asigna_oficinas
                                  
                                  cant_del_concepto_mas_antes = Val(Right(List1.List(t - 2), 9))
                                  cant_del_concepto_antes = Val(Right(List1.List(t - 1), 9))
                                  cant_del_concepto_despues = Val(Right(List1.List(t + 1), 9))
                                  cant_del_concepto_2despues = Val(Right(List1.List(t + 2), 9))
                                  cant_del_concepto = Val(Right(List1.List(t), 9))
                                  
          
                                  concepto_mas_antes$ = RTrim(UCase(Mid$(List1.List(t - 2), 22, 20)))
                                  concepto_antes$ = RTrim(UCase(Mid$(List1.List(t - 1), 22, 20)))
                                  concepto$ = RTrim(UCase(Mid$(List1.List(t), 22, 20)))
                                  concepto_despues$ = RTrim(UCase(Mid$(List1.List(t + 1), 22, 20)))
                                  concepto_2despues$ = RTrim(UCase(Mid$(List1.List(t + 2), 22, 20)))
                                  
                                  
                                 
                                 If csr$ = csr_antes$ Then
                                      If concepto_antes$ = "INVOICE COMMERCIAL" Then
                                        If concepto$ = "INVOICE" Then
                                            existe = 1    ' Todo
                                            GoTo revisa
                                        End If
                                        
                                      ElseIf concepto_despues$ = "INVOICE COMMERCIAL" Then
                                        If concepto$ = "INVOICE" Then
                                            existe = 1    ' Todo
                                            GoTo revisa
                                        End If
                                      End If
                                 End If
                                 
                                  
                                  
                                  
                                  
                                  
                                  ' verifica si antes es el mismo CSR, sino pasa a DESPUES del CSR
                                  If csr$ <> csr_antes$ Then
                                      If concepto_despues$ = "BF COMMERCIAL" Then
                                        If concepto$ = "BF" Then
                                                                           
                                        ' ---------------------------------
                                           If cant_del_concepto <> cant_del_concepto_despues Then
                                                 If csr$ = csr_2despues$ Then
                                                      If Format(cant_del_concepto_despues, "###0.0") = Format(cant_del_concepto_2despues, "###0.0") Then
                                                          ' NO ES DE LA POLIZA COMERCIAL
                                                          If oficina_user$ = oficina_csr$ Then  'Or oficina_USER$ = oficinaCSR2$ Or oficinaUSER2$ = oficina_csr$ Or oficinaUSER2$ = oficinaCSR2$ Then
                                                              If Left(b$, 7) = "MANAGER" Or b$ = "COMMERCIAL" Or b$ = "BF COMMERCIAL" Or b$ = "AGENT_COMMERCIAL" Then
                                                                existe = 1    ' Todo
                                                                Exit For
                                                              Else
                                                                existe = 9    ' compartido
                                                                Exit For
                                                              End If
                                                            
                                                          Else
                                                              existe = 9    ' compartido
                                                              Exit For
                                                          End If
                                                          
                                                      Else
                                                        
                                                          If (concepto_2despues$ = "BF COMMERCIAL^" And concepto_despues$ = "BF COMMERCIAL") Or (concepto_2despues$ = "BF COMMERCIAL" And concepto_despues$ = "BF COMMERCIAL^") Then
                                                              cantidad_BF_NO_commercial = (cant_del_concepto - cant_del_concepto_despues) / 2
                                                              c = cantidad_BF_NO_commercial + cant_del_concepto_despues
                                                              GoTo suma_la_cantidad
                                                          End If
                                                          
                                                          
                                                        
                                                          ' separa cantidad de BF
                                                          
                                                          If Left(b$, 7) = "MANAGER" Or b$ = "COMMERCIAL" Then
                                                             cantidad_BF_NO_commercial = (cant_del_concepto - cant_del_concepto_despues)
                                                          Else
                                                          
                                                             TOTAL_CONCEPTOS_BF = cant_del_concepto_2despues + cant_del_concepto
                                                             If cant_del_concepto_despues = TOTAL_CONCEPTOS_BF Then
                                                                  c = cant_del_concepto
                                                                  GoTo suma_la_cantidad
                                                             End If
                                                             
                                                             
                                                             cantidad_BF_NO_commercial = (cant_del_concepto - cant_del_concepto_despues) / 2
                                                             
                                                              ' separa NB
                                                             For Yy = 0 To List5.ListCount - 1
                                                               If csr$ = Left(List5.List(Yy), Len(List5.List(Yy)) - 3) Then
                                                                    cont = Right(List5.List(Yy), 2) - 1
                                                                    List5.RemoveItem Yy
                                                                    List5.AddItem csr$ + " " + Format(cont, "00")
                                                                   
                                                                    
                                                                    lista_users_shared.AddItem csr$ + " " + "01"
                                                                    
                                                                    lbltotal_users_managers.Caption = Format(Val(lbltotal_users_managers.Caption) - 1, "#0")
                                                                    lbltotal_user.Caption = Format(Val(lbltotal_user.Caption) + 0.5, "#0.0")
                                                                    
                                                                    lblsuma_total_user.Caption = Str(Val(lbltotal_users_managers.Caption) + Val(lbltotal_user.Caption))
                                                                    lbl_total_CSR_user.Caption = Str(Val(lblsuma_total_CSR.Caption) + Val(lblsuma_total_user.Caption))
                                                                    
                                                                    Exit For
                                                               End If
                                                             Next Yy
                                                          
                                                          End If
                                                          
                                                          
                                                          c = cantidad_BF_NO_commercial + cant_del_concepto_despues
                                                          
                                                         
                                                          
                                                          
                                                          
                                                          GoTo suma_la_cantidad
                                                      End If
                                                 End If
                                                 
                                                  
                                           ElseIf Format(cant_del_concepto, "###0.0") = Format(cant_del_concepto_despues, "###0.0") Then
                                                                                                                    
                                                          'If oficina_USER$ = oficina_csr$ Then
                                                          
                                                            If UCase(a$) = UCase(csr$) Then
                                                                  If Left(b$, 7) = "MANAGER" Or b$ = "COMMERCIAL" Or b$ = "BF COMMERCIAL" Or b$ = "AGENT_COMMERCIAL" Then
                                                                      existe = 1    ' Todo al user$
                                                                      Exit For
                                                                  Else
                                                                      existe = 9    ' compartido
                                                                      Exit For
                                                                  End If
                                                              
                                                            Else
                                                                                                                          
                                                                  If Left(b$, 7) = "MANAGER" Or b$ = "COMMERCIAL" Or b$ = "BF COMMERCIAL" Or b$ = "AGENT_COMMERCIAL" Then
                                                                      existe = 3    ' nada para agente
                                                                      Exit For
                                                                  Else
                                                                      existe = 9    ' compartido
                                                                      Exit For
                                                                  End If
                                                              
                                                            End If
                                                            
                                                         ' Else
                                                          
                                                          '    If UCase(a$) = UCase(csr$) Then
                                                          '        If Left(b$, 7) = "MANAGER" Or b$ = "COMMERCIAL" Or b$ = "BF COMMERCIAL" Or b$ = "AGENT_COMMERCIAL" Then
                                                          '            existe = 1    ' Todo al user$
                                                           '           Exit For
                                                           '       Else
                                                           '           existe = 9    ' compartido
                                                           '           Exit For
                                                           '       End If
                                                              
                                                            '  Else
                                                                                                                          
                                                            '      If Left(b$, 7) = "MANAGER" Or b$ = "COMMERCIAL" Or b$ = "BF COMMERCIAL" Or b$ = "AGENT_COMMERCIAL" Then
                                                            '          existe = 3    ' nada para agente
                                                            '          Exit For
                                                            '      Else
                                                            '          existe = 9    ' compartido
                                                             '         Exit For
                                                              '    End If
                                                              
                                                              'End If
                                                              
                                                              
                                                              
                                                          
                                                             ' If Right(b$, 10) = "COMMERCIAL" Then
                                                             '    existe = 3    ' nada porque es manager o agente cometrcial
                                                             '    Exit For
                                                             ' Else
                                                             '    existe = 9    ' compartido
                                                             '    Exit For
                                                             ' End If
                                               '           End If
                                           End If
                                        ' --------------------------------
                                        End If
                                      End If
                                                                                              
                                  
                                  End If
                                  
                                  
                                  
                                          
                                  
                                  
                                  
                                  si_es_poliza_commercial = 0
                                  
                                  If concepto_antes$ = "BF COMMERCIAL" And csr_antes$ = csr$ Then
                                       If Left(concepto_mas_antes$, 2) = "BF" And csr_mas_antes$ = csr$ And Format(cant_del_concepto_mas_antes, "0000.0") = Format(cant_del_concepto_antes, "0000.0") Then
                                          si_es_poliza_commercial = 0
                                          GoTo continue_right_here
                                       Else
                                          si_es_poliza_commercial = 1
                                       End If
                                    
                                  End If
                                  
                                  
                                  If concepto$ = "BF COMMERCIAL" And concepto_antes$ = "BF" And csr_antes$ = csr$ And b$ = "AGENT_COMMERCIAL" Then
                                    si_es_poliza_commercial = 1
                                  End If
                                  
                                  
                                  If concepto$ = "BF COMMERCIAL" And Left(UCase(concepto_despues$), 2) = "BF" And csr_despues$ = csr$ And b$ = "AGENT_COMMERCIAL" Then
                                    si_es_poliza_commercial = 1
                                  End If
                                  
                                  
                                  If concepto_despues$ = "BF COMMERCIAL" And csr_despues$ = csr$ Then
                                    si_es_poliza_commercial = 1
                                  End If
                                  
                                    
                                  If concepto_antes$ = "BF CALL CENTER" And csr_antes$ = csr$ Then
                                    si_es_poliza_commercial = 2
                                  End If
                                  
                                  If concepto_despues$ = "BF CALL CENTER" And csr_despues$ = csr$ Then
                                    si_es_poliza_commercial = 2
                                  End If
                                  
                                                               
                                  If concepto$ = "BF CALL CENTER" And concepto_antes$ = "BF" And csr_antes$ = csr$ Then
                                    si_es_poliza_commercial = 2
                                  End If
                                  
                                  
continue_right_here:
                                  
                                   ' verifica oficinas a que pertenecen
                                         For w = 1 To Val(ubicacion(0, 0))
                                              If UCase(RTrim(ubicacion(w, 0))) = UCase(agente$) Then
                                                      oficina_user$ = RTrim(ubicacion(w, 1))
                                                      oficina_user2$ = RTrim(ubicacion(w, 2))
                                              End If
                  
                                              If UCase(RTrim(ubicacion(w, 0))) = UCase(csr$) Then
                                                      oficina_csr$ = RTrim(ubicacion(w, 1))
                                                      oficina_csr2$ = RTrim(ubicacion(w, 2))
                                                                                                   
                                              End If
                                              
                                              ' csr_despues$
                                              
                                              
                                              If UCase(RTrim(ubicacion(w, 0))) = UCase(csr_despues$) Then
                                                      oficina_csr_next$ = RTrim(ubicacion(w, 1))
                                                      oficina_csr_next2$ = RTrim(ubicacion(w, 2))
                                                                                                   
                                              End If
                                              
                                              
                                              
                                         Next w
                                         
                                         
                                         
                                         If oficina_user2$ = "None" And oficina_csr2$ = "None" Then
                                                      oficina_user2$ = "Nada"
                                         End If
                  
                                         If oficina_user2$ = "" Then oficina_user2$ = " "
                                         
                                         If oficina_user2$ = "Nada" And oficina_csr2$ = "Nada" Then
                                                      oficina_user2$ = "None"
                                         End If
                                  
                                  
                                  
                                  
                                  
                                    If concepto_antes$ <> "BF COMMERCIAL" And concepto$ <> "BF COMMERCIAL" And concepto_despues$ <> "BF COMMERCIAL" And concepto_2despues$ <> "BF COMMERCIAL" Then
                                          ' NO ES DE LA POLIZA COMERCIAL
                                             If si_es_poliza_commercial = 0 Then
                                                   
                                                 If manager_agente$ = "" And manager_csr$ = "" Then  ' And manager_commercial$ = "" Then
                                                     
                                                    existe = 9   ' comparten
                                                    Exit For
                                                    
                                                 ElseIf manager_agente$ <> "" And manager_csr$ = "" Then
                                                                                                              
                                                    If (oficina_agente$ = oficina_csr$) Then
                                                    
                                                          existe = 3    ' cero para el agente
                                                          Exit For
                                                          
                                                    Else
                                                          existe = 9    ' comparten
                                                          Exit For
                                                    
                                                    End If
                                                    
                                                  ElseIf manager_agente$ = "" And manager_csr$ <> "" Then
                                                  
                                                    If (oficina_agente$ = oficina_csr$) Or (oficina_agente$ = oficina_csr2$) Then
                                                    
                                                          existe = 1    ' todo para el agente
                                                          Exit For
                                                          
                                                    Else
                                                          existe = 9    ' comparten
                                                          Exit For
                                                    
                                                    End If
                                                    
                                                    
                                                  ElseIf b$ = "AGENT_COMMERCIAL" And (a$ <> manager_agente$ And a$ <> manager_csr$) Then
                                                    
                                                    If (oficina_agente$ = oficina_csr$) Then
                                                    
                                                          existe = 3    ' cero para el agente
                                                          Exit For
                                                          
                                                    Else
                                                          existe = 9    ' comparten
                                                          Exit For
                                                    
                                                    End If
                                                    
                                                  Else
                                                  
                                                    
                                                        existe = 9  ' comparten
                                                        Exit For
                                                        
                                                    
                                                  End If
                                                    
                                             End If
                                  
                                    End If
                                         
                                         
                                         

                
                                  ' anula el status de agente comercial sobre manager gral comercial
                                  If manager_Oficinacommercial$ = agente$ Then
                                    si_es_poliza_commercial = 0
                                    existe = 3
                                    Exit For
                                  End If
                                  
                                  
                                  
                                  
                                      If concepto$ = "BF CALL CENTER" And si_es_poliza_commercial = 2 Then
                                             existe = 1
                                             Exit For
                                      End If
                                      
                                      
                                                                      
                                      If concepto$ = "BF COMMERCIAL" And si_es_poliza_commercial = 1 Then
                                             existe = 3
                                             Exit For
                                      End If
                                   
                                    
                                       
                                       
                                      If Left(concepto$, 2) = "BF" And concepto$ <> "BF COMMERCIAL" And concepto$ <> "BF CALL CENTER" And si_es_poliza_commercial = 1 Then
                                         
                                      
                                      
                                             existe = 1
                                             Exit For
                                      End If
                                  
                                  
                                      If concepto$ = "BF" And concepto_despues$ = "BF CALL CENTER" And si_es_poliza_commercial = 2 Then
                                             existe = 1
                                             Exit For
                                      End If
                                  
                                  
          
                                                                                  
                                         
                                         
                  
                                         If (oficina_user$ = oficina_csr$ Or oficina_user2$ = oficina_csr2$ Or oficina_user$ = oficina_csr2$ Or oficina_user2$ = oficina_csr$) Then
                                                    If concepto$ = "BF COMMERCIAL" Then
                                                      existe = 3
                                                      Exit For
                                                    Else
                                                    
                                                   
                                                    
                                                     ' es manager el agente?
                                                         es_manager = 0
                                                         For z = 0 To lista_managers.ListCount - 1
                                                             a$ = RTrim(UCase(Left$(lista_managers.List(z), 20)))
                                                             bb$ = LTrim(RTrim(Right(UCase(lista_managers.List(z)), Len(lista_managers.List(z)) - 20)))
                         
                                                             If a$ = csr$ Then
                                                                  If bb$ = "MANAGER" Or bb$ = "MANAGER_COMMERCIAL" Or bb$ = "COMMERCIAL" Then
                                                                          es_manager = 1
                                                                          Exit For
                                                                  End If
                                                             End If
                                                                
                                                             If a$ = agente$ Then
                                                                  If bb$ = "MANAGER" Or bb$ = "MANAGER_COMMERCIAL" Or bb$ = "COMMERCIAL" Then
                                                                          es_manager = 2
                                                                          Exit For
                                                                  End If
                                                             End If
                                                                
                                                         Next z
                                                         
                                                         If es_manager = 1 Then
                                                            existe = 1
                                                            
                                                         ElseIf es_manager = 2 Then
                                                            existe = 3
                                                         Else
                                                            existe = 9
                                                         End If
                                                         
                                                         
                                                         
                                                    
                                                    
                                                      'existe = 1    ' tenia existe=1   8/17/2022
                                                      Exit For
                                                    End If
                                         Else
                                                     existe = 9
                                                      Exit For
                  
                                         End If
                  
                                         
            
                                         existe = 7
                                         Exit For
                                         
                                         
                     
          ElseIf b$ = concepto$ And b$ = "BF COMMERCIAL" Then
                                  
                                            existe = 3
                                            Exit For
              
          ElseIf b$ = concepto$ And b$ = "BF" Then
                                            existe = 1
                                            Exit For
                
            
          End If
         
        End If
        
       ' If existe = -1 Then
       '   existe = 10
       '   Exit For
       ' End If
        
        
     Next Y
     
     
salto_temp:
     
     If existe = 1 Or existe = 9 Or existe = 11 Or existe = 3 Then GoTo brinca
     
     
     ' verifica si es CALL CENTER
     
     ' ///////////////////////////////////////////////////////////////////////////////////////////
     
      concepto_antes$ = RTrim(UCase(Mid$(List1.List(t - 1), 22, 20)))
      concepto$ = RTrim(UCase(Mid$(List1.List(t), 22, 20)))
      concepto_despues$ = RTrim(UCase(Mid$(List1.List(t + 1), 22, 20)))
      concepto_2despues$ = RTrim(UCase(Mid$(List1.List(t + 2), 22, 20)))
      
      
      csr$ = RTrim(UCase(Left$(List1.List(t), 20)))
      csr_despues$ = RTrim(UCase(Left$(List1.List(t + 1), 20)))
      csr_2despues$ = RTrim(UCase(Left$(List1.List(t + 2), 20)))
      csr_antes$ = RTrim(UCase(Left$(List1.List(t - 1), 20)))
      csr_mas_antes$ = RTrim(UCase(Left$(List1.List(t - 2), 20)))
                                  
                                  
     
     If concepto_despues$ = "BF CALL CENTER" And concepto$ = "BF" And csr_despues$ = csr$ Then
                   cant_del_concepto = Val(Right(List1.List(t), 9))
                   c = Val(Right(List1.List(t), 9))
                   existe = 1
                   GoTo revisa
                   
     End If
     
     
     If concepto_antes$ = "BF" And concepto$ = "BF CALL CENTER" And csr_despues$ = csr$ Then
                   cant_del_concepto = Val(Right(List1.List(t), 9))
                   c = Val(Right(List1.List(t), 9))
                   existe = 3
                   GoTo revisa
                   
     End If
     
     
     
     
     
     If concepto_despues$ = "BF COMMERCIAL" And concepto$ = "BF" And csr_despues$ = csr$ Then
                  ' c = Val(Right(List1.List(t), 9))
                  ' existe = 1
                  ' GoTo revisa
                   
     End If
     
     
     If concepto_antes$ = "BF" And concepto$ = "BF COMMERCIAL" And csr_despues$ = csr$ Then
                 '  c = Val(Right(List1.List(t), 9))
                 '  existe = 3
                 '  GoTo revisa
                   
     End If
     
     
     
     ' //////////////////////////////////////////////////////////////////////////////////////////
     
     
     
     
     
     
     
     
     
     
  
  ' verifica si es commercial el AGENTE
     'existe = 0
             For Y = 0 To lista_managers.ListCount - 1
                           a$ = RTrim(UCase(Left$(lista_managers.List(Y), 20)))
                           b$ = LTrim(RTrim(Right(UCase(lista_managers.List(Y)), Len(lista_managers.List(Y)) - 20)))
                
                           If a$ = agente$ Then   ' csmbie csr$ por agente$
                                  If b$ = "COMMERCIAL" Or b$ = "MANAGER_COMMERCIAL" Or b$ = "AGENT_COMMERCIAL" Or b$ = "BF COMMERCIAL" Then
                                  
                                  GoSub check_manager
                                  
                                  'a$ = RTrim(UCase(Left$(lista_managers.List(Y), 20)))
                                  'b$ = LTrim(RTrim(Right(UCase(lista_managers.List(Y)), Len(lista_managers.List(Y)) - 20)))
                                   
                                   
                                  
                                 GoSub Asigna_oficinas
                                  
                                  csr$ = RTrim(UCase(Left$(List1.List(t), 20)))
                                  csr_despues$ = RTrim(UCase(Left$(List1.List(t + 1), 20)))
                                  csr_2despues$ = RTrim(UCase(Left$(List1.List(t + 2), 20)))
                                  csr_antes$ = RTrim(UCase(Left$(List1.List(t - 1), 20)))
                                  
                                  
                                  
                                  cant_del_concepto_antes = Val(Right(List1.List(t - 1), 9))
                                  cant_del_concepto_despues = Val(Right(List1.List(t + 1), 9))
                                  cant_del_concepto_2despues = Val(Right(List1.List(t + 2), 9))
                                  cant_del_concepto = Val(Right(List1.List(t), 9))
                                  
          
                                  concepto_antes$ = RTrim(UCase(Mid$(List1.List(t - 1), 22, 20)))
                                  concepto$ = RTrim(UCase(Mid$(List1.List(t), 22, 20)))
                                  concepto_despues$ = RTrim(UCase(Mid$(List1.List(t + 1), 22, 20)))
                                  concepto_2despues$ = RTrim(UCase(Mid$(List1.List(t + 2), 22, 20)))
                                  
                                  
                                  
                                  If csr$ = csr_antes$ Then
                                      If concepto_antes$ = "INVOICE COMMERCIAL" Then
                                        If concepto$ = "INVOICE" Then
                                            existe = 3    ' nada
                                            GoTo revisa
                                        End If
                                        
                                      ElseIf concepto_despues$ = "INVOICE COMMERCIAL" Then
                                        If concepto$ = "INVOICE" Then
                                            existe = 3    ' nada
                                            GoTo revisa
                                        End If
                                      End If
                                 End If
                                  
                                  
                                  
                                  ' verifica si antes es el mismo CSR, sino pasa a DESPUES del CSR
                                  If csr$ <> csr_antes$ Then
                                      
                                      If concepto$ = "BF COMMERCIAL^" Then
                                         c = cant_del_concepto
                                         GoTo comercial_saltado
                                      End If
                                  
                                  
                                      If concepto_despues$ = "BF COMMERCIAL" Or concepto_despues$ = "BF COMMERCIAL^" Then
                                         If concepto$ = "BF" Then
                                         ' ----------------------------------------------
                                           If Format(cant_del_concepto, "###0.0") <> Format(cant_del_concepto_despues, "###0.0") Then
                                                 If csr$ = csr_2despues$ Or csr_2despues$ = "" Then
                                                      If Format(cant_del_concepto_despues, "###0.0") = Format(cant_del_concepto_2despues, "###0.0") Then
                                                          ' NO ES DE LA POLIZA COMERCIAL, verifica si es manager la otra persona
                                                          
                                                          If oficina_user$ = oficina_csr$ Then ' Or oficinaUSER1$ = oficinaCSR2$ Or oficinaUSER2$ = oficinaCSR1$ Or oficinaUSER2$ = oficinaCSR2$ Then
                                                              
                                                              If Left(b$, 7) = "MANAGER" Or b$ = "COMMERCIAL" Or b$ = "BF COMMERCIAL" Or b$ = "AGENT_COMMERCIAL" Then
                                                                existe = 3    ' nada porque es manager del agente
                                                                Exit For
                                                              Else
                                                                existe = 9    ' compartido
                                                                Exit For
                                                              End If
                                                            
                                                          Else
                                                              existe = 9    ' compartido
                                                              Exit For
                                                          End If
                                                          
                                                      Else
                                                        
                                                          ' separa cantidad de BF
                                                          If Left(b$, 7) = "MANAGER" Or b$ = "COMMERCIAL" Or b$ = "AGENT_COMMERCIAL" Then
                                                             c = 0
                                                             GoTo suma_la_cantidad
                                                             
                                                          Else
                                                             cantidad_BF_NO_commercial = (cant_del_concepto - cant_del_concepto_despues) / 2
                                                             
                                                              TOTAL_CONCEPTOS_BF = cant_del_concepto_2despues + cant_del_concepto
                                                             If cant_del_concepto_despues = TOTAL_CONCEPTOS_BF Then
                                                                  c = cant_del_concepto
                                                                  GoTo suma_la_cantidad
                                                             End If
                                                             
                                                             
                                                              ' separa NB
                                                             For Yy = 0 To List5.ListCount - 1
                                                               If csr$ = Left(List5.List(Yy), Len(List5.List(Yy)) - 3) Then
                                                                    cont = Right(List5.List(Yy), 2) - 1
                                                                    List5.RemoveItem Yy
                                                                    List5.AddItem csr$ + " " + Format(cont, "00")
                                                                   
                                                                    
                                                                    lista_users_shared.AddItem csr$ + " " + "01"
                                                                    
                                                                    lbltotal_users_managers.Caption = Format(Val(lbltotal_users_managers.Caption) - 1, "#0")
                                                                    lbltotal_user.Caption = Format(Val(lbltotal_user.Caption) + 0.5, "#0.0")
                                                                    
                                                                    lblsuma_total_user.Caption = Str(Val(lbltotal_users_managers.Caption) + Val(lbltotal_user.Caption))
                                                                    lbl_total_CSR_user.Caption = Str(Val(lblsuma_total_CSR.Caption) + Val(lblsuma_total_user.Caption))
                                                                    
                                                                    
                                                                    Exit For
                                                               End If
                                                             Next Yy
                                                          
                                                          
                                                          End If
                                                          
                                                          
                                                          c = cantidad_BF_NO_commercial + cant_del_concepto_despues
                                                          
                                                           
                                                          
                                                          GoTo suma_la_cantidad
                                                          
                                                      End If
                                                 End If
                                           ElseIf Format(cant_del_concepto, "###0.0") = Format(cant_del_concepto_despues, "###0.0") Then
                                                          
                                                           'If oficina_USER$ = oficina_csr$ Then
                                                          
                                                            If UCase(a$) = UCase(csr$) Then
                                                                  If Left(b$, 7) = "MANAGER" Or b$ = "COMMERCIAL" Or b$ = "BF COMMERCIAL" Or b$ = "AGENT_COMMERCIAL" Then
                                                                      existe = 1    ' Todo al user$
                                                                      Exit For
                                                                  Else
                                                                      existe = 9    ' compartido
                                                                      Exit For
                                                                  End If
                                                              
                                                            Else
                                                                                                                          
                                                                  If Left(b$, 7) = "MANAGER" Or b$ = "COMMERCIAL" Or b$ = "BF COMMERCIAL" Or b$ = "AGENT_COMMERCIAL" Then
                                                                      existe = 3    ' nada para agente
                                                                      Exit For
                                                                  Else
                                                                      existe = 9    ' compartido
                                                                      Exit For
                                                                  End If
                                                              
                                                            End If
                                                            
                                                                                                                  
                                           ElseIf Val(Format(cant_del_concepto, "###0.0")) = (Val(Format(cant_del_concepto_despues, "###0.0")) + (Val(Format(cant_del_concepto, "###0.0")) - Val(Format(cant_del_concepto_despues, "###0.0")))) And Right(concepto$, 1) = "^" Then
                                                          
                                                           If UCase(a$) = UCase(csr$) Then
                                                                  If Left(b$, 7) = "MANAGER" Or b$ = "COMMERCIAL" Or b$ = "BF COMMERCIAL" Or b$ = "AGENT_COMMERCIAL" Then
                                                                      existe = 1    ' Todo al user$
                                                                      Exit For
                                                                  Else
                                                                      existe = 9    ' compartido
                                                                      Exit For
                                                                  End If
                                                              
                                                            Else
                                                                                                                          
                                                                  If Left(b$, 7) = "MANAGER" Or b$ = "COMMERCIAL" Or b$ = "BF COMMERCIAL" Or b$ = "AGENT_COMMERCIAL" Then
                                                                      existe = 3    ' nada para agente
                                                                      Exit For
                                                                  Else
                                                                      existe = 9    ' compartido
                                                                      Exit For
                                                                  End If
                                                              
                                                            End If
                                                          
                                                          
                                           End If
                                         ' -----------------------------------------------------
                                        
                                         
                                         
                                         End If
                                      End If
                                                                                              
                                  
                                  
                                  ElseIf csr$ = csr_antes$ Then
                                  '  si son el mismo csr
                                  
                                      If concepto$ = "BF COMMERCIAL^" Then
                                         c = cant_del_concepto
                                         GoTo comercial_saltado
                                      End If
                                      
                                  
                                      If concepto_antes$ = "BF COMMERCIAL" Then
                                         If concepto$ <> "BF" Then
                                         ' ----------------------------------------------
                                           If Format(cant_del_concepto, "###0.0") <> Format(cant_del_concepto_antes, "###0.0") Then
                                                 If csr$ = csr_antes$ Then
                                                                                                              
                                                          ' separa cantidad de BF
                                                          If oficina_agente$ = oficina_csr$ Then
                                                            If Left(concepto$, 2) = "BF" Then
                                                               If Left(b$, 7) = "MANAGER" Or b$ = "COMMERCIAL" Or b$ = "AGENT_COMMERCIAL" Then
                                                             
                                                                   c = 0
                                                                   GoTo suma_la_cantidad
                                                               Else
                                                                   c = cant_del_concepto / 2
                                                                   GoTo suma_la_cantidad
                                                               End If
                                                               
                                                             Else
                                                                   c = cant_del_concepto / 2
                                                                   GoTo suma_la_cantidad
                                                             
                                                             End If
                                                             
                                                          Else
                                                                c = cant_del_concepto / 2
                                                                GoTo suma_la_cantidad
                                                          
                                                          End If
                                                          
                                                      
                                                 End If
                                           'ElseIf Format(cant_del_concepto, "###0.0") = Format(cant_del_concepto_despues, "###0.0") Then
                                                          
                                                          
                                                          
                                                          
                                           End If
                                         ' -----------------------------------------------------
                                        
                                         
                                         
                                         End If
                                      End If
                                                                                              
                                  
                                  
                                  
                                  
                                  End If   '  cierre del if
                                  
                                  
                                  
                                  
                                  si_es_poliza_commercial = 0
                                  
                                  
                                 
                                  
                                  
                                  If List1.List(t + 1) <> "" Then
                                    cant_despues_concepto = Val(RTrim(UCase(Right$(List1.List(t + 1), 10))))
                                  End If
                                  
                                  
                                  
                                  If concepto_antes$ = "BF COMMERCIAL" And csr_antes$ = csr$ Then
                                    si_es_poliza_commercial = 1
                                  End If
                                  
                                  If concepto_despues$ = "BF COMMERCIAL" And csr_despues$ = csr$ Then
                                    si_es_poliza_commercial = 1
                                  End If
                                  
                                  
                                  If concepto$ = "BF COMMERCIAL" And csr_despues$ = "" Then
                                    si_es_poliza_commercial = 1
                                  End If
                                  
                                  If concepto$ = "BF COMMERCIAL" Then
                                    si_es_poliza_commercial = 1
                                  End If
                                  
                                  If (concepto_despues$ = "BF CALL CENTER" And csr_despues$ <> csr$) And (concepto$ = "BF CALL CENTER" And csr_antes$ = csr$) Then
                                    si_es_poliza_commercial = 3
                                    GoTo brinca_caso
                                  End If
                                  
                                  
                                  
                                  
                                  
                                  
                                  If concepto_antes$ = "BF CALL CENTER" And csr_antes$ = csr$ Then
                                    si_es_poliza_commercial = 2
                                  End If
                                  
                                  If concepto_despues$ = "BF CALL CENTER" And csr_despues$ = csr$ Then
                                    si_es_poliza_commercial = 2
                                  End If
                                  
                                  If concepto$ = "BF CALL CENTER" And csr_despues$ = csr$ Then
                                    si_es_poliza_commercial = 2
                                  End If
                                    
                                  If concepto$ = "BF CALL CENTER" And csr_antes$ = csr$ Then
                                    si_es_poliza_commercial = 2
                                  End If
                                  
                                  
                                                                   
                                  
brinca_caso:
                                     
                                     
                                      If concepto$ = "BF" And si_es_poliza_commercial = 2 And concepto_despues$ = "BF CALL CENTER" And csr_despues$ = csr$ Then
                                             existe = 3  '
                                             Exit For
                                      End If
                                     
                                     
                                      If concepto$ = "BF CALL CENTER" And si_es_poliza_commercial = 2 And concepto_antes$ = "BF" And csr_antes$ = csr$ Then
                                             existe = 1  '
                                             Exit For
                                      End If
                                      
                                      
                                      
                                     
                                   
                                      If concepto$ = "BF COMMERCIAL" And si_es_poliza_commercial = 1 Then
                                             existe = 1
                                             Exit For
                                      End If
                                   
                                                                                                           
                                      If concepto$ = "BF CALL CENTER" And si_es_poliza_commercial = 2 Then
                                             existe = 3
                                             Exit For
                                      End If
                                      
                                      
                                      If concepto$ = "BF CALL CENTER" And si_es_poliza_commercial = 3 Then
                                             existe = 9
                                             Exit For
                                      End If
                                      
                                      
                                      
                                      
                                      
                                       
                                       
                                     ' If concepto$ = "BF ENDO FEE" And si_es_poliza_commercial = 1 Then
                                     '        existe = 1
                                     '        Exit For
                                     ' End If
                                      
                                      
                                       
                                      If Left(concepto$, 2) = "BF" And concepto$ <> "BF COMMERCIAL" And concepto$ <> "BF CALL CENTER" And si_es_poliza_commercial = 1 Then
                                             
                                             
                                             
                                             
                                             
                                             
                                             If b$ = "AGENT_COMMERCIAL" Or b$ = "COMMERCIAL" Then
                                                
                                                If (concepto$ = "BF" Or concepto$ = "BF ENDO FEE" Or concepto$ = "BF PAYMENT FEE") And (concepto_despues$ = "BF COMMERCIAL" Or concepto_antes$ = "BF COMMERCIAL") Then
                                                
                                                      ' If cant_del_concepto <> cant_despues_concepto Then
                                                      
                                                          ' verifica si la cantidad del BF es parte de un BF COMMERCIAL
                                                          ' ***************************************************************************************
                                                                     
                                                                     cont_veces = 0
                                                                     
Obten_BF:
                                                                     Set Rs = New ADODB.Recordset

                                                                     
                                                                     If cont_veces = 0 Then
                                                                       r$ = "and iitem.InvoiceItemName='BF COMMERCIAL' "
                                                                     Else
                                                                       r$ = ""
                                                                     End If
                                                                     
                                                                     sSelect = "SELECT " & _
                                                                     "recdtl.[IdReceiptHDR] as [Receipt #] " & _
                                                                     ",rechdr.Date " & _
                                                                     ",iitem.InvoiceItemName as [Invoice Item] " & _
                                                                     ",emp.Username as [User] " & _
                                                                     ",csr.Username as [CSR] " & _
                                                                     ",recdtl.Amount " & _
                                                                     ",ofc.Office " & _
                                                                     ",rechdr.IdOffice, rechdr.AmountPaid, rechdr.IdEmployeeCSR2, rechdr.BalanceDue, rechdr.IdCustomer  " & _
                                                                     "FROM  [ReceiptsDTL] recdtl " & _
                                                                     "inner join  ReceiptsHDR  rechdr on recdtl.IdReceiptHDR=rechdr.IDReceiptHDR " & _
                                                                     "inner join PoliciesHDR polhdr on polhdr.IdPoliciesHDR=rechdr.IdPoliciesHDR " & _
                                                                     "inner join EmployeeInfo emp on emp.IDEmployee=rechdr.IdEmployeeUSR " & _
                                                                     "inner join EmployeeInfo csr on csr.IDEmployee=rechdr.IdEmployeeCSR1 " & _
                                                                     "inner join OfficesCatalog ofc on ofc.IdOffice=rechdr.IdOffice " & _
                                                                     "inner join OfficesCatalog ofcpol on ofcpol.IdOffice=rechdr.IdOffice " & _
                                                                     "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
                                                                     "inner join Customers cus on cus.IdCustomer=polhdr.IdCustomer " & _
                                                                     "inner join InvoiceItemCatalog iitem on iitem.IdInvoiceItem=recdtl.IdInvoiceItem " & _
                                                                     "inner join BankAccountCatalog bank on bank.IdBankAccount=iitem.IdBankAccount " & _
                                                                     "inner join CompanysAGICatalog ja on ja.IdCompanyAGI=polhdr.IdCompanyAGI " & _
                                                                     "where iitem.IdInvoiceItem in (2,35,3,11,15,22,23,33,9,30,33,17,13,31,32, 36, 37,39,42,43) " & _
                                                                     "and cast(rechdr.Date as Date) >= '" + txtdatefrom.Text + "' AND cast( rechdr.DATE as Date) <= '" + txtdateto.Text + "' " & _
                                                                     "and rechdr.Active=1 " + r$ + " and (csr.username='" + csr$ + "' or emp.Username='" + csr$ + "') " & _
                                                                     "order by [Receipt #], rechdr.IdCustomer"
        
    
                                                                     ' ---------------------------------------------------------------------------
    
                                                                     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
                                                                     Rs.Open sSelect, base, adOpenStatic, adLockOptimistic
    
                                                                     Rs.MoveLast

                                                                     Rs.MoveFirst
                                                                     ' Assuming that rs is your ADO recordset
                                                                     grid.Rows = Rs.RecordCount + 1

                                                                     rsVar = Rs.GetString(adClipString, Rs.RecordCount)

                                                                     grid.cols = Rs.Fields.Count + 1
    
                                                                     grid.TextMatrix(0, 0) = ""
                                                                     ' Set column names in the grid
                                                                     For i = 0 To Rs.Fields.Count - 1
                                                                         grid.TextMatrix(0, i + 1) = Rs.Fields(i).name
                                                                     Next

                                                                     grid.row = 1
                                                                     grid.col = 1

                                                                     ' Set range of cells in the grid
                                                                     grid.RowSel = grid.Rows - 1
                                                                     grid.ColSel = grid.cols - 1
                                                                     grid.clip = rsVar

                                                                     ' Reset the grid's selected range of cells
                                                                     grid.RowSel = grid.row
                                                                     grid.ColSel = grid.col

                                                                     Rs.Close

                                                                     Set Rs = Nothing
                                                                       
                                                                                                                                                                                                             
                                                                     encontrado1 = 0
                                                                     For w = 1 To grid.Rows - 1
                                                                         grid.row = w
                                                                         grid.col = 6
                                                                         cantidad_comercial = Val(grid.Text)
                                                                                
                                                                         If Format(cantidad_comercial, "###0.0") = Format(cant_del_concepto, "###0.0") Or Format((cantidad_comercial + 0.01), "###0.0") = Format(cant_del_concepto, "###0.0") Then
                                                                            encontrado1 = 1
                                                                            Exit For
                                                                         End If
                                                                     Next w
                                                                     
                                                                     If encontrado1 = 1 Then
                                                                        existe = 3
                                                          
                                                                     End If
                                                      
                                                           
                                                           'c = (cant_del_concepto - cant_despues_concepto) / 2
                                                           'GoTo suma_la_cantidad
                                                           
                                                           
                                                      'Else
                                                      '     existe = 3
                                                      'End If
                                                'ElseIf concepto$ = "BF ENDO FEE" And (concepto_antes$ = "BF COMMERCIAL" Or concepto_despues$ = "BF COMMERCIAL") Then
                                                '      existe = 3   '  cero
                                                      
                                                'ElseIf concepto$ = "BF PAYMENT FEE" And (concepto_antes$ = "BF COMMERCIAL" Or concepto_despues$ = "BF COMMERCIAL") Then
                                                '      existe = 3  ' cero
                                                      
                                                Else
                                                      existe = 9   '  mitad     tenia existe=9  9/14/2022
                                                End If
                                                
                                                
                                                
                                               
                                             Else
                                                      existe = 3    '   cero    tenia existe=1   8/17/2022
                                                      
                                             End If
                                             
                                             
                                             'existe = 3   ' existe estaba como 3   8/17/2022
                                             Exit For
                                      End If
                                       
                                                                    
                                   
                                   
                                   
                                            ' verifica oficinas a que pertenecen
                                         For w = 1 To Val(ubicacion(0, 0))
                                              If UCase(RTrim(ubicacion(w, 0))) = UCase(agente$) Then
                                                      oficina_user$ = RTrim(ubicacion(w, 1))
                                                      oficina_user2$ = RTrim(ubicacion(w, 2))
                                              End If
                  
                                              If UCase(RTrim(ubicacion(w, 0))) = UCase(csr$) Then
                                                      oficina_csr$ = RTrim(ubicacion(w, 1))
                                                      oficina_csr2$ = RTrim(ubicacion(w, 2))
                                              End If
                                         Next w
                                         
                                         
                                         
                                         If oficina_user2$ = "None" And oficina_csr2$ = "None" Then
                                                      oficina_user2$ = "Nada"
                                         End If
                  
                                         If oficina_user2$ = "" Then oficina_user2$ = " "
                                         
                                         
                                         If oficina_user2$ = "Nada" And oficina_csr2$ = "Nada" Then
                                                     oficina_user2$ = "None"
                                         End If
                  
                  
                                         If (oficina_user$ = oficina_csr$ Or oficina_user2$ = oficina_csr2$ Or oficina_user$ = oficina_csr2$ Or oficina_user2$ = oficina_csr$) Then
                                                    
                                                    If b$ = "AGENT_COMMERCIAL" Then
                                                         
                                                         ' es manager del agente
                                                         es_manager = 0
                                                         For z = 0 To lista_managers.ListCount - 1
                                                             a$ = RTrim(UCase(Left$(lista_managers.List(z), 20)))
                                                             bb$ = LTrim(RTrim(Right(UCase(lista_managers.List(z)), Len(lista_managers.List(z)) - 20)))
                         
                                                             If a$ = csr$ Then
                                                                  If bb$ = "MANAGER" Or bb$ = "MANAGER_COMMERCIAL" Then
                                                                          es_manager = 1
                                                                          Exit For
                                                                  End If
                                                             End If
                                                                
                                                         Next z
                                                         
                                                         If es_manager = 1 Then
                                                            existe = 1
                                                         Else
                                                            existe = 9
                                                         End If
                                                      
                                                    Else
                                                      existe = 3    '   8/17/2022
                                                      
                                                    End If
                                                      
                                                      ' estaba existe=3
                                                      Exit For
                                         Else
                                                      existe = 9
                                                      Exit For
                  
                                         End If
                             
                                                           
                                        
                                         
                                         
                                            existe = 6
                                            Exit For
                                            
                                            
                                            
                                 ' ElseIf b$ = concepto$ And b$ <> "BF COMMERCIAL" Then
                                  '          existe = 8
                                  '          Exit For
                                            
                                   ElseIf b$ = concepto$ And b$ = "BF COMMERCIAL" Then
                                      ' If manager_Oficinacommercial$ = a$ Then
                                            existe = 1
                                      ' Else
                                      '      existe = 11
                                      ' End If
                                            Exit For
              
              
                                  ElseIf b$ = concepto$ And b$ = "BF" Then
                                            existe = 3
                                            Exit For
                                  End If
         
                           End If
        
        
             Next Y
  
      If existe = 3 Or existe = 8 Or existe = 9 Or existe = 11 Or existe = 1 Then GoTo brinca
  
  
  
  
  
     
  
             ' verifica si es manager
             existe = -1
             For Y = 0 To lista_managers.ListCount - 1
                         a$ = RTrim(UCase(Left$(lista_managers.List(Y), 20)))
                         b$ = LTrim(RTrim(Right(UCase(lista_managers.List(Y)), Len(lista_managers.List(Y)) - 20)))
                         
                         If a$ = csr$ Then ' Or agente$ = a$ Then
                                    If b$ = "MANAGER" Or b$ = "MANAGER_COMMERCIAL" Then
                                                                                                       
                                         ' verifica oficinas a que pertenecen
                                         For w = 1 To Val(ubicacion(0, 0))
                                              If UCase(RTrim(ubicacion(w, 0))) = UCase(agente$) Then
                                                      oficina_user$ = RTrim(ubicacion(w, 1))
                                                      oficina_user2$ = RTrim(ubicacion(w, 2))
                                              End If
                  
                                              If UCase(RTrim(ubicacion(w, 0))) = UCase(csr$) Then
                                                      oficina_csr$ = RTrim(ubicacion(w, 1))
                                                      oficina_csr2$ = RTrim(ubicacion(w, 2))
                                              End If
                                         Next w
                  
                  
                                         If oficina_user2$ = "None" And oficina_csr2$ = "None" Then
                                                      oficina_user2$ = "Nada"
                                         End If
                  
                                         If oficina_user2$ = "" Then oficina_user2$ = " "
                                         
                                         If oficina_user2$ = "Nada" And oficina_csr2$ = "Nada" Then
                                                      oficina_user2$ = "None"
                                         End If
                  
                                         If (oficina_user$ = oficina_csr$ Or oficina_user2$ = oficina_csr2$ Or oficina_user$ = oficina_csr2$ Or oficina_user2$ = oficina_csr$) Then
                                                    
                                                      existe = 1
                                                      Exit For
                                         Else
                                                      existe = 8
                                                      Exit For
                  
                                         End If
                  
                                    
                                               
                    
                                    ElseIf b$ = concepto$ Then
                                                existe = 2
                                                Exit For
                                    End If
         
                         End If
        
        
        
        
                         If a$ = csr2$ Then ' Or agente$ = a$ Then
                                    If b$ = "MANAGER_COMMERCIAL" Then
                                                                                                       
                                         ' verifica oficinas a que pertenecen
                                         For w = 1 To Val(ubicacion(0, 0))
                                              If UCase(RTrim(ubicacion(w, 0))) = UCase(agente$) Then
                                                      oficina_user$ = RTrim(ubicacion(w, 1))
                                                      oficina_user2$ = RTrim(ubicacion(w, 2))
                                              End If
                  
                                              If UCase(RTrim(ubicacion(w, 0))) = UCase(csr$) Then
                                                      oficina_csr$ = RTrim(ubicacion(w, 1))
                                                      oficina_csr2$ = RTrim(ubicacion(w, 2))
                                              End If
                                         Next w
                  
                  
                                         If oficina_user2$ = "None" And oficina_csr2$ = "None" Then
                                                      oficina_user2$ = "Nada"
                                         End If
                  
                                         If oficina_user2$ = "" Then oficina_user2$ = " "
                                         
                                         If oficina_user2$ = "Nada" And oficina_csr2$ = "Nada" Then
                                                      oficina_user2$ = "None"
                                         End If
                                         
                                         
                                         
                                          
                  
                                         If (oficina_user$ = oficina_csr$ Or oficina_user2$ = oficina_csr2$ Or oficina_user$ = oficina_csr2$ Or oficina_user2$ = oficina_csr$) Then
                                                    
                                                      existe = 1
                                                      Exit For
                                         Else
                                                      existe = 8
                                                      Exit For
                  
                                         End If
                  
                                    
                                               
                    
                                    ElseIf b$ = concepto$ Then
                                                existe = 2
                                                Exit For
                                    End If
         
                         End If
        
        
        
                         If a$ = agente$ And csr$ <> agente$ Then
          
                                    If b$ = "MANAGER" Or b$ = "MANAGER_COMMERCIAL" Then
                                    
                                    ' verifica oficinas a que pertenecen
                                          For w = 1 To Val(ubicacion(0, 0))
                                                If UCase(RTrim(ubicacion(w, 0))) = UCase(agente$) Then
                                                            oficina_user$ = RTrim(ubicacion(w, 1))
                                                            oficina_user2$ = RTrim(ubicacion(w, 2))
                                                End If
                  
                                                If UCase(RTrim(ubicacion(w, 0))) = UCase(csr$) Then
                                                            oficina_csr$ = RTrim(ubicacion(w, 1))
                                                            oficina_csr2$ = RTrim(ubicacion(w, 2))
                                                End If
                                          Next w
                  
                                          If oficina_user2$ = "None" And oficina_csr2$ = "None" Then
                                              oficina_user2$ = "Nada"
                                          End If
                                          
                                          If oficina_user2$ = "" Then oficina_user2$ = " "
                                          
                                          
                                          If oficina_user2$ = "Nada" And oficina_csr2$ = "Nada" Then
                                              oficina_user2$ = "None"
                                          End If
                                          
                  
                  
                                          If oficina_user$ = oficina_csr$ Or (oficina_user2$ = oficina_csr2$ And oficina_user2$ <> "") Or oficina_user$ = oficina_csr2$ Or oficina_user2$ = oficina_csr$ Then
                                                            existe = 3     'c=0
                                                            Exit For
                                          Else
                                                            existe = 8     'c=c/2
                                                            Exit For
                                          End If
                  
                  
                                                
                    
                                    ElseIf b$ = concepto$ Then
                                                existe = 2   'c=c
                                                Exit For
                                    End If
        
        
        
           
                         End If
        
                         If a$ = agente$ Then
        
                         End If
        
        
             Next Y
     
'            If existe = 1 Or existe = 3 Then GoTo brinca
  
     
  
     
brinca:

     If BF_CALL = 1 Then
         For k = 0 To List2.ListCount - 1
              concepto$ = Left(List2.List(k), 20)
              recibo$ = Mid$(List2.List(k), 21, 6)
              usuario$ = Mid$(List2.List(k), 28, 20)
              cantidad$ = Right(List2.List(k), 9)
              
              If Format(Val(cant_guardada_CALL_CENTER$), "###0.0") = Format(Val(cantidad$), "###0.0") And LTrim(RTrim(UCase(concepto$))) <> "BF CALL CENTER" Then
                existe = 1   ' asigna todo
                If par_encontrado = 1 Then
                   par_encontrado = par_encontrado + 1
                End If
                Exit For
              End If
              
           Next k

           BF_CALL = 0
           
     End If

  
     If UCase(concepto$) = "BF CALL CENTER" Then
       If (UCase(oficina_user$) = "JA - PHONE SALES") Then
          
       Else
           existe = 5  ' asigna CERO
           BF_CALL = 1
           par_encontrado = par_encontrado + 1
           cant_guardada_CALL_CENTER$ = LTrim(Str(cant_del_concepto))  'cantidad$
           USUARIO_DE_PS$ = csr$
       End If
       
       
     End If
     
     
     If par_encontrado = 2 Then
       par_encontrado = 0
       
     End If
     
     
     ' detecta si el invoice con JENNIFERR esta en la lista de facturas ya pasadas de tiempo
     
     
     
    
  ' ---------------------------  DETECTA SI EL AGENTE FUE EL COBRADOR DEL CHEQUE ----------------------------------------
    If RTrim(LTrim(UCase(concepto$))) = "LATE FEE" Then
      
     agente_cobrador$ = ""
     cobrador$ = ""
 
     sSelect = "select idemployeeUSR, IdEmployeeCSR1, IdEmployeeCSR2  from [ReceiptsDTL] recdtl " & _
     "inner join  ReceiptsHDR  rechdr on rechdr.IDReceiptHDR =recdtl.IdReceiptHDR " & _
     "inner join InvoiceItemCatalog iitem on iitem.IdInvoiceItem=recdtl.IdInvoiceItem " & _
     "where rechdr.Active=1 and recdtl.[IdReceiptHDR] in (" + recibo$ + ")"
     
     ' and iitem.InvoiceItemName='Late Fee'"  se quito esta parte   4/2/2024
 
     Rs.Open sSelect, base, adOpenUnspecified
     agente_cobrador$ = Rs(0)
     agente_csr1$ = Rs(1)
     agente_csr1$ = Rs(2)
     Rs.Close
     
     
      sSelect = "select username from EmployeeInfo where IDEmployee='" + agente_cobrador$ + "'"
      Rs.Open sSelect, base, adOpenUnspecified
      cobrador$ = UCase$(Rs(0))
      Rs.Close
       
    End If
    
    
     
     
  
     If agente$ = "JENNIFERR" And (UCase(concepto$) = "INVOICE") Then
       If UCase(concepto$) = "INVOICE" Then
       
        hayado = 0
        For q = 0 To lista_invoices30.ListCount - 1
            
            cant_invoice = Val(Format(Mid$(lista_invoices30.List(q), 8, 9), "00000.00"))
            
            
            If cant_del_concepto = cant_invoice Then
             csr$ = ""
             n$ = LTrim(RTrim(lista_invoices30.List(q)))
             For Y = Len(n$) To 1 Step -1
                If Mid$(n$, Y, 1) <> Space(1) Then
                   conta = conta + 1
                   csr$ = LTrim(RTrim(Right$(n$, conta)))
                Else
                   Exit For
                End If
             Next Y
 
             agente1$ = csr$
  
             r$ = RTrim(Left(n$, Len(n$) - (conta)))
             conta = 0
             For Y = Len(r$) To 1 Step -1
               If Mid$(r$, Y, 1) <> Space(1) Then
                     conta = conta + 1
                     csr$ = LTrim(RTrim(Right$(r$, conta)))
               Else
                     Exit For
               End If
             Next Y
             csr1$ = csr$
            
            
             If UCase(csr$) = UCase(csr1$) Or UCase(csr$) = UCase(agente1$) Then
               existe = 1
               hayado = 1
               Exit For
             End If
            
            End If
            
            
        Next q
       
        If hayado = 1 Then
           GoTo revisa
        End If
       
        existe = 9
        GoTo revisa
        
       End If
       
        
        
     ElseIf csr$ = "JENNIFERR" And (UCase(concepto$) = "INVOICE" Or UCase(concepto$) = "LATE FEE") Then
        
        If UCase(cobrador$) = "JENNIFERR" And UCase(concepto$) = "LATE FEE" Then
           existe = 3
        Else
           existe = 9
        End If
        
     ElseIf cobrador$ = agente$ And UCase(concepto$) = "LATE FEE" Then
     
        If UCase(cobrador$) = agente$ And UCase(concepto$) = "LATE FEE" Then
           existe = 1
        Else
           existe = 9
        End If
                        
        
     End If
     
     
  
    ' If Alta_Comercial = 1 Then
    '    existe = 5  ' asigna CERO
    ' End If
  
     
     'If Es_Monterrey = 1 Then
     '  existe = 1
    ' End If
  
revisa:
  
     
  
  
     If existe = 1 Then
       c = (Val(Right(List1.List(t), 9)))
     ElseIf existe = 2 Then
       c = (Val(Right(List1.List(t), 9))) '/ 2
     ElseIf existe = 3 Then
       c = 0
     ElseIf existe = 4 Then
       c = (Val(Right(List1.List(t), 9)))
     ElseIf existe = 5 Then
       c = 0
     ElseIf existe = 6 Then
       c = (Val(Right(List1.List(t), 9))) / 2
       ' c = 0
     ElseIf existe = 8 Then
        c = (Val(Right(List1.List(t), 9))) / 2
       
     ElseIf existe = 7 Then
       c = (Val(Right(List1.List(t), 9))) / 2   ' tenia c = 0
    ElseIf existe = 9 Then
       c = (Val(Right(List1.List(t), 9))) / 2
       
       
     ElseIf existe = 10 Then
       c = cant_del_concepto / 2
       'c = 0
       
     ElseIf existe = 11 Then   ' BF COMMERCIAL
       c = 0
       
     Else
       c = (Val(Right(List1.List(t), 9))) / 2
     End If
  End If
  
suma_la_cantidad:

  If (UCase(concepto$)) = "INVOICE" Then
  
     If csr$ = agente$ Or Es_Monterrey = 1 Then
  
       List9.AddItem cant_del_concepto
       
     Else
     
        If oficina_csr$ = oficina_agente$ Then
       
            If manager_agente$ = agente$ Then
                List9.AddItem 0
            ElseIf manager_csr$ = csr$ Then
                List9.AddItem cant_del_concepto
            Else
                List9.AddItem cant_del_concepto / 2
            End If
       
        Else
        
            List9.AddItem cant_del_concepto / 2
            
        End If
     
     End If
     
  End If


    
comercial_saltado:
    
  gtotal = gtotal + c
  Alta_Comercial = 0

Next t



If par_encontrado = 1 Then
   gtotal = gtotal + (Val(cant_guardada_CALL_CENTER$) / 2)

End If
   
   
   
If UCase(lblagent.Caption) = "JENNIFERR" Then
    GoTo brinca_por_recaudacion
End If


' descuenta el invoice de mas de 30 dias si existe

For t = 0 To lista_invoices30.ListCount - 1
  csr$ = ""
  conta = 0
  
  
  
  
  n$ = LTrim(RTrim(lista_invoices30.List(t)))
  For Y = Len(n$) To 1 Step -1
    If Mid$(n$, Y, 1) <> Space(1) Then
         conta = conta + 1
         csr$ = Right$(n$, conta)
    Else
         Exit For
    End If
  Next Y
 
  agente1$ = csr$
  
  
  
  
  r$ = RTrim(Left(n$, Len(n$) - (conta)))
  
  conta = 0
  For Y = Len(r$) To 1 Step -1
    If Mid$(r$, Y, 1) <> Space(1) Then
         conta = conta + 1
         csr$ = Right$(r$, conta)
    Else
         Exit For
    End If
  Next Y
  csr1$ = csr$
  
  
  ' verifica si esta asignado
  existe = 0
  For k = 0 To List1.ListCount - 1
    '  List1.AddItem Format(r$, "!@@@@@@@@@@@@@@@@@@@@") + " " + Format(concepto$, "!@@@@@@@@@@@@@@@@@@@@") + Space(1) + cantidad$
    nom_fact$ = UCase(RTrim(Left(List1.List(k), 20)))
    If UCase(nom_fact$) = UCase(agente1$) Or UCase(nom_fact$) = UCase(csr1$) Then
       existe = 1
       Exit For
    Else
      
    
    End If
    
    
  Next k
  
  
  If existe = 0 Then
     GoTo no_hagas
  End If
  
  
  
  ' asigna la oficina a user y a csr y a agente
   
   For Y = 1 To Val(ubicacion(0, 0))
      If UCase(RTrim(ubicacion(Y, 0))) = UCase(csr1$) Then
         oficina_csr$ = ubicacion(Y, 1)
         Exit For
      End If
   Next Y
   
   
   For Y = 1 To Val(ubicacion(0, 0))
      If UCase(RTrim(ubicacion(Y, 0))) = UCase(user1$) Then
         oficina_user$ = ubicacion(Y, 1)
         Exit For
      End If
   Next Y
      
      
   For Y = 1 To Val(ubicacion(0, 0))
      If UCase(RTrim(ubicacion(Y, 0))) = UCase(agente1$) Then
         oficina_agente$ = ubicacion(Y, 1)
         Exit For
      End If
   Next Y
    
    
    
    
    
    
  
  
  
  
  r$ = RTrim(Left(r$, Len(r$) - conta - 1))
  
  
  For w = 0 To List2.ListCount - 1
    '          linea$ = Format(concepto$, "!@@@@@@@@@@@@@@@@@@@@") + Format(recibo$, "000000") + " " + Format(r$, "!@@@@@@@@@@@@@@@@@@@@") + " " + Format(cantidad$, "000000.00")
        recibox$ = Mid$(List2.List(w), 21, 6)
        If recibox$ = Left(r$, 6) Then
          cant = Val(Right(List2.List(w), 9))
          Exit For
        End If
        
  Next w
  
  pos = InStr(1, r$, "$")
  'cant = Val(Right(r$, Len(r$) - (pos)))
  
  existe = 0
  
  If RTrim(UCase(lblagent.Caption)) = RTrim(UCase(agente1$)) And RTrim(UCase(lblagent.Caption)) = RTrim(UCase(csr1$)) And Right(oficina_csr$, 9) <> "MONTERREY" Then
    
      
        gtotal = gtotal - cant
        existe = 1
      
  End If
  
  
  If RTrim(UCase(lblagent.Caption)) <> RTrim(UCase(csr1$)) And RTrim(UCase(lblagent.Caption)) = RTrim(UCase(agente1$)) And Right(oficina_csr$, 9) <> "MONTERREY" Then
      
         gtotal = gtotal - (cant / 2)
         existe = 1
    
  End If
  
     
   
  
   If RTrim(UCase(lblagent.Caption)) <> RTrim(UCase(agente1$)) And RTrim(UCase(lblagent.Caption)) = RTrim(UCase(csr1$)) And Right(oficina_csr$, 9) <> "MONTERREY" Then
       
    
        gtotal = gtotal - (cant / 2)
        If gtotal < 0 Then gtotal = 0
        existe = 1
   
       
      
  
  End If
  
  
no_hagas:
  
  
Next t



brinca_por_recaudacion:


X = Redondear(gtotal, 2)  ' redondea a 2 decimales
If X < 0 Then X = 0
lbltotal_invoices.Caption = Format(X, "$###,##0.00")




' realiza total de cada agente
Erase tabla
    csr$ = ""
    For t = 0 To List1.ListCount - 1
      nombre$ = Left(List1.List(t), 20)
      existe = 0
      For Y = 0 To 19
         If tabla(Y, 0) = "" Then
            Exit For
         End If
      
         If tabla(Y, 0) = nombre$ Then
            tabla(Y, 1) = tabla(Y, 1) + Val(Right(List1.List(t), 9))
            existe = 1
            Exit For
         End If
         
      Next Y
      If existe = 0 Then
         tabla(Y, 0) = nombre$
         tabla(Y, 1) = tabla(Y, 1) + Val(Right(List1.List(t), 9))
      End If
        
        
        
       
    Next t
    
    

carga_treeview

  Exit Sub
  
  
  
  
' //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'   R u t i n a s                        R u t i n a s                        R u t i n a s                                  R u t i n a s
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\



  
  
' *************************************
check_manager:
' *************************************


     ' verifica si CSR es manager
     existe = 0
     Es_Monterrey = 0
     
     For z = 0 To lista_managers.ListCount - 1
       a$ = RTrim(UCase(Left$(lista_managers.List(z), 20)))
       b$ = LTrim(RTrim(Right(UCase(lista_managers.List(z)), Len(lista_managers.List(z)) - 20)))
       
       If UCase(b$) = "BF COMMERCIAL" Then
          GoTo brincado_aqui3
       End If
       
       
       'If a$ = "MFUENTES" Then Stop
       
       
       If a$ = UCase(userx$) Then '
          If b$ = "MANAGER" Or b$ = "MANAGER_COMMERCIAL" Or b$ = "COMMERCIAL" Then
             manager_csr$ = a$
             existe = 1
            ' Exit For
          End If
        End If
        
        
        If a$ = UCase(csr$) Then '
          If b$ = "MANAGER" Or b$ = "MANAGER_COMMERCIAL" Or b$ = "COMMERCIAL" Then
             manager_csr$ = a$
             existe = 1
             'Exit For
          End If
          
        End If
        
        
        If a$ = UCase(user$) Then
            If b$ = "MANAGER" Or b$ = "MANAGER_COMMERCIAL" Or b$ = "COMMERCIAL" Then
             manager_user$ = a$
             existe = 1
             'Exit For
            End If
             
        End If
          
          
          
        If a$ = UCase(agente$) Then
          If b$ = "MANAGER" Or b$ = "MANAGER_COMMERCIAL" Or b$ = "COMMERCIAL" Then
             manager_agente$ = a$
             existe = 1
             'Exit For
          End If
             
        End If
        
        
        If a$ = UCase(agente$) Or a$ = UCase(csr$) Or a$ = UCase(user$) Then
          If b$ = "COMMERCIAL" Then ' And Alta_Comercial = 0 Then
             manager_commercial$ = a$
             
             'Alta_Comercial = 1
             existe = 1
             'Exit For
          End If
             
        End If
        
        
        
        If a$ = UCase(agente$) Or a$ = UCase(csr$) Or a$ = UCase(user$) Then
          If b$ = "MANAGER_COMMERCIAL" Then
             manager_Oficinacommercial$ = a$
             existe = 1
             'Exit For
          End If
             
        End If
        
        
        If a$ = UCase(agente$) Or a$ = UCase(csr$) Or a$ = UCase(user$) Then
          If b$ = "AGENT_COMMERCIAL" Then
             Agent_Oficinacommercial$ = a$
             existe = 1
             'Exit For
          End If
             
        End If
        
        
        If a$ = UCase(agente$) Or a$ = UCase(csr$) Or a$ = UCase(user$) Then
            If b$ = "MONTERREY" Then
               Es_Monterrey = 1
              ' Exit For
            End If
        End If
        
        
        If existe = 1 Or Es_Monterrey = 1 Then
             Exit For
        End If
        
brincado_aqui3:
        
      Next z
      
      Return
      
      
 ' ==============================================
check_commercial:
 ' ==============================================
 
 ' verifica si CSR o USER es manager-COMERCIAL
        existe = 0
        CSR2_Es_manager_commercial = 0
    
        For z = 0 To lista_managers.ListCount - 1
              a$ = RTrim(UCase(Left$(lista_managers.List(z), 20)))
              b$ = LTrim(RTrim(Right(UCase(lista_managers.List(z)), Len(lista_managers.List(z)) - 20)))
                
              If a$ = agente$ Or a$ = csr$ Or a$ = user$ Then
                    If b$ = "COMMERCIAL" Or b$ = "MANAGER_COMMERCIAL" And Alta_Comercial = 0 Then
                           manager_commercial$ = a$
                           Alta_Comercial = 1
                           existe = 1
             
                    End If
              End If
        
                            
              If a$ = agente$ Or a$ = csr$ Or a$ = user$ Then
                   If b$ = "AGENT_COMMERCIAL" Then
                          Agent_Oficinacommercial$ = a$
                          existe = 1
                   End If
              End If
             
                
              If a$ = csr2$ Then
                    If b$ = "COMMERCIAL" Or b$ = "MANAGER_COMMERCIAL" Or b$ = "AGENT_COMMERCIAL" Then
                          CSR2_Es_manager_commercial = 1
                          existe = 1
             
                    End If
              End If
                   
        
        Next z
        
        If existe = 1 Then
                    contador_comercial = contador_comercial + 1
       
                    If Agent_Oficinacommercial$ <> "" Then
                           Commercial$(contador_comercial, 0) = user$
                           Commercial$(contador_comercial, 1) = cantidad$
                           Commercial$(contador_comercial, 2) = Agent_Oficinacommercial$
                    Else
                           Commercial$(contador_comercial, 0) = user$
                           Commercial$(contador_comercial, 1) = cantidad$
                           Commercial$(contador_comercial, 2) = manager_csr2$
                    End If
        End If
        
        
        Return
      
      
      
   
      
      
      
      
' =============================================
agentes_de_factura:
' =============================================
      
   If (UCase$(LTrim(concepto$)) = "INVOICE") Or (UCase$(LTrim(concepto$)) = "LATE FEE") Then
    
        user_original$ = user$
        csr_original$ = csr$
   
        X = agentes_de_invoice(recibo$)
   
            csr$ = UCase(csr_original$)
            user$ = UCase(user_original$)
     
   End If
  
  
   Return
   
   
   
' =============================================
Asigna_oficinas:
' =============================================
   
   
   oficina_agente$ = ""
   oficina_user$ = ""
   oficina_user2$ = ""
   oficina_csr$ = ""
   oficina_csr2$ = ""
   
   ' asigna la oficina a user y a csr y a agente,
   
   For Y = 1 To Val(ubicacion(0, 0))
      If UCase(RTrim(ubicacion(Y, 0))) = UCase(csr$) Then
         oficina_csr$ = ubicacion(Y, 1)
      End If
      
      If UCase(RTrim(ubicacion(Y, 0))) = UCase(user$) Then
         oficina_user$ = ubicacion(Y, 1)
      End If
      
      If UCase(RTrim(ubicacion(Y, 0))) = UCase(agente$) Then
         oficina_agente$ = ubicacion(Y, 1)
      End If
   Next Y
   
   Return
   
      
End Sub

Private Sub lista_invoices30_Click()
On Error Resume Next
lista_invoices30.Selected(lista_invoices30.ListIndex) = False
Text1.SetFocus

End Sub

Private Sub lista_managers_Click()
On Error Resume Next
lista_managers.Selected(lista_managers.ListIndex) = False
Text1.SetFocus

End Sub

Private Sub lista_users_shared_Click()
On Error Resume Next
lista_users_shared.Selected(lista_users_shared.ListIndex) = False
Text1.SetFocus
End Sub


Private Sub op_invoice_Click(Index As Integer)
On Error Resume Next
lblmsg.Caption = "Please, wait a moment... loading all data"
msg.Visible = True
msg.Refresh

expiracion_invoices = Index
calcula_invoices
btncargar_excel_Click

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
seg = seg + 1
If seg >= 2 Then
  seg = 0
  cargado = False
  Timer1.Enabled = False
End If


End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
'MsgBox TreeView1.SelectedItem
On Error Resume Next

' Left(List1.List(t), 20)

fila_selecta = TreeView1.SelectedItem.Index
concepto_selecto$ = TreeView1.Nodes(fila_selecta)
agente_selecto$ = ""

For Y = fila_selecta To 1 Step -1
  r$ = TreeView1.Nodes(Y)
  If Left(r$, 1) <> Space(1) Then
     agente_selecto$ = RTrim(Left(r$, 20))
     Exit For
  End If
Next Y



encontrado = 0
For t = 0 To List1.ListCount - 1
  n$ = Right(List1.List(t), Len(List1.List(t)) - 20)
  a$ = RTrim(Left(List1.List(t), 20))
  
  If n$ = TreeView1.SelectedItem And a$ = agente_selecto$ Then
     agente$ = Left(List1.List(t), 20)
     encontrado = 1
     Exit For
  End If
Next t

     
List3.Clear

For t = 0 To List2.ListCount - 1
   concepto$ = Left(List2.List(t), 20)
   recibo$ = Mid$(List2.List(t), 21, 6)
   usuario$ = Mid$(List2.List(t), 28, 20)
   cantidad$ = Right(List2.List(t), 9)
   If UCase(RTrim(agente$)) = UCase(RTrim(usuario$)) Then
      If LTrim(RTrim(Left(TreeView1.SelectedItem, 20))) = RTrim(concepto$) Then
        ' checa si el invoice es mayor de 30 dias
         existe = 0
         For Y = 0 To lista_invoices30.ListCount - 1
           If Format(recibo$, "00000") = Left(lista_invoices30.List(Y), 5) Then
              List3.AddItem recibo$ + " " + Format(Format(Val(cantidad$) * (-1), "$##0.00"), "@@@@@@@")
              existe = 1
              Exit For
           End If
         Next Y
         
         If existe = 0 Then
            List3.AddItem recibo$ + " " + Format(Format(cantidad$, "$##0.00"), "@@@@@@@")
         End If
      End If
   End If
   
Next t



lbltotal_unidades.Caption = Format(List3.ListCount, "###,##0")

End Sub



Public Sub carga_treeview()
On Error Resume Next
   
 ' CARGA LAS NOTAS
 nf = FreeFile
    Open "c:\goals\test.txt" For Output Shared As #nf
    
    n$ = ""
    For t = 0 To List1.ListCount - 1
      nombre$ = RTrim(Left(List1.List(t), 20))
      If n$ = nombre$ Then
        Print #nf, Chr$(9) + Right(List1.List(t), Len(List1.List(t)) - 20)
      Else
        For z = 0 To 19
          If RTrim(tabla(z, 0)) = RTrim(nombre$) Then
             nombre$ = Format(nombre$, "!@@@@@@@@@@@@@@@@@@@@") + Space(1) + Format(tabla(z, 1), "###,##0.00")
             Exit For
          End If
         Next z
      
        Print #nf, nombre$
        Print #nf, Chr$(9) + Right(List1.List(t), Len(List1.List(t)) - 20)
        n$ = RTrim(Left(nombre$, 20))
      End If
    Next t
    
   Close nf



If Right$(file_name, 1) <> "\" Then file_name = file_name & "\"
file_name = "c:\goals\test.txt"
LoadTreeViewFromFile file_name, TreeView1
   




End Sub

Public Sub carga_treeview_extra()
On Error Resume Next
   
 Tipo$ = TreeView1.SelectedItem
 tipo_nombre$ = Left(Tipo$, 20)
 'tipo_cantidad
   
   
 ' CARGA LAS NOTAS
 nf = FreeFile
    Open "c:\goals\test.txt" For Output Shared As #nf
    
    n$ = ""
    For t = 0 To List1.ListCount - 1
      nombre$ = Left(List1.List(t), 20)
      If n$ = nombre$ Then
        Print #nf, Chr$(9) + Right(List1.List(t), Len(List1.List(t)) - 20)
      Else
        For z = 0 To 19
          If tabla(z, 0) = nombre$ Then
             nombre$ = nombre$ + Space(1) + Format(tabla(z, 1), "###,##0.00")
             Exit For
          End If
         Next z
      
        Print #nf, nombre$
        Print #nf, Chr$(9) + Right(List1.List(t), Len(List1.List(t)) - 20)
        n$ = Left(nombre$, 20)
      End If
    Next t
    
   Close nf



If Right$(file_name, 1) <> "\" Then file_name = file_name & "\"
file_name = "c:\goals\test.txt"
LoadTreeViewFromFile file_name, TreeView1
   

End Sub

Public Sub carga_attributos()
On Error Resume Next

Dim sSelect As String
Dim Rs As ADODB.Recordset
    
Set Rs = New ADODB.Recordset



lista_managers.Clear
lista_PhoneSales.Clear



 grid.Clear
 
 sSelect = "Select idemppayrolllink, idemployeeLAE, idpaytype, initials, manager, MTY, Commercial, PhoneSales, Exception1,Typeofexception1," & _
 "Exception2,Typeofexception2,Exception3,Typeofexception3,Exception4,Typeofexception4,Exception5,Typeofexception5,deb_mod, nb_deb_mod, porc_mod, autorizado, " & _
 "office1, office2, emp.firstname, emp.lastname1, emp.username, managercommercial from [payrollconfig] payroll " & _
 "inner join EmployeeInfo emp on emp.IDEmployee=payroll.IdEmployeeLAE "
 
 '"where emp.username='" + txtagent.Text + "'"
  ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
        
     ' Permitir redimensionar las columnas
    grid.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid.DataSource = Rs
                         
    Rs.Close



For t = 1 To grid.Rows - 1
   grid.row = t
   grid.col = 8
   phone_sales1$ = grid.Text
   
   grid.col = 27
   username1$ = grid.Text
   
   grid.col = 28
   Manager_Commercial1$ = grid.Text
   
   
 If Val(phone_sales1$) = 1 Then
    n$ = Format(username1$, "!@@@@@@@@@@@@@@@@@@@@") + Space(2) + Format("PHONE SALES", "!@@@@@@@@@@@@@@@@@@")
    lista_PhoneSales.AddItem n$
 End If
 
 
 grid.col = 5
 manager1$ = grid.Text
 
 grid.col = 6
 monterrey1$ = grid.Text
 
 grid.col = 7
 Commercial1$ = grid.Text
 
 grid.col = 28
 Manager_Commercial1$ = grid.Text
 
 
 existe = 0
 If Val(manager1$) = 1 Or Val(monterrey1$) = 1 Or Val(Commercial1$) = 1 Or Val(Manager_Commercial1$) = 1 Or Val(Manager_Commercial1$) = 2 Then
 
   grid.col = 9
   excep1$ = grid.Text
    
   grid.col = 10
   tipo_excep1$ = grid.Text
   
   grid.col = 11
   excep2$ = grid.Text
   
   grid.col = 12
   tipo_excep2$ = grid.Text
   
   grid.col = 13
   excep3$ = grid.Text
   
   grid.col = 14
   tipo_excep3$ = grid.Text
   
   grid.col = 15
   excep4$ = grid.Text
   
   grid.col = 16
   tipo_excep4$ = grid.Text
   
   grid.col = 17
   excep5$ = grid.Text
   
   grid.col = 18
   tipo_excep5$ = grid.Text
   
   
   ' checa si hay alguna excepcion
   If Val(tipo_excep1$) = 1 Then
      a$ = cboexcepcion.List(Val(excep1$))
      n$ = Format(username1$, "!@@@@@@@@@@@@@@@@@@@@") + Space(2) + Format(a$, "!@@@@@@@@@@@@@@@@@@")
      lista_managers.AddItem n$
      existe = 1
   End If
   
   If Val(tipo_excep2$) = 1 Then
      a$ = cboexcepcion.List(Val(excep2$))
      n$ = Format(username1$, "!@@@@@@@@@@@@@@@@@@@@") + Space(2) + Format(a$, "!@@@@@@@@@@@@@@@@@@")
      lista_managers.AddItem n$
      existe = 1
   End If
   
   If Val(tipo_excep3$) = 1 Then
      a$ = cboexcepcion.List(Val(excep3$))
      n$ = Format(username1$, "!@@@@@@@@@@@@@@@@@@@@") + Space(2) + Format(a$, "!@@@@@@@@@@@@@@@@@@")
      lista_managers.AddItem n$
      existe = 1
   End If
   
   If Val(tipo_excep4$) = 1 Then
      a$ = cboexcepcion.List(Val(excep4$))
      n$ = Format(username1$, "!@@@@@@@@@@@@@@@@@@@@") + Space(2) + Format(a$, "!@@@@@@@@@@@@@@@@@@")
      lista_managers.AddItem n$
      existe = 1
   End If
   
   If Val(tipo_excep5$) = 1 Then
      a$ = cboexcepcion.List(Val(excep5$))
      n$ = Format(username1$, "!@@@@@@@@@@@@@@@@@@@@") + Space(2) + Format(a$, "!@@@@@@@@@@@@@@@@@@")
      lista_managers.AddItem n$
      existe = 1
   End If
      
   If Val(monterrey1$) = 1 Then
     existe = 2
     n$ = Format(username1$, "!@@@@@@@@@@@@@@@@@@@@") + Space(2) + Format("MONTERREY", "!@@@@@@@@@@@@@@@@@@")
     lista_managers.AddItem n$
   End If
 
    
   If Val(Commercial1$) = 1 Then
     existe = 2
     n$ = Format(username1$, "!@@@@@@@@@@@@@@@@@@@@") + Space(2) + Format("COMMERCIAL", "!@@@@@@@@@@@@@@@@@@")
     lista_managers.AddItem n$
   End If
 
 
   If Val(Manager_Commercial1$) = 1 Then
     existe = 2
     n$ = Format(username1$, "!@@@@@@@@@@@@@@@@@@@@") + Space(2) + Format("MANAGER_COMMERCIAL", "!@@@@@@@@@@@@@@@@@@")
     lista_managers.AddItem n$
   End If
 
 
   If Val(Manager_Commercial1$) = 2 Then
     existe = 2
     n$ = Format(username1$, "!@@@@@@@@@@@@@@@@@@@@") + Space(2) + Format("AGENT_COMMERCIAL", "!@@@@@@@@@@@@@@@@@@")
     lista_managers.AddItem n$
   End If
     
   If existe = 0 Then
     n$ = Format(username1$, "!@@@@@@@@@@@@@@@@@@@@") + Space(2) + Format("MANAGER", "!@@@@@@@@@@@@@@@@@@")
     lista_managers.AddItem n$
   End If
   
   
 
 End If
 
Next t



End Sub

Public Sub separa_campos_inv()
On Error Resume Next


GoTo sigue


grid.Clear
grid.cols = 10
grid.Rows = Grid3.Rows

Grid3.row = 0  '1
grid.row = 0

   Grid3.col = 3   ' RECEIPT
   grid.col = 1
   grid.Text = Grid3.Text

   Grid3.col = 4  ' DATE
   grid.col = 2
   grid.Text = Grid3.Text
   
   Grid3.col = 5  ' CUST ID
   grid.col = 3
   grid.Text = Grid3.Text
   
   Grid3.col = 7  ' FOR
   grid.col = 4
   grid.Text = Grid3.Text
   
   Grid3.col = 10  ' NON-FID
   grid.col = 5
   grid.Text = Grid3.Text
   
   Grid3.col = 12  ' USER
   grid.col = 6
   grid.Text = Grid3.Text
      
   Grid3.col = 13  ' CSR
   grid.col = 7
   grid.Text = Grid3.Text
   
   Grid3.col = 19  ' amount paid
   grid.col = 8
   grid.Text = Grid3.Text
   
   Grid3.col = 20  ' balance due
   grid.col = 9
   grid.Text = Grid3.Text


For t = 1 To Grid3.Rows
   
   
   Grid3.row = t
   grid.row = t
   
   Grid3.col = 0
   grid.col = 0
   grid.Text = Grid3.Text
   
      
   
   Grid3.col = 3   ' RECEIPT
   grid.col = 1
   grid.Text = Grid3.Text

   Grid3.col = 4  ' DATE
   grid.col = 2
   grid.Text = Grid3.Text
   
   Grid3.col = 5  ' CUST ID
   grid.col = 3
   grid.Text = Grid3.Text
   
   Grid3.col = 7  ' FOR
   grid.col = 4
   grid.Text = Grid3.Text
   
   Grid3.col = 10  ' NON-FID
   grid.col = 5
   grid.Text = Grid3.Text
   
   Grid3.col = 12  ' USER
   grid.col = 6
   grid.Text = Grid3.Text
      
   Grid3.col = 13  ' CSR
   grid.col = 7
   grid.Text = Grid3.Text
   
   Grid3.col = 19  ' amount paid
   grid.col = 8
   grid.Text = Grid3.Text
   
   Grid3.col = 20  ' balance due
   grid.col = 9
   grid.Text = Grid3.Text
   
   
Next t




' transfiere datos necesarios solamente

sigue:


Grid3.Clear
Grid3.cols = 10
Grid3.Rows = grid.Rows

Grid3.row = 0
grid.row = 0
For t = 0 To 9
  Grid3.col = t
  grid.col = t
  Grid3.Text = grid.Text
  
Next t


Erase tabla_INV



List4.Clear
For z = 1 To grid.Rows - 1
   grid.row = z
   
      grid.col = 3
      cust_id = grid.Text
      
      If cust_id <> "" Then
        List4.AddItem Format(cust_id, "00000") + " " + Format(z, "0000")
      End If
   
   
Next z



cont = 0
For t = 0 To List4.ListCount - 1
   fila = Val(Right(List4.List(t), 4))
   grid.row = fila
   Grid3.row = t + 1
   
   Grid3.col = 0
   cont = cont + 1
   Grid3.Text = cont
   For Y = 1 To 9
      Grid3.col = Y
      grid.col = Y
      Grid3.Text = grid.Text
      tabla_INV(t + 1, Y) = grid.Text
   Next Y
Next t


'For z = 2 To grid.Rows
   'grid3.row = z - 1
   'grid.row = z
   
   
   
 '  For Y = 0 To 9
    ' grid3.col = Y
    ' grid.col = Y
    ' grid3.Text = grid.Text
     
 '    tabla_INV(z - 1, Y) = grid.Text
 '  Next Y
'Next z
   
   
   
   
 


setup_grid3



calcula_invoices

End Sub

Public Sub calcula_invoices()
On Error Resume Next
Dim sSelect As String
Dim Rs As ADODB.Recordset
    
Set Rs = New ADODB.Recordset


lista_invoices30.Clear

For t = 1 To Grid3.Rows
  Grid3.row = t
  Grid3.col = 3
  cust_id = Val(Grid3.Text)
  
  Grid3.col = 4
  concepto$ = LTrim(RTrim(UCase(Grid3.Text)))
  
         Grid3.col = 1
         recibo$ = Grid3.Text
         
       '  If recibo$ = "249540" Then Stop
         
         Grid3.col = 2
         fecha_final_recibo$ = Format(Grid3.Text, "mm/dd/yyyy")
         
         Grid3.col = 7
         cantidad = Val(Grid3.Text)
         
         Grid3.col = 5
         usuario1$ = (Grid3.Text)
         
         Grid3.col = 6
         csr1$ = Grid3.Text
         
         
     existe = 0
  If (UCase$(LTrim(concepto$)) = "INVOICE") Then
     
         
         sSelect = "select rechdr.Date, IDReceiptHDR, idemployeeUSR, idemployeeCSR1, balancedue from ReceiptsBalancePayments recbalpay " & _
         "inner join ReceiptsHDR rechdr on recbalpay.IdReceiptsHDRWBalance=rechdr.IDReceiptHDR " & _
         "Where IdReceiptsHDRPayBalance='" + Format(recibo$, "#000000") + "'"
         
          ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
          Rs.Open sSelect, base, adOpenUnspecified
    
          fecha_de_creacion_recibo$ = Format(Rs(0), "mm/dd/yyyy")
          reciboHDR$ = Rs(1)
          
          user2$ = Rs(2)
          csr2$ = Rs(3)
          
          balance_factura$ = Rs(4)
          
          Rs.Close
          
          
          
          
          sSelect = "select IdReceiptsHDRwBalance from ReceiptsBalancePayments where IdReceiptsHDRPayBalance='" + Format(recibo$, "#000000") + "'"
          Rs.Open sSelect, base, adOpenUnspecified
          ID_reciboHDRW$ = Rs(0)
          Rs.Close
              
    
    
          sSelect = "select idemployeeUSR, idemployeeCSR1 from ReceiptsHDR where IDReceiptHDR='" + ID_reciboHDRW$ + "'"
          Rs.Open sSelect, base, adOpenUnspecified
          user1$ = Rs(0)
          csr1$ = Rs(1)
          Rs.Close
          
                    
          
          sSelect = "select username from EmployeeInfo where IDEmployee='" + user1$ + "'"
          Rs.Open sSelect, base, adOpenUnspecified
          user1$ = Rs(0)
          Rs.Close
           
           
          sSelect = "select username from EmployeeInfo where IDEmployee='" + csr1$ + "'"
          Rs.Open sSelect, base, adOpenUnspecified
          csr1$ = Rs(0)
          Rs.Close
               
               
               
            ' ==========================================================================
               
            X = agentes_de_invoice(recibo$)
            
           
   
     
            
               csr1$ = UCase(csr_original$)
               user1$ = UCase(user_original$)
        
            
               
               
          ' ==========================================================================
               
         
         
         ' If recibo$ = "224897" Then
         '   Stop
         
       '   End If
         
          sSelect = "select balancedue, amountpaid from ReceiptsHDR where IDReceiptHDR='" + reciboHDR$ + "'"
          Rs.Open sSelect, base, adOpenUnspecified
    
          balance_factura$ = Rs(0)
          cantidad_pagada$ = Rs(1)
          
          
            Rs.Close
          
          
         
         
         
         fecha_factura$ = fecha_de_creacion_recibo$
         ano_factura$ = Right(fecha_factura$, 4)
         ano_actual$ = Format(Now, "yyyy")
         
         If Val(ano_factura$) < Val(ano_actual$) Then
            dia_fecha_factura$ = Format(fecha_factura$, "y")
            dias_pasados = (365 - Val(dia_fecha_factura$))
            dia_fecha_actual$ = Format(Now, "y")
            
            total_dias = dias_pasados + Val(dia_fecha_actual$)
            'If total_dias > 30 Then
               dias_que_han_pasado = total_dias
            'End If
            
         ElseIf Val(ano_factura$) = Val(ano_actual$) Then
            dia_fecha_actual$ = Format(fecha_final_recibo$, "y")
         
            dia_fecha_factura$ = Format(fecha_factura$, "y")
            total_dias = Val(dia_fecha_factura$)
            
            dias_que_han_pasado = (Val(dia_fecha_actual$) - total_dias)
            
         End If
         
         
         If expiracion_invoices = 0 Then
            If dias_que_han_pasado > 31 Then
                If balance_factura$ <> "" Then
                    
                    r$ = recibo$ + Space(1) + Format(Format(balance_factura$, "$#,##0.00"), "@@@@@@@@@") + Space(2) + Format(csr1$, "!@@@@@@@@@@@") + Space(1) + Format(user1$, "!@@@@@@@@@@@")
                    existe = 0
                    For g = 0 To lista_invoices30.ListCount - 1
                       If r$ = lista_invoices30.List(g) Then
                          existe = 1
                          Exit For
                       End If
                    Next g
                
                    If existe = 0 Then
                       lista_invoices30.AddItem recibo$ + Space(1) + Format(Format(balance_factura$, "$#,##0.00"), "@@@@@@@@@") + Space(2) + Format(csr1$, "!@@@@@@@@@@@") + Space(1) + Format(user1$, "!@@@@@@@@@@@")
                    End If
                    
                End If
            End If
         Else
         
            If dias_que_han_pasado > 61 Then
                If balance_factura$ <> "" Then
                    
                    r$ = recibo$ + Space(1) + Format(Format(balance_factura$, "$#,##0.00"), "@@@@@@@@@") + Space(2) + Format(csr1$, "!@@@@@@@@@@@") + Space(1) + Format(user1$, "!@@@@@@@@@@@")
                    existe = 0
                    For g = 0 To lista_invoices30.ListCount - 1
                       If r$ = lista_invoices30.List(g) Then
                          existe = 1
                          Exit For
                       End If
                    Next g
                
                    If existe = 0 Then
                       lista_invoices30.AddItem recibo$ + Space(1) + Format(Format(balance_factura$, "$#,##0.00"), "@@@@@@@@@") + Space(2) + Format(csr1$, "!@@@@@@@@@@@") + Space(1) + Format(user1$, "!@@@@@@@@@@@")
                    End If
                    
                End If
            End If
         End If
         
         dias_que_han_pasado = 0
      
   End If
  
  
Next t


End Sub

Public Sub calcula_grandtotal()
On Error Resume Next
lblmsg.Caption = "Please, wait a moment... calculating grand total"

msg.Visible = True
msg.Refresh

grandtotal = 0
'For t = 0 To 1  ' lista_agentes.ListCount - 1
   lista_agentes.ListIndex = 0
   
   grandtotal = grandtotal + Val(Format(lbltotal_invoices.Caption, "000000.00"))

'Next t
lblgrandtotal.Caption = Format(grandtotal, "$###,##0.00")

msg.Visible = False
msg.Refresh


End Sub

Public Sub asigna_limite()
On Error Resume Next

'BF_user = Val(Format(lbltotal_invoices.Caption, "000000.00"))


BF_user = Val(Format(lbltotal_bf.Caption, "000000.00"))
nb_user = Val(Format(lbltotal_NB.Caption, "00.0"))



    
   Dim sSelect As String
Dim Rs As ADODB.Recordset




' detecta si es de phone sales
existe = 0
For t = 0 To lista_PhoneSales.ListCount - 1
  a$ = Left(lista_PhoneSales.List(t), 20)
  If UCase(RTrim(a$)) = agente$ Then
     existe = 1
     Exit For
  End If
Next t
 
carga_tier_del_agente
b = categoria_tier
 
 
 
' =====================================================================================================
' ======   CARGA TIER de AGENTE INDEPENDIENTE
' ======


' revisa en que rango queda

 If categoria_tier = 9 Or categoria_tier = 17 Or categoria_tier = 18 Then
 
    Set Rs = New ADODB.Recordset
    
    NB_Goal1 = ""
    sSelect = "SELECT NBgoalmin From tierscatalog where tier='1' and (idjobtitle='2' or idjobtitle='37') and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    NB_Goal1 = Rs(0)
                        
    Rs.Close
    
    
    NB_Goal2 = ""
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT NBgoalmin From tierscatalog where tier='2' and (idjobtitle='2' or idjobtitle='37') and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    NB_Goal2 = Rs(0)
                        
    Rs.Close
    
    
    
    NB_Goal3 = ""
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT NBgoalmin From tierscatalog where tier='3' and (idjobtitle='2' or idjobtitle='37') and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    NB_Goal3 = Rs(0)
                        
    Rs.Close
    

    NB_Goal4 = ""
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT NBgoalmin From tierscatalog where tier='4' and (idjobtitle='2' or idjobtitle='37') and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    NB_Goal4 = Rs(0)
                        
    Rs.Close


    NB_Goal5 = ""
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT NBgoalmin From tierscatalog where tier='5' and (idjobtitle='2' or idjobtitle='37') and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    NB_Goal5 = Rs(0)
                        
    Rs.Close



    'commission1 = 50
    'commission2 = 13
    'commission3 = 15
    'commission4 = 20
    'commission5 = 25
    
    Set Rs = New ADODB.Recordset
    
    
    commission1 = 0
    sSelect = "SELECT PercComm, flatamount From tierscatalog where tier='1' and (idjobtitle='2' or idjobtitle='37') and idpaytype='" + Format(categoria_tier, "#0") + "'  and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    commission1 = Rs(0)
    flatamount = Rs(1)
    
                          
    Rs.Close
    
    If commission1 = 0 Then
       commission1 = flatamount
    End If
    
    
    commission2 = 0
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT PercComm From tierscatalog where tier='2' and (idjobtitle='2' or idjobtitle='37') and idpaytype='" + Format(categoria_tier, "#0") + "'  and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    commission2 = Rs(0)
                        
    Rs.Close
    
    
    commission3 = 0
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT PercComm From tierscatalog where tier='3' and (idjobtitle='2' or idjobtitle='37') and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    commission3 = Rs(0)
                        
    Rs.Close
    
    
    
    commission4 = 0
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT PercComm From tierscatalog where tier='4' and (idjobtitle='2' or idjobtitle='37') and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    commission4 = Rs(0)
                        
    Rs.Close
    
    
    commission5 = 0
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT PercComm From tierscatalog where tier='5' and (idjobtitle='2' or idjobtitle='37') and idpaytype='" + Format(categoria_tier, "#0") + "'  and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    commission5 = Rs(0)
                        
    Rs.Close



Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmin From tierscatalog where tier='1' and (idjobtitle='2' or idjobtitle='37') and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango1a = Rs(0)
                        
    Rs.Close


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmax From tierscatalog where tier='1' and (idjobtitle='2' or idjobtitle='37') and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango1b = Rs(0)
                        
    Rs.Close



    Salary = 20


    'rango1a = 1350  '1250
    'rango1b = 1499  '1349
    Tablax_BF(0).Caption = Format(rango1a, "$###,##0") + " - " + Format(rango1b, "$###,##0")
    tablax_NB(0).Caption = Str(NB_Goal1)
    
    If commission1 = 50 Then
      tablax_comm(0).Caption = Format(commission1, "$##0")
    Else
      tablax_comm(0).Caption = Format(commission1, "00") + "%"
    End If


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmin From tierscatalog where tier='2' and (idjobtitle='2' or idjobtitle='37') and idpaytype='" + Format(categoria_tier, "#0") + "'  and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango2a = Rs(0)
                        
    Rs.Close


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmax From tierscatalog where tier='2' and (idjobtitle='2' or idjobtitle='37') and idpaytype='" + Format(categoria_tier, "#0") + "'  and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango2b = Rs(0)
                        
    Rs.Close


    'rango2a = 1500  '1350
    'rango2b = 1900  '1801  '1800
    Tablax_BF(1).Caption = Format(rango2a, "$###,##0") + " - " + Format(rango2b, "$###,##0")
    tablax_NB(1).Caption = Str(NB_Goal2)
    tablax_comm(1).Caption = Format(commission2, "00") + "%"


Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmin From tierscatalog where tier='3' and (idjobtitle='2' or idjobtitle='37') and idpaytype='" + Format(categoria_tier, "#0") + "'  and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango3a = Rs(0)
                        
    Rs.Close


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmax From tierscatalog where tier='3' and (idjobtitle='2' or idjobtitle='37') and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango3b = Rs(0)
                        
    Rs.Close



    'rango3a = 1901 '1801
    ' rango3b = 2200 '2101 '2100
    Tablax_BF(2).Caption = Format(rango3a, "$###,##0") + " - " + Format(rango3b, "$###,##0")
    tablax_NB(2).Caption = Str(NB_Goal3)
    tablax_comm(2).Caption = Format(commission3, "00") + "%"


    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmin From tierscatalog where tier='4' and (idjobtitle='2' or idjobtitle='37') and idpaytype='" + Format(categoria_tier, "#0") + "'  and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango4a = Rs(0)
                        
    Rs.Close


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmax From tierscatalog where tier='4' and (idjobtitle='2' or idjobtitle='37') and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango4b = Rs(0)
                        
    Rs.Close
    
    If rango4b > 999999 Then
       rango4b = 9999
    End If

    'rango4a = 2201  '2101
    'rango4b = 3100  '2999  '2998
    Tablax_BF(3).Caption = Format(rango4a, "$###,##0") + " - " + Format(rango4b, "$###,##0")
    tablax_NB(3).Caption = Str(NB_Goal4)
    tablax_comm(3).Caption = Format(commission4, "00") + "%"
    
    


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmin From tierscatalog where tier='5' and (idjobtitle='2' or idjobtitle='37') and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango5a = Rs(0)
                        
    Rs.Close


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmax From tierscatalog where tier='5' and (idjobtitle='2' or idjobtitle='37') and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango5b = Rs(0)
                        
    Rs.Close


    If rango5b > 999999 Then
       rango5b = 9999
    End If

    'rango5a = 3101  '2999
    'rango5b = 9999
    Tablax_BF(4).Caption = Format(rango5a, "$###,##0") + " - " + Format(rango5b, "$###,##0")
    tablax_NB(4).Caption = Str(NB_Goal5)
    tablax_comm(4).Caption = Format(commission5, "00") + "%"
    
    If NB_Goal5 > 0 Then
    
    Tablax_BF(4).Visible = True
    tablax_NB(4).Visible = True
    tablax_comm(4).Visible = True
    tablax_salary(4).Visible = True
    'tablax_DED_DMV(4).Visible = True
    
    Tablax_BF(5).Visible = True
    tablax_NB(5).Visible = True
    tablax_comm(5).Visible = True
    tablax_salary(5).Visible = True
    
    Else
    
    Tablax_BF(4).Visible = False
    tablax_NB(4).Visible = False
    tablax_comm(4).Visible = False
    tablax_salary(4).Visible = False
    'tablax_DED_DMV(4).Visible = false
    
    
    End If
    
    
    GoTo continua_aqui

 End If

' =============================================================================================
' ===========   TERMINA AQUI





 
If existe = 0 Then

' revisa en que rango queda

   ' NB_Goal1 = 6 '5
   ' NB_Goal2 = 9 '8
   ' NB_Goal3 = 11 '10
   ' NB_Goal4 = 13 '12
   ' NB_Goal5 = 15 '14

Set Rs = New ADODB.Recordset
    sSelect = "SELECT NBgoalmin From tierscatalog where tier='1'  and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"    ' SE removio esto:  and idjobtitle='16'
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    NB_Goal1 = Rs(0)
                        
    Rs.Close
    
    
    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT NBgoalmin From tierscatalog where tier='2'  and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    NB_Goal2 = Rs(0)
                        
    Rs.Close
    
    
    
    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT NBgoalmin From tierscatalog where tier='3' and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    NB_Goal3 = Rs(0)
                        
    Rs.Close
    


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT NBgoalmin From tierscatalog where tier='4' and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    NB_Goal4 = Rs(0)
                        
    Rs.Close



    Set Rs = New ADODB.Recordset
    sSelect = "SELECT NBgoalmin From tierscatalog where tier='5' and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    NB_Goal5 = Rs(0)
                        
    Rs.Close



    'commission1 = 50
    'commission2 = 13
    'commission3 = 15
    'commission4 = 20
    'commission5 = 25
    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT PercComm, flatamount From tierscatalog where tier='1' and idpaytype='" + Format(categoria_tier, "#0") + "'  and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    commission1 = Rs(0)
    flatamount = Rs(1)
    
                          
    Rs.Close
    
    If commission1 = 0 Then
       commission1 = flatamount
    End If
    
    
    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT PercComm From tierscatalog where tier='2' and idpaytype='" + Format(categoria_tier, "#0") + "'  and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    commission2 = Rs(0)
                        
    Rs.Close
    
    
    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT PercComm From tierscatalog where tier='3'  and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    commission3 = Rs(0)
                        
    Rs.Close
    
    
    
    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT PercComm From tierscatalog where tier='4'  and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    commission4 = Rs(0)
                        
    Rs.Close
    
    
    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT PercComm From tierscatalog where tier='5'  and idpaytype='" + Format(categoria_tier, "#0") + "'  and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    commission5 = Rs(0)
                        
    Rs.Close



Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmin From tierscatalog where tier='1'  and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango1a = Rs(0)
                        
    Rs.Close


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmax From tierscatalog where tier='1'  and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango1b = Rs(0)
                        
    Rs.Close



    Salary = 20


    'rango1a = 1350  '1250
    'rango1b = 1499  '1349
    Tablax_BF(0).Caption = Format(rango1a, "$###,##0") + " - " + Format(rango1b, "$###,##0")
    tablax_NB(0).Caption = Str(NB_Goal1)
    
    If commission1 = 50 Then
      tablax_comm(0).Caption = Format(commission1, "$##0")
    Else
      tablax_comm(0).Caption = Format(commission1, "00") + "%"
    End If


    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmin From tierscatalog where tier='2'  and idpaytype='" + Format(categoria_tier, "#0") + "'  and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango2a = Rs(0)
                        
    Rs.Close


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmax From tierscatalog where tier='2'  and idpaytype='" + Format(categoria_tier, "#0") + "'  and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango2b = Rs(0)
                        
    Rs.Close


    'rango2a = 1500  '1350
    'rango2b = 1900  '1801  '1800
    Tablax_BF(1).Caption = Format(rango2a, "$###,##0") + " - " + Format(rango2b, "$###,##0")
    tablax_NB(1).Caption = Str(NB_Goal2)
    tablax_comm(1).Caption = Format(commission2, "00") + "%"


Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmin From tierscatalog where tier='3'  and idpaytype='" + Format(categoria_tier, "#0") + "'  and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango3a = Rs(0)
                        
    Rs.Close


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmax From tierscatalog where tier='3'  and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango3b = Rs(0)
                        
    Rs.Close



    'rango3a = 1901 '1801
    ' rango3b = 2200 '2101 '2100
    Tablax_BF(2).Caption = Format(rango3a, "$###,##0") + " - " + Format(rango3b, "$###,##0")
    tablax_NB(2).Caption = Str(NB_Goal3)
    tablax_comm(2).Caption = Format(commission3, "00") + "%"


    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmin From tierscatalog where tier='4'  and idpaytype='" + Format(categoria_tier, "#0") + "'  and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango4a = Rs(0)
                        
    Rs.Close


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmax From tierscatalog where tier='4'  and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango4b = Rs(0)
                        
    Rs.Close
    
    If rango4b > 999999 Then
       rango4b = 9999
    End If

    'rango4a = 2201  '2101
    'rango4b = 3100  '2999  '2998
    Tablax_BF(3).Caption = Format(rango4a, "$###,##0") + " - " + Format(rango4b, "$###,##0")
    tablax_NB(3).Caption = Str(NB_Goal4)
    tablax_comm(3).Caption = Format(commission4, "00") + "%"
    
    


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmin From tierscatalog where tier='5'  and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango5a = Rs(0)
                        
    Rs.Close


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmax From tierscatalog where tier='5'  and idpaytype='" + Format(categoria_tier, "#0") + "' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango5b = Rs(0)
                        
    Rs.Close


    If rango5b > 999999 Then
       rango5b = 9999
    End If

    'rango5a = 3101  '2999
    'rango5b = 9999
    Tablax_BF(4).Caption = Format(rango5a, "$###,##0") + " - " + Format(rango5b, "$###,##0")
    tablax_NB(4).Caption = Str(NB_Goal5)
    tablax_comm(4).Caption = Format(commission5, "00") + "%"
    
    If NB_Goal5 > 0 Then
    
    For w = 0 To 4
      Tablax_BF(w).Visible = True
      tablax_NB(w).Visible = True
      tablax_comm(w).Visible = True
      tablax_salary(w).Visible = True
    Next w
    'tablax_DED_DMV(4).Visible = True
    
    Tablax_BF(5).Visible = True
    tablax_NB(5).Visible = True
    tablax_comm(5).Visible = True
    tablax_salary(5).Visible = True
    
    
    Else
    
      For w = 0 To 3
        Tablax_BF(w).Visible = True
        tablax_NB(w).Visible = True
        tablax_comm(w).Visible = True
        tablax_salary(w).Visible = True
        'tablax_DED_DMV(4).Visible = false
      Next w
    
       Tablax_BF(4).Visible = False
        tablax_NB(4).Visible = False
        tablax_comm(4).Visible = False
        tablax_salary(4).Visible = False
        
        Tablax_BF(5).Visible = True
        tablax_NB(5).Visible = True
        tablax_comm(5).Visible = True
        tablax_salary(5).Visible = True
    
    
    End If

Else


    'NB_Goal1 = 5   ' estaba 10
    'NB_Goal2 = 8   ' estaba 12
    'NB_Goal3 = 12  ' estaba 15
    'NB_Goal4 = 15  ' estaba 18
    
    
    'GoTo salta_sql1
    
    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT NBgoalmin From tierscatalog where tier='1' and idjobtitle='2' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    NB_Goal1 = Rs(0)
                        
    Rs.Close
    
    
    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT NBgoalmin From tierscatalog where tier='2' and idjobtitle='2' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    NB_Goal2 = Rs(0)
                        
    Rs.Close
    
    
    
    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT NBgoalmin From tierscatalog where tier='3' and idjobtitle='2' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    NB_Goal3 = Rs(0)
                        
    Rs.Close
    


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT NBgoalmin From tierscatalog where tier='4' and idjobtitle='2' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    NB_Goal4 = Rs(0)
                        
    Rs.Close




    Set Rs = New ADODB.Recordset
    sSelect = "SELECT NBgoalmin From tierscatalog where tier='5' and idjobtitle='2' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    NB_Goal5 = Rs(0)
                        
    Rs.Close
    

salta_sql1:

    

    'commission1 = 50
    'commission2 = 10
    'commission3 = 13
    'commission4 = 15
    
    'GoTo salta_sql2
    
    
        
    
     Set Rs = New ADODB.Recordset
    sSelect = "SELECT PercComm, flatamount From tierscatalog where tier='1' and idjobtitle='2'  and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    commission1 = Rs(0)
    flatamount = Rs(1)
    
                          
    Rs.Close
    
    If commission1 = 0 Then
       commission1 = flatamount
    End If
    
    
    
    
    
    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT PercComm From tierscatalog where tier='2' and idjobtitle='2' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    commission2 = Rs(0)
                        
    Rs.Close
    
    
    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT PercComm From tierscatalog where tier='3' and idjobtitle='2' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    commission3 = Rs(0)
                        
    Rs.Close
    
    
    
    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT PercComm From tierscatalog where tier='4' and idjobtitle='2' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    commission4 = Rs(0)
                        
    Rs.Close
    
    
    
    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT PercComm From tierscatalog where tier='5' and idjobtitle='2' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    commission5 = Rs(0)
                        
    Rs.Close



Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmin From tierscatalog where tier='1' and idjobtitle='2'  and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango1a = Rs(0)
                        
    Rs.Close


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmax From tierscatalog where tier='1' and idjobtitle='2' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango1b = Rs(0)
                        
    Rs.Close


    
salta_sql2:
    
    Salary = 20
    
    

    'rango1a = 1050
    'rango1b = 1200  '1199
    Tablax_BF(0).Caption = Format(rango1a, "$###,##0") + " - " + Format(rango1b, "$###,##0")
    tablax_NB(0).Caption = Str(NB_Goal1)
       
    If commission1 = 50 Then
      tablax_comm(0).Caption = Format(commission1, "$##0")
    Else
      tablax_comm(0).Caption = Format(commission1, "00") + "%"
    End If

    'GoTo salta_sql3



Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmin From tierscatalog where tier='2' and idjobtitle='2'  and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango2a = Rs(0)
                        
    Rs.Close


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmax From tierscatalog where tier='2' and idjobtitle='2' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango2b = Rs(0)
                        
    Rs.Close


salta_sql3:

    'rango2a = 1200
    'rango2b = 1501  '1500
    Tablax_BF(1).Caption = Format(rango2a, "$###,##0") + " - " + Format(rango2b, "$###,##0")
    tablax_NB(1).Caption = Str(NB_Goal2)
    tablax_comm(1).Caption = Format(commission2, "00") + "%"


'GoTo salta_sql4

Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmin From tierscatalog where tier='3' and idjobtitle='2' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango3a = Rs(0)
                        
    Rs.Close


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmax From tierscatalog where tier='3' and idjobtitle='2' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango3b = Rs(0)
                        
    Rs.Close
    
    
salta_sql4:
    
    
    'rango3a = 1501
    'rango3b = 2001 '2000
    Tablax_BF(2).Caption = Format(rango3a, "$###,##0") + " - " + Format(rango3b, "$###,##0")
    tablax_NB(2).Caption = Str(NB_Goal3)
    tablax_comm(2).Caption = Format(commission3, "00") + "%"

 '   GoTo salta_sql5


Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmin From tierscatalog where tier='4' and idjobtitle='2'  and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango4a = Rs(0)
                        
    Rs.Close


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmax From tierscatalog where tier='4' and idjobtitle='2' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango4b = Rs(0)
                        
    Rs.Close
    
    If rango4b > 999999 Then
       rango4b = 9999
    End If
    
salta_sql5:
    
    
    'rango4a = 2001
    'rango4b = 9999
    Tablax_BF(3).Caption = Format(rango4a, "$###,##0") + " - " + Format(rango4b, "$###,##0")
    tablax_NB(3).Caption = Str(NB_Goal4)
    tablax_comm(3).Caption = Format(commission4, "00") + "%"



Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmin From tierscatalog where tier='5' and idjobtitle='2' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango5a = Rs(0)
                        
    Rs.Close


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmax From tierscatalog where tier='5' and idjobtitle='2' and active='1'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango5b = Rs(0)
                        
    Rs.Close
    
    If rango5b > 999999 Then
       rango5b = 9999
    End If
    
    
    
    'rango4a = 2001
    'rango4b = 9999
    Tablax_BF(4).Caption = Format(rango5a, "$###,##0") + " - " + Format(rango5b, "$###,##0")
    tablax_NB(4).Caption = Str(NB_Goal5)
    tablax_comm(4).Caption = Format(commission5, "00") + "%"


    
    If NB_Goal5 > 0 Then
      For w = 0 To 4
         Tablax_BF(w).Visible = True
         tablax_NB(w).Visible = True
         tablax_comm(w).Visible = True
         tablax_salary(w).Visible = True
      Next w
     
    Else
    
      For w = 0 To 3
        Tablax_BF(w).Visible = True
        tablax_NB(w).Visible = True
        tablax_comm(w).Visible = True
        tablax_salary(w).Visible = True
       'tablax_DED_DMV(4).Visible = False
      Next w
      
        Tablax_BF(4).Visible = False
        tablax_NB(4).Visible = False
        tablax_comm(4).Visible = False
        tablax_salary(4).Visible = False
    
    End If

End If





continua_aqui:

' ****************************************************************************************************************************************
' ****************************************************************************************************************************************
' ****************************************************************************************************************************************
' ****************************************************************************************************************************************



bf_final = 0



If BF_user >= 11000 Then
  bf_final = 6
ElseIf BF_user >= rango5a And BF_user <= rango5b Then
  bf_final = 5
ElseIf BF_user >= rango4a And BF_user <= rango4b Then
  bf_final = 4
ElseIf BF_user >= rango3a And BF_user <= rango3b Then
  bf_final = 3
ElseIf BF_user >= rango2a And BF_user <= rango2b Then
  bf_final = 2
ElseIf BF_user >= rango1a And BF_user <= rango1b Then
  bf_final = 1


Else
  bf_final = 0
End If


 ' SI es INDEPENDIENTE
 If categoria_tier = 9 Or categoria_tier = 17 Or categoria_tier = 18 Then
     bf_final = 1
 End If
 
 



If categoria_tier = 1 Then
  
  tablax_NB(5) = "$12,500"
ElseIf categoria_tier = 8 Then
  tablax_NB(5) = "$9,000"
  
Else
  tablax_NB(5) = "---"
End If





nb_final = 0
If nb_user >= NB_Goal5 And NB_Goal5 <> "" Then
  nb_final = 5
ElseIf nb_user >= NB_Goal4 Then ' tier4
  nb_final = 4

ElseIf nb_user >= NB_Goal3 Then 'tier3
  nb_final = 3

ElseIf nb_user >= NB_Goal2 Then 'tier2
  nb_final = 2

ElseIf nb_user >= NB_Goal1 Then 'tier1
  nb_final = 1
  
ElseIf nb_user >= 6 And nb_user <= 8.5 And BF_user >= 2701 And categoria_tier = 1 Then  'tier especial FULL TIME
  nb_final = 9
  bf_final = 9
  
  
ElseIf nb_user >= 3 And nb_user <= 3.5 And BF_user >= 2101 And categoria_tier = 8 Then 'tier especial PART TIME
  nb_final = 8
  bf_final = 8
  
  
Else
  nb_final = 0

End If


 ' SI es INDEPENDIENTE
 If categoria_tier = 9 Or categoria_tier = 17 Or categoria_tier = 18 Then
   nb_final = 1
 End If



  For t = 0 To 5
    img_arrow(t).Visible = False
    img_arrow2(t).Visible = False
  Next t

penal = 0





If bf_final = 1 And nb_final = 1 Then       ' ANTERIOR = If bf_final = 1 or (nb_final = 1 And bf_final >= 2) Then
  img_arrow(0).Visible = True
   penal = 0
  
  'If commission1 = 50 Then
  '   commis = Val(commission1)
  '   porcentaje_comm$ = ""
  'Else
  
     X = (Val(Format(lbltotal_invoices.Caption, "000000.00")) * (commission1 - penal)) / 100
     commis = Format(X, "$##0.00")
     porcentaje_comm$ = Format(commission1 - penal, "00") + "%"
     
  'End If
  
ElseIf bf_final = 9 And nb_final = 9 Then  ' CATEGORIA ESPECIAL tiempo completo
  
  img_arrow(4).Visible = True
  penal = 0
  
  X = (Val(Format(lbltotal_invoices.Caption, "000000.00")) * (6)) / 100
  commis = Format(X, "$##0.00")
  porcentaje_comm$ = "6%"
  
  
ElseIf bf_final = 8 And nb_final = 8 Then  ' CATEGORIA ESPECIAL PART TIME
  
  img_arrow(4).Visible = True
  penal = 0
  
  X = (Val(Format(lbltotal_invoices.Caption, "000000.00")) * (6)) / 100
  commis = Format(X, "$##0.00")
  porcentaje_comm$ = "6%"
  
  
  
ElseIf bf_final = 1 And nb_final >= 1 And nb_user >= Val(tablax_NB(0).Caption) Then
   img_arrow(0).Visible = True
   penal = 0
  
   X = (Val(Format(lbltotal_invoices.Caption, "000000.00")) * (commission1 - penal)) / 100
  commis = Format(X, "$##0.00")
  porcentaje_comm$ = Format(commission1 - penal, "00") + "%"
  
  
  
  
ElseIf bf_final = 2 And nb_final >= 2 And nb_user >= Val(tablax_NB(0).Caption) Then     ' val(tablax_NB(1).Caption) estaba un 8 en de lugar de
  img_arrow(1).Visible = True
  penal = 0
  X = (Val(Format(lbltotal_invoices.Caption, "000000.00")) * (commission2 - penal)) / 100
  commis = Format(X, "$##0.00")
  porcentaje_comm$ = Format(commission2 - penal, "00") + "%"
  
ElseIf bf_final = 2 And nb_final < 2 And nb_user >= Val(tablax_NB(0).Caption) Then
  img_arrow2(0).Visible = True
  penal = 2
  'If commission1 = 50 Then
  '   commis = Val(commission1)
  '   porcentaje_comm$ = ""
  'Else
  
     X = (Val(Format(lbltotal_invoices.Caption, "000000.00")) * (commission2 - penal)) / 100
     commis = Format(X, "$##0.00")
     porcentaje_comm$ = Format(commission2 - penal, "00") + "%"
     
  'End If
  
  
ElseIf bf_final = 3 And nb_final >= 3 And nb_user >= Val(tablax_NB(0).Caption) Then
  img_arrow(2).Visible = True
  penal = 0
  X = (Val(Format(lbltotal_invoices.Caption, "000000.00")) * (commission3 - penal)) / 100
  commis = Format(X, "$##0.00")
  porcentaje_comm$ = Format(commission3 - penal, "00") + "%"
  
  
ElseIf bf_final = 3 And nb_final < 3 And nb_user >= Val(tablax_NB(0).Caption) Then
  img_arrow2(1).Visible = True
  penal = 2
  X = (Val(Format(lbltotal_invoices.Caption, "000000.00")) * (commission3 - penal)) / 100
  commis = Format(X, "$##0.00")
  porcentaje_comm$ = Format(commission3 - penal, "00") + "%"

  
ElseIf bf_final = 4 And nb_final >= 4 And nb_user >= Val(tablax_NB(0).Caption) Then
  img_arrow(3).Visible = True
  penal = 0
  X = (Val(Format(lbltotal_invoices.Caption, "000000.00")) * (commission4 - penal)) / 100
  commis = Format(X, "$##0.00")
  porcentaje_comm$ = Format(commission4 - penal, "00") + "%"

ElseIf bf_final = 4 And nb_final < 4 And nb_user >= Val(tablax_NB(0).Caption) Then
  img_arrow2(2).Visible = True
  penal = 2
  X = (Val(Format(lbltotal_invoices.Caption, "000000.00")) * (commission4 - penal)) / 100
  commis = Format(X, "$##0.00")
  porcentaje_comm$ = Format(commission4 - penal, "00") + "%"

  
ElseIf bf_final = 5 And nb_final >= 5 And nb_user >= Val(tablax_NB(0).Caption) Then
  img_arrow(4).Visible = True
  penal = 0
  X = (Val(Format(lbltotal_invoices.Caption, "000000.00")) * (commission5 - penal)) / 100
  commis = Format(X, "$##0.00")
  porcentaje_comm$ = Format(commission5 - penal, "00") + "%"

ElseIf bf_final = 5 And nb_final < 5 And nb_user >= Val(tablax_NB(0).Caption) Then
  img_arrow2(3).Visible = True
  penal = 2
  X = (Val(Format(lbltotal_invoices.Caption, "000000.00")) * (commission5 - penal)) / 100
  commis = Format(X, "$##0.00")
  porcentaje_comm$ = Format(commission5 - penal, "00") + "%"

ElseIf bf_final = 6 And nb_user >= Val(tablax_NB(0).Caption) Then
  img_arrow(5).Visible = True
  penal = 0
  X = ((Val(Format(lbltotal_invoices.Caption, "000000.00")) * (commission5 - penal)) / 100) + 500
  commis = Format(X, "$##0.00")
  porcentaje_comm$ = Format(commission5 - penal, "00") + "%"

End If


If nb_final = 0 Then
  For t = 0 To 5
    img_arrow(t).Visible = False
    img_arrow2(t).Visible = False
  Next t
  penal = 0
  commis = 0
  porcentaje_comm$ = ""
  
End If


lblcommission.Caption = Format(commis, "$###,##0.00")

If chklock3.Value = 0 Then
  lblporcentaje.Caption = porcentaje_comm$

Else
  ' de lo contrario si es modificado la comision por autorizacion del manager
  For t = 0 To 5
    img_arrow(t).Visible = False
    img_arrow2(t).Visible = False
  Next t
  
  
  
End If




' SI es INDEPENDIENTE
 If categoria_tier = 9 Or categoria_tier = 17 Or categoria_tier = 18 Then
  For t = 1 To 5
    img_arrow(t).Visible = False
    img_arrow2(t).Visible = False
    
     Tablax_BF(t).Visible = False
     tablax_NB(t).Visible = False
     tablax_comm(t).Visible = False
     tablax_salary(t).Visible = False
      
  Next t
  
    Tablax_BF(0).Caption = "$1 - $99,999"
    tablax_NB(0).Caption = "-"
    tablax_salary(0).Caption = ""
 
 Else
 
    tablax_salary(0).Caption = "20 x hour"
      
  
 End If





End Sub

Public Sub carga_iniciales()
On Error Resume Next

Dim sSelect As String
Dim Rs As ADODB.Recordset
    
Set Rs = New ADODB.Recordset



lblinitials.Caption = ""

    sSelect = "Select idemppayrolllink, idemployeeLAE, idpaytype, initials, manager, MTY, Commercial, PhoneSales, Exception1,Typeofexception1," & _
    "Exception2,Typeofexception2,Exception3,Typeofexception3,Exception4,Typeofexception4,Exception5,Typeofexception5,deb_mod, nb_deb_mod, porc_mod, autorizado, " & _
    "office1, office2, emp.firstname, emp.lastname1, emp.username from [payrollconfig] payroll " & _
    "inner join EmployeeInfo emp on emp.IDEmployee=payroll.IdEmployeeLAE "
    
' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
        
     ' Permitir redimensionar las columnas
    grid.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid.DataSource = Rs
                         
    Rs.Close
    
    


For t = 1 To grid.Rows - 1

  grid.row = t
  grid.col = 27
  username1$ = grid.Text
  
  grid.col = 4
  iniciales1$ = grid.Text
  
  grid.col = 6
  monterrey1$ = grid.Text
  
  
  grid.col = 25
  nombre1$ = grid.Text
  
  grid.col = 26
  apellido1$ = grid.Text
  
  nombre_completo1$ = nombre1$ + " " + apellido1$
  
  If UCase(RTrim(username1$)) = UCase(agente$) Then
     lblinitials.Caption = iniciales1$
     lblfull_name.Caption = UCase(nombre_completo1$)
     
     If Val(monterrey1$) = 1 Then
       icon_mex.Visible = True
     Else
       icon_mex.Visible = False
     End If


    grid.col = 23
    ofic1$ = grid.Text

     sSelect = "select office from officescatalog where idoffice='" + ofic1$ + "'"
     Rs.Open sSelect, base, adOpenUnspecified
     oficina1$ = Rs(0)
     Rs.Close
         
     
     lbloficina.Caption = oficina1$
     
     grid.col = 1
     id1$ = grid.Text
     
     If Val(id1$) > 0 Then
       lblidentificacion.Caption = id1$
     Else
       lblidentificacion.Caption = ""
     End If
     
     
     grid.col = 2
     id_LAE1$ = grid.Text
     
     If Val(id_LAE1$) > 0 Then
       lbllae.Caption = id_LAE1$
     Else
       lbllae.Caption = ""
     End If
     
     
        grid.col = 19
        ded_mod1$ = grid.Text
     
        btndeduction.Visible = Val(ded_mod1$)
        
        grid.col = 20
        nb_ded_mod1$ = grid.Text
               
        btnNB_deduction.Visible = Val(nb_ded_mod1$)
        
        grid.col = 21
        porc_mod1$ = grid.Text
        
        btnporcentaje.Visible = Val(porc_mod1$)
        
        
        
        'chklock1.Enabled = attr.ded_mod
        chklock1.Value = Val(ded_mod1$)
        
        'chklock2.Enabled = attr.nb_ded_mod
        chklock2.Value = Val(nb_ded_mod1$)
    
        'chklock3.Enabled = attr.porc_mod
        chklock3.Value = Val(porc_mod1$)

     
     Exit For
  End If
Next t






End Sub





Private Sub txtdatefrom_Click()
calen = 0
If txtdatefrom.Text = "" Then
   Calendar1.Today
Else
   Calendar1.Value = txtdatefrom.Text
End If
Calendar1.Visible = True
End Sub

Private Sub txtdatefrom_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then Exit Sub
If KeyAscii = Asc("/") Then Exit Sub

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
  Exit Sub
End If
End Sub

Private Sub txtdatefrom_LostFocus()
Calendar1.Visible = False
End Sub

Private Sub txtdatepayday_Click()
calen = 2
If txtdatepayday.Text = "" Then
   Calendar1.Today
Else
   Calendar1.Value = txtdatepayday.Text
End If

Calendar1.Visible = True

End Sub

Private Sub txtdatepayday_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then Exit Sub
If KeyAscii = Asc("/") Then Exit Sub

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
  Exit Sub
End If

End Sub


Private Sub txtdatepayday_LostFocus()
Calendar1.Visible = False

End Sub

Private Sub txtdateto_Click()
calen = 1
If txtdateto.Text = "" Then
   Calendar1.Today
Else
   Calendar1.Value = txtdateto.Text
End If
Calendar1.Visible = True

End Sub

Private Sub txtdateto_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then Exit Sub
If KeyAscii = Asc("/") Then Exit Sub

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
  Exit Sub
End If
End Sub


Private Sub txtdateto_LostFocus()
Calendar1.Visible = False

End Sub

Private Sub txtnotes_Click()
On Error Resume Next
If lista_agentes.ListCount = 0 Then
  Exit Sub
End If


If cargado = True Then
   Exit Sub
End If

Load forma_nota
forma_nota.Show 1
cargado = True
Timer1.Enabled = True
End Sub

Private Sub txtnotes_GotFocus()
On Error Resume Next
If lista_agentes.ListCount = 0 Then
  Exit Sub
End If


If cargado = True Then
   Exit Sub
End If

Load forma_nota
forma_nota.Show 1
cargado = True
Timer1.Enabled = True

End Sub

Private Sub txtnotes_LostFocus()
On Error Resume Next


   notes(num_fila_agente) = RTrim(txtnotes.Text)
   
End Sub





Public Sub carga_ubicaciones()
On Error Resume Next


Dim sSelect As String
   
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset

    grid.Clear
    
    sSelect = "Select idemployeeLAE, office1, office2, emp.username from [payrollconfig] payroll " & _
    "inner join EmployeeInfo emp on emp.IDEmployee=payroll.IdEmployeeLAE "
    
    '"where emp.username='" + txtagent.Text + "'"


     

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
        
     ' Permitir redimensionar las columnas
    grid.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid.DataSource = Rs
                         
    Rs.Close
    
Erase ubicacion

   
For t = 1 To grid.Rows - 1
   grid.row = t
   grid.col = 2
   oficina1$ = grid.Text
   
   
   grid.col = 3
   oficina2$ = grid.Text
   
   grid.col = 4
   UserName$ = grid.Text
   
   
   'oficina1$ = ""
   'oficina2$ = "empty" + Format(t, "00")
   'If UCase(UserName$) = "JDIAZ" Then Stop
     
   
   ubicacion(t, 0) = UCase(UserName$)
   
        
    sSelect = "Select office from officescatalog where idoffice='" + oficina1$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    oficina1$ = Rs(0)
    Rs.Close
   
           
   ubicacion(t, 1) = UCase(oficina1$)
     
     
    sSelect = "Select office from officescatalog where idoffice='" + oficina2$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    oficina2$ = Rs(0)
    Rs.Close
    
    If oficina2$ = "-1" Or oficina2$ = "0" Then
       oficina2$ = ""
    End If
    
   ubicacion(t, 2) = UCase(oficina2$)
     
Next t


ubicacion(0, 0) = grid.Rows - 1



End Sub





Public Sub carga_NB()
On Error Resume Next


agente$ = lista_agentes.List(lista_agentes.ListIndex)


' ************************************************************
  
 Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
      Dim rsVar As Variant
   Dim i As Integer
    
    Set Rs = New ADODB.Recordset
    'Checa_status
   
    user1$ = ""
    'sSelect = "SELECT login From employees where login='hnavarro'"
    
    
    sSelect = "SELECT " & _
    "recdtl.[IdReceiptHDR] as [Receipt #] " & _
    ",rechdr.Date " & _
    ",iitem.InvoiceItemDesc as [Invoice Item] " & _
    ",emp.Username as [User] " & _
    ",csr.Username as [CSR] " & _
    ",recdtl.Amount " & _
    ",ofc.Office " & _
    ",rechdr.IdOffice " & _
    "FROM  [ReceiptsDTL] recdtl " & _
    "inner join  ReceiptsHDR  rechdr on recdtl.IdReceiptHDR=rechdr.IDReceiptHDR " & _
    "inner join PoliciesHDR polhdr on polhdr.IdPoliciesHDR=rechdr.IdPoliciesHDR " & _
    "inner join EmployeeInfo emp on emp.IDEmployee=rechdr.IdEmployeeUSR " & _
    "inner join EmployeeInfo csr on csr.IDEmployee=rechdr.IdEmployeeCSR1 " & _
    "inner join OfficesCatalog ofc on ofc.IdOffice=rechdr.IdOffice " & _
    "inner join OfficesCatalog ofcpol on ofcpol.IdOffice=rechdr.IdOffice " & _
    "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
    "inner join Customers cus on cus.IdCustomer=polhdr.IdCustomer " & _
    "inner join InvoiceItemCatalog iitem on iitem.IdInvoiceItem=recdtl.IdInvoiceItem " & _
    "inner join BankAccountCatalog bank on bank.IdBankAccount=iitem.IdBankAccount " & _
    "inner join CompanysAGICatalog ja on ja.IdCompanyAGI=polhdr.IdCompanyAGI " & _
    "where iitem.IdInvoiceItem in (1,6,20) " & _
    "and cast(rechdr.Date as Date) >= '" + txtdatefrom.Text + "' " & _
    "AND cast( rechdr.DATE as Date) <= '" + txtdateto.Text + "' " & _
    "and rechdr.Active=1 order by [Receipt #]"
    
    
    
    ' ---------------------------------------------------------------------------
    
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
   Rs.Open sSelect, base, adOpenStatic, adLockOptimistic
    
    
   Rs.MoveLast

   Rs.MoveFirst
   ' Assuming that rs is your ADO recordset
   Grid1.Rows = Rs.RecordCount + 1

   rsVar = Rs.GetString(adClipString, Rs.RecordCount)

   Grid1.cols = Rs.Fields.Count + 1
    
    
    
   Grid1.TextMatrix(0, 0) = ""
   ' Set column names in the grid
   For i = 0 To Rs.Fields.Count - 1
      Grid1.TextMatrix(0, i + 1) = Rs.Fields(i).name
   Next

   Grid1.row = 1
   Grid1.col = 1

   ' Set range of cells in the grid
   Grid1.RowSel = Grid1.Rows - 1
   Grid1.ColSel = Grid1.cols - 1
   Grid1.clip = rsVar

   ' Reset the grid's selected range of cells
   Grid1.RowSel = Grid1.row
   Grid1.ColSel = Grid1.col

   Rs.Close

   Set Rs = Nothing


' ----------------------------------------------------------------------------



    
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
'    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
 '   grid1.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
  '  Set grid1.DataSource = Rs
    ' user1$ = Rs(0)
          
                        
                         
  '  Rs.Close
    
    
    For t = 1 To Grid1.Rows - 1
       Grid1.row = t
       Grid1.col = 0
       Grid1.Text = t
    Next t
    
    
    
  setup_grid1

  CARGA_AGENTES

'carga_attributos

  
End Sub

Public Sub carga_GI()
On Error Resume Next

' ************************************************************
  
 Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    Dim rsVar As Variant
   Dim i As Integer
   
    
    Set Rs = New ADODB.Recordset
    'Checa_status
   
    user1$ = ""
   
    
    If agente$ <> "" Then  ' And agente$ <> "JENNIFERR" Then
    
      r$ = " and (emp.username='" + agente$ + "' or csr.username='" + agente$ + "' or csr2.username='" + agente$ + "') "
      
    '  r$ = "and goals.IdEmployee='" + lbllae.Caption + "' "
      
   ' ElseIf agente$ = "JENNIFERR" Then
    
      'r$ = " and (emp.username='" + agente$ + "' or csr.username='" + agente$ + "' or csr2.username='" + agente$ + "')  or ( iitem.InvoiceItemName='Invoice' )  "
    
   
      
    Else
    
      r$ = " "
    
    End If
    
    
    Grid2.Visible = False
    
    sSelect = "SELECT " & _
    "recdtl.[IdReceiptHDR] as [Receipt #] " & _
    ",rechdr.Date " & _
    ",iitem.InvoiceItemName as [Invoice Item] " & _
    ",emp.Username as [User] " & _
    ",csr.Username as [CSR] " & _
    ",recdtl.Amount " & _
    ",ofc.Office " & _
    ",rechdr.IdOffice, rechdr.AmountPaid, rechdr.IdEmployeeCSR2, rechdr.BalanceDue, rechdr.IdCustomer, csr2.username as [csr2]  " & _
    "FROM  [ReceiptsDTL] recdtl " & _
    "inner join  ReceiptsHDR  rechdr on recdtl.IdReceiptHDR=rechdr.IDReceiptHDR " & _
    "inner join PoliciesHDR polhdr on polhdr.IdPoliciesHDR=rechdr.IdPoliciesHDR " & _
    "left join EmployeeInfo emp on emp.IDEmployee=rechdr.IdEmployeeUSR " & _
    "left join EmployeeInfo csr on csr.IDEmployee=rechdr.IdEmployeeCSR1 " & _
    "left join EmployeeInfo csr2 on csr2.IDEmployee=rechdr.IdEmployeeCSR2 " & _
    "inner join OfficesCatalog ofc on ofc.IdOffice=rechdr.IdOffice " & _
    "inner join OfficesCatalog ofcpol on ofcpol.IdOffice=rechdr.IdOffice " & _
    "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
    "inner join Customers cus on cus.IdCustomer=polhdr.IdCustomer " & _
    "inner join InvoiceItemCatalog iitem on iitem.IdInvoiceItem=recdtl.IdInvoiceItem " & _
    "inner join BankAccountCatalog bank on bank.IdBankAccount=iitem.IdBankAccount " & _
    "inner join CompanysAGICatalog ja on ja.IdCompanyAGI=polhdr.IdCompanyAGI " & _
    "where iitem.IdInvoiceItem in (2,35,3,11,15,22,23,33,9,30,33,17,13,31,32, 36, 37,39,42,43) " & _
    "and cast(rechdr.Date as Date) >= '" + txtdatefrom.Text + "' AND cast( rechdr.DATE as Date) <= '" + txtdateto.Text + "' " & _
    "and rechdr.Active=1 " + r$ + "order by [Receipt #], rechdr.IdCustomer"
    
    GoTo aaa
    
    
 sSelect = "SELECT rechdr.[IdReceiptHDR] as [Receipt #] ,rechdr.Date ,iitem.InvoiceItemName as [Invoice Item] ,emp.Username as [User] ,csr.Username as [CSR] ,recdtl.amount ,ofc.Office , " & _
"rechdr.IdOffice, rechdr.AmountPaid, rechdr.IdEmployeeCSR2, rechdr.BalanceDue, rechdr.IdCustomer, csr2.Username as [CSR2]  FROM  [EmplGoalsCalc] goals " & _
"inner join ReceiptsHDR rechdr on rechdr.IDReceiptHDR= goals.IdReceiptHDR " & _
 "inner join  ReceiptsDTL  recdtl on recdtl.IdReceiptHDR=rechdr.IDReceiptHDR " & _
"inner join InvoiceItemCatalog iitem on iitem.IdInvoiceItem=goals.IdInvoiceItem " & _
"inner join EmployeeInfo emp on emp.IDEmployee=rechdr.IdEmployeeUSR " & _
"left join EmployeeInfo csr on csr.IDEmployee=rechdr.IdEmployeeCSR1 " & _
"left join EmployeeInfo csr2 on csr2.IDEmployee=rechdr.IdEmployeeCSR2 " & _
"inner join OfficesCatalog ofc on ofc.IdOffice=rechdr.IdOffice " & _
"where iitem.IdInvoiceItem in (2,35,3,11,15,22,23,33,9,30,33,17,13,31,32, 36, 37,39, 42,43) " & _
"and cast(rechdr.Date as Date) >= '" + txtdatefrom.Text + "' AND cast( rechdr.DATE as Date) <= '" + txtdateto.Text + "' " & _
"and rechdr.Active=1 and goals.Active=1 " + r$ & _
"order by emp.IdEmployee,goals.IdInvoiceItem, goals.IdReceiptHDR"

aaa:
    
    ' ---------------------------------------------------------------------------
    
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
   Rs.Open sSelect, base, adOpenStatic, adLockOptimistic
    
    
   Rs.MoveLast

   Rs.MoveFirst
   ' Assuming that rs is your ADO recordset
   Grid2.Rows = Rs.RecordCount + 1

   rsVar = Rs.GetString(adClipString, Rs.RecordCount)

   Grid2.cols = Rs.Fields.Count + 1
    
    
    
   Grid2.TextMatrix(0, 0) = ""
   ' Set column names in the grid
   For i = 0 To Rs.Fields.Count - 1
      Grid2.TextMatrix(0, i + 1) = Rs.Fields(i).name
   Next

   Grid2.row = 1
   Grid2.col = 1

   ' Set range of cells in the grid
   Grid2.RowSel = Grid2.Rows - 1
   Grid2.ColSel = Grid2.cols - 1
   Grid2.clip = rsVar

   ' Reset the grid's selected range of cells
   Grid2.RowSel = Grid2.row
   Grid2.ColSel = Grid2.col

   Rs.Close

   Set Rs = Nothing


   If agente$ = "JENNIFERR" Then
      For t = 1 To Grid2.Rows - 1
          Grid2.row = t
          Grid2.col = 4
          userk$ = UCase(Grid2.Text)
          
          Grid2.col = 5
          csrk$ = UCase(Grid2.Text)
          
          If userk$ <> "JENNIFERR" And csrk$ <> "JENNIFERR" Then
             Grid2.col = 5
             Grid2.Text = "JENNIFERR"
          
          End If
         
      Next t
  
   End If


' ----------------------------------------------------------------------------

GoTo salida
    
    
    

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    
    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
    ' user1$ = Rs(0)
          
                        
                         
    Rs.Close
    
    
salida:
    
    
    For t = 1 To Grid2.Rows - 1
       Grid2.row = t
       Grid2.col = 0
       Grid2.Text = t
    Next t
    
    
 setup_grid2

  Grid2.Visible = True

If agente$ = "" Then
  CARGA_AGENTES
End If


'carga_attributos
End Sub

Public Sub carga_inv()
On Error Resume Next

' ************************************************************
  
 Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    
    If agente$ = "" Then Exit Sub
    
    
    Set Rs = New ADODB.Recordset
    'Checa_status
   
    user1$ = ""
    'sSelect = "SELECT login From employees where login='hnavarro'"
    
    
    fecha1$ = txtdatefrom.Text
    
    
    diadelano = Format(fecha1$, "y")
    dia30 = diadelano - 35
    fecha2$ = Format(dia30, "mm/dd")
    
    If dia30 <= 0 Then
       ano2$ = Format(Val((Format(Now, "yyyy"))) - 1, "####")
    Else
       ano2$ = (Format(Now, "yyyy"))
    End If
    
    fecha30$ = fecha2$ + "/" + ano2$
    
    
    If agente$ <> "" Then
    
      r$ = " and (emp.username='" + agente$ + "' or csr.username='" + agente$ + "') "
      
   
    Else
    
      r$ = " "
    
    End If
    
    
       
    
    
     Grid3.Visible = False
    
    sSelect = "SELECT " & _
    "recdtl.[IdReceiptHDR] as [Receipt #] " & _
    ",rechdr.Date ,cus.Idcustomer " & _
    ",iitem.InvoiceItemName as [For] " & _
    ",emp.Username as [User] " & _
    ",csr.Username as [CSR] " & _
    ",rechdr.BalanceDue as [Balance Due] " & _
    ",ofc.Office " & _
    ",rechdr.IdOffice " & _
    "FROM  [ReceiptsDTL] recdtl " & _
    "inner join  ReceiptsHDR  rechdr on recdtl.IdReceiptHDR=rechdr.IDReceiptHDR " & _
    "inner join PoliciesHDR polhdr on polhdr.IdPoliciesHDR=rechdr.IdPoliciesHDR " & _
    "inner join EmployeeInfo emp on emp.IDEmployee=rechdr.IdEmployeeUSR " & _
    "inner join EmployeeInfo csr on csr.IDEmployee=rechdr.IdEmployeeCSR1 " & _
    "inner join OfficesCatalog ofc on ofc.IdOffice=rechdr.IdOffice " & _
    "inner join OfficesCatalog ofcpol on ofcpol.IdOffice=rechdr.IdOffice " & _
    "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
    "inner join Customers cus on cus.IdCustomer=polhdr.IdCustomer " & _
    "inner join InvoiceItemCatalog iitem on iitem.IdInvoiceItem=recdtl.IdInvoiceItem " & _
    "inner join BankAccountCatalog bank on bank.IdBankAccount=iitem.IdBankAccount " & _
    "inner join CompanysAGICatalog ja on ja.IdCompanyAGI=polhdr.IdCompanyAGI " & _
    "where iitem.IdInvoiceItem=19 and " & _
    "cast(rechdr.Date as Date) >= '" + fecha30$ + "' " & _
    "AND cast( rechdr.DATE as Date) <= '" + txtdateto.Text + "' " + r$ & _
    "and rechdr.Active=1 order by [Receipt #]"
    
    
   

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    Grid3.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
     Set Grid3.DataSource = Rs
     
     'user1$ = Rs(0)
          
                        
                         
    Rs.Close
    
    
    
    
   

    
    
    
    
    
    
    
    grid.Rows = Grid3.Rows ' - 1
    grid.cols = Grid3.cols
    
    grid.Clear
    
    For t = 0 To Grid3.Rows - 1
      grid.row = t
      Grid3.row = t
      
      For Y = 0 To Grid3.cols - 1
         grid.col = Y
         Grid3.col = Y
         grid.Text = Grid3.Text
      Next Y
    Next t
    
      
   
    
    
    sSelect = "SELECT " & _
    "recdtl.[IdReceiptHDR] as [Receipt #] " & _
    ",rechdr.Date ,cus.Idcustomer " & _
    ",iitem.InvoiceItemName as [For] " & _
    ",emp.Username as [User] " & _
    ",csr.Username as [CSR] " & _
    ",rechdr.BalanceDue as [Balance Due] " & _
    ",ofc.Office " & _
    ",rechdr.IdOffice " & _
    "FROM  [ReceiptsDTL] recdtl " & _
    "inner join  ReceiptsHDR  rechdr on recdtl.IdReceiptHDR=rechdr.IDReceiptHDR " & _
    "inner join PoliciesHDR polhdr on polhdr.IdPoliciesHDR=rechdr.IdPoliciesHDR " & _
    "inner join EmployeeInfo emp on emp.IDEmployee=rechdr.IdEmployeeUSR " & _
    "inner join EmployeeInfo csr on csr.IDEmployee=rechdr.IdEmployeeCSR1 " & _
    "inner join OfficesCatalog ofc on ofc.IdOffice=rechdr.IdOffice " & _
    "inner join OfficesCatalog ofcpol on ofcpol.IdOffice=rechdr.IdOffice " & _
    "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
    "inner join Customers cus on cus.IdCustomer=polhdr.IdCustomer " & _
    "inner join InvoiceItemCatalog iitem on iitem.IdInvoiceItem=recdtl.IdInvoiceItem " & _
    "inner join BankAccountCatalog bank on bank.IdBankAccount=iitem.IdBankAccount " & _
    "inner join CompanysAGICatalog ja on ja.IdCompanyAGI=polhdr.IdCompanyAGI " & _
    "where iitem.IdInvoiceItem in (35) and " & _
    "cast(rechdr.Date as Date) >= '" + txtdatefrom.Text + "' " & _
    "AND cast( rechdr.DATE as Date) <= '" + txtdateto.Text + "' " + r$ & _
    "and rechdr.Active=1 order by [Receipt #]"
    
    
    
        
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    ' Asignar el recordset al FlexGrid
    Set Grid3.DataSource = Rs
                        
                        
    Rs.Close
    
    

    
    
    
    fila = grid.Rows
    
    total_filas = grid.Rows + (Grid3.Rows - 1)
    grid.Rows = total_filas
    
       
    For t = 1 To Grid3.Rows - 1
      
      grid.row = fila
      Grid3.row = t
      
      For Y = 0 To Grid3.cols - 1
         grid.col = Y
         Grid3.col = Y
         grid.Text = Grid3.Text
      Next Y
      
      fila = fila + 1
    Next t
    
    
    
    
       
    
    For t = 1 To fila
       grid.row = t
       grid.col = 0
       grid.Text = t
    Next t
    
    
    Grid3.Clear
    Grid3.Rows = grid.Rows
    
    For t = 0 To grid.Rows - 1
      grid.row = t
      Grid3.row = t
      
      For Y = 0 To grid.cols - 1
         grid.col = Y
         Grid3.col = Y
         Grid3.Text = grid.Text
      Next Y
    Next t
    
    
setup_grid3



separa_campos_inv

 Grid3.Visible = True

  'CARGA_AGENTES

'carga_attributos
End Sub

Public Sub carga_tiers()
On Error Resume Next

Dim sSelect As String
Dim Rs As ADODB.Recordset
    

Exit Sub

Set Rs = New ADODB.Recordset
    sSelect = "SELECT NBgoalmin From tierscatalog where tier='1' and idjobtitle='16'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    NB_Goal1 = Rs(0)
                        
    Rs.Close
    
    
    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT NBgoalmin From tierscatalog where tier='2'  and idjobtitle='16'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    NB_Goal2 = Rs(0)
                        
    Rs.Close
    
    
    
    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT NBgoalmin From tierscatalog where tier='3' and idjobtitle='16'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    NB_Goal3 = Rs(0)
                        
    Rs.Close
    


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT NBgoalmin From tierscatalog where tier='4' and idjobtitle='16'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    NB_Goal4 = Rs(0)
                        
    Rs.Close





    Set Rs = New ADODB.Recordset
    sSelect = "SELECT NBgoalmin From tierscatalog where tier='5' and idjobtitle='16'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    NB_Goal5 = Rs(0)
                        
    Rs.Close







 Set Rs = New ADODB.Recordset
    sSelect = "SELECT flatamount From tierscatalog where tier='1' and idjobtitle='16'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    commission1 = Rs(0)
                        
    Rs.Close
    
    
    
    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT percComm From tierscatalog where tier='2' and idjobtitle='16'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    commission2 = Rs(0)
                        
    Rs.Close
    
    
    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT percComm From tierscatalog where tier='3' and idjobtitle='16'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    commission3 = Rs(0)
                        
    Rs.Close
    
    
    
    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT percComm From tierscatalog where tier='4' and idjobtitle='16'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    commission4 = Rs(0)
                        
    Rs.Close
    
    
    
    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT percComm From tierscatalog where tier='5' and idjobtitle='16'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    commission5 = Rs(0)
                        
    Rs.Close
    
    
    
    
    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmin From tierscatalog where tier='1' and idjobtitle='16'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango1a = Rs(0)
                        
    Rs.Close


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmax From tierscatalog where tier='1' and idjobtitle='16'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango1b = Rs(0)
                        
    Rs.Close


    If rango1b > 999999 Then rango1b = 9999
    
    
    
 Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmin From tierscatalog where tier='2' and idjobtitle='16'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango2a = Rs(0)
                        
    Rs.Close


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmax From tierscatalog where tier='2' and idjobtitle='16'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango2b = Rs(0)
                        
    Rs.Close
    
    If rango2b > 999999 Then rango2b = 9999
    

Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmin From tierscatalog where tier='3' and idjobtitle='16'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango3a = Rs(0)
                        
    Rs.Close


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmax From tierscatalog where tier='3' and idjobtitle='16'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango3b = Rs(0)
                        
    Rs.Close
    
    
    If rango3b > 999999 Then rango3b = 9999
    
    
Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmin From tierscatalog where tier='4' and idjobtitle='16'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango4a = Rs(0)
                        
    Rs.Close


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmax From tierscatalog where tier='4' and idjobtitle='16'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango4b = Rs(0)
                        
    Rs.Close
    
    If rango4b > 999999 Then rango4b = 9999
    
' -------------------------

Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmin From tierscatalog where tier='5' and idjobtitle='16'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango5a = Rs(0)
                        
    Rs.Close


    Set Rs = New ADODB.Recordset
    sSelect = "SELECT bfgoalmax From tierscatalog where tier='5' and idjobtitle='16'"
     
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
       
    ' Asignar el recordset al FlexGrid
    rango5b = Rs(0)
                        
    Rs.Close


If rango5b > 999999 Then rango5b = 9999
  

End Sub

Public Sub carga_tier_del_agente()
On Error Resume Next

Dim sSelect As String
Dim Rs As ADODB.Recordset
    
Set Rs = New ADODB.Recordset



 sSelect = "Select idemppayrolllink, idemployeeLAE, idpaytype, initials, manager, MTY, Commercial, PhoneSales, Exception1,Typeofexception1," & _
    "Exception2,Typeofexception2,Exception3,Typeofexception3,Exception4,Typeofexception4,Exception5,Typeofexception5,deb_mod, nb_deb_mod, porc_mod, autorizado, " & _
    "office1, office2, emp.firstname, emp.lastname1, emp.username from [payrollconfig] payroll " & _
    "inner join EmployeeInfo emp on emp.IDEmployee=payroll.IdEmployeeLAE "
    ' "where emp.username='" + txtagent.Text + "'"


  ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
        
     ' Permitir redimensionar las columnas
    grid.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid.DataSource = Rs
                         
    Rs.Close



' verifica si existe el agente
existe = 0
For t = 1 To grid.Rows - 1
 
 grid.row = t
 grid.col = 27
 username1$ = grid.Text
 
 
 If RTrim(UCase(username1$)) = RTrim(UCase(agente$)) Then
    existe = 1
     
    grid.col = 3
    categoria_tier = Val(grid.Text)
    
    Exit For
    
 End If
Next t


If existe = 0 Then
   categoria_tier = 1
End If




End Sub

Public Sub setup_grid1()
On Error Resume Next

Grid1.ColWidth(0) = 600
Grid1.ColWidth(1) = 1100  ' recibo#
Grid1.ColWidth(2) = 2300 ' fecha
Grid1.ColWidth(3) = 3500 ' factura
Grid1.ColWidth(4) = 1400 ' user
Grid1.ColWidth(5) = 1400  ' csr
Grid1.ColWidth(6) = 900   ' cantidad
Grid1.ColWidth(7) = 1700  ' oficina
Grid1.ColWidth(8) = 800  ' IDoficina


End Sub

Public Sub setup_grid2()
On Error Resume Next

Grid2.ColWidth(0) = 600
Grid2.ColWidth(1) = 1100  ' recibo#
Grid2.ColWidth(2) = 2300 ' fecha
Grid2.ColWidth(3) = 3500 ' factura
Grid2.ColWidth(4) = 1400 ' user
Grid2.ColWidth(5) = 1400  ' csr
Grid2.ColWidth(6) = 900   ' cantidad
Grid2.ColWidth(7) = 1700  ' oficina
Grid2.ColWidth(8) = 800  ' IDoficina

End Sub

Public Sub setup_grid3()
On Error Resume Next



Grid3.ColWidth(0) = 600
Grid3.ColWidth(1) = 1100  ' recibo#
Grid3.ColWidth(2) = 2300 ' fecha
Grid3.ColWidth(3) = 1200 ' idcustomer
Grid3.ColWidth(4) = 1900 ' for
Grid3.ColWidth(5) = 1400  ' user
Grid3.ColWidth(6) = 1400   ' csr
Grid3.ColWidth(7) = 1200  ' cantidad
Grid3.ColWidth(8) = 1800  ' Oficina
Grid3.ColWidth(9) = 800 ' idoficina




End Sub

Public Function agentes_de_invoice(recibo As String)
On Error Resume Next

Dim sSelect As String
Dim Rs As ADODB.Recordset
    
Set Rs = New ADODB.Recordset

csr2$ = ""
creador_factura$ = ""
creador_late_fee$ = ""
unico = 0

 sSelect = "select idemployeeUSR, idemployeeCSR1, idemployeeCSR2 from ReceiptsBalancePayments recbalpay " & _
         "inner join ReceiptsHDR rechdr on recbalpay.IdReceiptsHDRWBalance=rechdr.IDReceiptHDR " & _
         "Where IdReceiptsHDRPayBalance='" + recibo + "'"
         
          ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
          Rs.Open sSelect, base, adOpenUnspecified
    
          user1$ = Rs(0)
          csr1$ = Rs(1)
          csr2$ = Rs(2)
          Rs.Close
          
          
         
          
          
          If user1$ = "" Or csr1$ = "" Then
            
             Exit Function
          End If
          
          
          If csr2$ = "" Then
             creador_factura$ = user1$
             CSR_factura$ = csr1$
          Else
             creador_factura$ = csr1$
             CSR_factura$ = csr2$
             
          End If
          
          
           
          
          
          ' ///////////////////////////
          
           sSelect = "select idemployeeUSR, idemployeeCSR1, idemployeeCSR2 from [ReceiptsDTL] recdtl " & _
          "inner join  ReceiptsHDR  rechdr on rechdr.IDReceiptHDR =recdtl.IdReceiptHDR " & _
          "inner join InvoiceItemCatalog iitem on iitem.IdInvoiceItem=recdtl.IdInvoiceItem " & _
          "where rechdr.Active=1 and recdtl.[IdReceiptHDR] in (" + recibo + ") and iitem.InvoiceItemName='Invoice'"
          
          Rs.Open sSelect, base, adOpenUnspecified
          id_user_originalx$ = Rs(0)
          csrx$ = Rs(1)
          csr2x$ = Rs(2)
                    
          Rs.Close
          
          
          If csrx$ = csr2x$ Then
             unico = 1
          End If
          
           ' asigna usuario primero
       '   If csr2x$ <> "" Then
       '      creador_factura$ = csr2x$
       '      CSR_factura$ = csrx$
       '   Else
       '      creador_factura$ = csrx$
       '      CSR_factura$ = id_user_originalx$
       '   End If
          
          ' /////////////////////////////
          
          
          sSelect = "select idemployeeUSR, idemployeeCSR1, idemployeeCSR2 from [ReceiptsDTL] recdtl " & _
          "inner join  ReceiptsHDR  rechdr on rechdr.IDReceiptHDR =recdtl.IdReceiptHDR " & _
          "inner join InvoiceItemCatalog iitem on iitem.IdInvoiceItem=recdtl.IdInvoiceItem " & _
          "where rechdr.Active=1 and recdtl.[IdReceiptHDR] in (" + recibo + ") and iitem.InvoiceItemName='Late Fee'"
          
          Rs.Open sSelect, base, adOpenUnspecified
          id_user_original$ = Rs(0)
          csr$ = Rs(1)
          csr2$ = Rs(2)
                    
          Rs.Close
          
          
          
        '
          
          
          If UCase(concepto$) = "LATE FEE" And csr2$ <> "" Then
             creador_late_fee$ = csr2$
         
          End If
          
          
              
          
          
          
          sSelect = "select username from EmployeeInfo where IDEmployee='" + creador_factura$ + "'"
          Rs.Open sSelect, base, adOpenUnspecified
          user_original$ = UCase(Rs(0))
          Rs.Close
           
           
          sSelect = "select username from EmployeeInfo where IDEmployee='" + CSR_factura$ + "'"      ' aqui estaba  id_user_originalx$
          Rs.Open sSelect, base, adOpenUnspecified
          csr_original$ = UCase(Rs(0))
          Rs.Close
          
          
          
          If UCase(concepto$) = "LATE FEE" Then
            sSelect = "select username from EmployeeInfo where IDEmployee='" + creador_late_fee$ + "'"
            Rs.Open sSelect, base, adOpenUnspecified
            user_original$ = UCase(Rs(0))
            Rs.Close
          End If
          
          
          
continua_aqui:
          
          
          If UCase(agente$) = "JENNIFERR" And (user_original$ <> agente$ And csr_original$ <> agente$) Then
             
             csr_original$ = agente$
                      
             Exit Function
             
          End If
          
          
          
          
          
          
          If unico = 1 Then
              
              csr_original$ = user_original$
             
          End If
          
          
          
End Function

Public Sub carga_factura()
On Error Resume Next

' ************************************************************
  
 Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    Dim rsVar As Variant
   Dim i As Integer
   
    
    Set Rs = New ADODB.Recordset
    'Checa_status
   
    user1$ = ""
   
    
    
    sSelect = "SELECT " & _
    "recdtl.[IdReceiptHDR] as [Receipt #] " & _
    ",rechdr.Date " & _
    ",iitem.InvoiceItemName as [Invoice Item] " & _
    ",emp.Username as [User] " & _
    ",csr.Username as [CSR] " & _
    ",recdtl.Amount " & _
    ",ofc.Office " & _
    ",rechdr.IdOffice, rechdr.AmountPaid, rechdr.IdEmployeeCSR2, rechdr.BalanceDue, rechdr.IdCustomer, polhdr.PolicyNumber  " & _
    "FROM  [ReceiptsDTL] recdtl " & _
    "inner join  ReceiptsHDR  rechdr on recdtl.IdReceiptHDR=rechdr.IDReceiptHDR " & _
    "inner join PoliciesHDR polhdr on polhdr.IdPoliciesHDR=rechdr.IdPoliciesHDR " & _
    "inner join EmployeeInfo emp on emp.IDEmployee=rechdr.IdEmployeeUSR " & _
    "inner join EmployeeInfo csr on csr.IDEmployee=rechdr.IdEmployeeCSR1 " & _
    "inner join OfficesCatalog ofc on ofc.IdOffice=rechdr.IdOffice " & _
    "inner join OfficesCatalog ofcpol on ofcpol.IdOffice=rechdr.IdOffice " & _
    "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
    "inner join Customers cus on cus.IdCustomer=polhdr.IdCustomer " & _
    "inner join InvoiceItemCatalog iitem on iitem.IdInvoiceItem=recdtl.IdInvoiceItem " & _
    "inner join BankAccountCatalog bank on bank.IdBankAccount=iitem.IdBankAccount " & _
    "inner join CompanysAGICatalog ja on ja.IdCompanyAGI=polhdr.IdCompanyAGI " & _
    "where iitem.IdInvoiceItem in (2,35,3,11,15,22,23,33,9,30,33,17,13,31,32, 36, 37,39,42,43) " & _
    "and cast(rechdr.Date as Date) >= '" + txtdatefrom.Text + "' AND cast( rechdr.DATE as Date) <= '" + txtdateto.Text + "' " & _
    "and rechdr.IdCustomer='" + ID_Cliente$ + "' and (iitem.InvoiceItemName='Invoice' or iitem.InvoiceItemName='Late Fee') " & _
    "and rechdr.Active=1 order by [Receipt #]"
    
    
    
    ' ---------------------------------------------------------------------------
    
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
   Rs.Open sSelect, base, adOpenStatic, adLockOptimistic
    
    
   Rs.MoveLast

   Rs.MoveFirst
   ' Assuming that rs is your ADO recordset
   grid8.Rows = Rs.RecordCount + 1

   rsVar = Rs.GetString(adClipString, Rs.RecordCount)

   grid8.cols = Rs.Fields.Count + 1
    
    
    
   grid8.TextMatrix(0, 0) = ""
   ' Set column names in the grid
   For i = 0 To Rs.Fields.Count - 1
      grid8.TextMatrix(0, i + 1) = Rs.Fields(i).name
   Next

   grid8.row = 1
   grid8.col = 1

   ' Set range of cells in the grid
   grid8.RowSel = grid8.Rows - 1
   grid8.ColSel = grid8.cols - 1
   grid8.clip = rsVar

   ' Reset the grid's selected range of cells
   grid8.RowSel = grid8.row
   grid8.ColSel = grid8.col

   Rs.Close

   Set Rs = Nothing


' ----------------------------------------------------------------------------

   
    
salida:
    
    
    For t = 1 To grid8.Rows - 1
       grid8.row = t
       grid8.col = 0
       grid8.Text = t
    Next t
    
    

End Sub

Public Sub calcula_NB()
On Error Resume Next

 Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
      Dim rsVar As Variant
   Dim i As Integer
    
    Set Rs = New ADODB.Recordset
    'Checa_status
   
    user1$ = ""
    
    If agente$ <> "" Then
    
      r$ = "' and (emp.username='" + agente$ + "' or csr.username='" + agente$ + "') "
    Else
    
      r$ = "' "
    
    End If
    
    Grid1.Visible = False
    
    sSelect = "SELECT " & _
    "recdtl.[IdReceiptHDR] as [Receipt #] " & _
    ",rechdr.Date " & _
    ",iitem.InvoiceItemDesc as [Invoice Item] " & _
    ",emp.Username as [User] " & _
    ",csr.Username as [CSR] " & _
    ",recdtl.Amount " & _
    ",ofc.Office " & _
    ",rechdr.IdOffice " & _
    "FROM  [ReceiptsDTL] recdtl " & _
    "inner join  ReceiptsHDR  rechdr on recdtl.IdReceiptHDR=rechdr.IDReceiptHDR " & _
    "inner join PoliciesHDR polhdr on polhdr.IdPoliciesHDR=rechdr.IdPoliciesHDR " & _
    "inner join EmployeeInfo emp on emp.IDEmployee=rechdr.IdEmployeeUSR " & _
    "inner join EmployeeInfo csr on csr.IDEmployee=rechdr.IdEmployeeCSR1 " & _
    "inner join OfficesCatalog ofc on ofc.IdOffice=rechdr.IdOffice " & _
    "inner join OfficesCatalog ofcpol on ofcpol.IdOffice=rechdr.IdOffice " & _
    "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
    "inner join Customers cus on cus.IdCustomer=polhdr.IdCustomer " & _
    "inner join InvoiceItemCatalog iitem on iitem.IdInvoiceItem=recdtl.IdInvoiceItem " & _
    "inner join BankAccountCatalog bank on bank.IdBankAccount=iitem.IdBankAccount " & _
    "inner join CompanysAGICatalog ja on ja.IdCompanyAGI=polhdr.IdCompanyAGI " & _
    "where iitem.IdInvoiceItem in (1,20) " & _
    "and cast(rechdr.Date as Date) >= '" + txtdatefrom.Text + "' " & _
    "AND cast( rechdr.DATE as Date) <= '" + txtdateto.Text + r$ & _
    "and rechdr.Active=1 order by emp.Username, csr.Username"
    
    
    
    
    ' ---------------------------------------------------------------------------
    
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
   Rs.Open sSelect, base, adOpenStatic, adLockOptimistic
    
    
   Rs.MoveLast

   Rs.MoveFirst
   ' Assuming that rs is your ADO recordset
   Grid1.Rows = Rs.RecordCount + 1

   rsVar = Rs.GetString(adClipString, Rs.RecordCount)

   Grid1.cols = Rs.Fields.Count + 1
    
    
    
   Grid1.TextMatrix(0, 0) = ""
   ' Set column names in the grid
   For i = 0 To Rs.Fields.Count - 1
      Grid1.TextMatrix(0, i + 1) = Rs.Fields(i).name
   Next

   Grid1.row = 1
   Grid1.col = 1

   ' Set range of cells in the grid
   Grid1.RowSel = Grid1.Rows - 1
   Grid1.ColSel = Grid1.cols - 1
   Grid1.clip = rsVar

   ' Reset the grid's selected range of cells
   Grid1.RowSel = Grid1.row
   Grid1.ColSel = Grid1.col

   Rs.Close

   Set Rs = Nothing


' ----------------------------------------------------------------------------


    
    
    For t = 1 To Grid1.Rows - 1
       Grid1.row = t
       Grid1.col = 0
       Grid1.Text = t
    Next t
    
    
    
  setup_grid1

  Grid1.Visible = True


End Sub

Public Sub carga_tabla()
On Error Resume Next

txtresultado.Text = ""


For t = 1 To contador_usuarios
  r$ = ""
  r$ = Format(matrix_NB(t, 0), "!@@@@@@@@@@@@@@@") + " " + Format(matrix_NB(t, 1), "!@@@@@@@@@@@@@@@") + " "
  For Y = 2 To 4
     r$ = r$ + Format(Val(matrix_NB(t, Y)), "00.0") + " "
  Next Y
    
  For Y = 5 To 12
     r$ = r$ + Format(matrix_NB(t, Y), "0") + " "
  Next Y
    
    
  r$ = r$ + Chr$(13)
  
  
  If letra$ = "-" Then
     txtresultado.Text = txtresultado.Text + r$
  Else
     If matrix_NB(t, 0) = agente$ Or matrix_NB(t, 1) = agente$ Then
         txtresultado.Text = txtresultado.Text + r$
     End If
  End If
  
  
Next t




End Sub

Public Sub Calcula_cantidades_old()
On Error Resume Next

   
    

Dim gtotal As Double
List1.Clear
List2.Clear
List3.Clear
List9.Clear
List10.Clear

total_facturas_propias = 0
total_facturas_ajenas = 0

'List4.Clear
'List5.Clear

Dim Commercial$(50, 3)
Dim sSelect As String
Dim Rs As ADODB.Recordset
    
Set Rs = New ADODB.Recordset

contador_comercial = 0

Alta_Comercial = 0
par_encontrado = 0
For t = 1 To Grid2.Rows - 1
   
   
   NO_CAMBIAR = False
   csr$ = ""
   Grid2.col = 5
   Grid2.row = t
   csr$ = UCase(Grid2.Text)
  
   Grid2.col = 4
   user$ = UCase(Grid2.Text)
   
   
   invoice_3_users = 0
   
   ' == S T A R T ======================================================================
   
   csr2$ = ""
   Grid2.col = 10
   csr2$ = UCase(Grid2.Text)
   
   If csr2$ <> "" Then
    sSelect = "select username from EmployeeInfo where IDEmployee='" + csr2$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    csr2$ = UCase(Rs(0))
    
    Rs.Close
   End If
   
   
   
   
   
   Grid2.col = 1
   recibo$ = Grid2.Text
   
   
   
   
   ' verifica si existe un agente tercero en un invoice
   
   Grid2.col = 3
   concepto$ = Grid2.Text
   
   Grid2.col = 6
   cantidad$ = Grid2.Text
   
  
   Grid2.col = 12
   ID_Cliente$ = Grid2.Text
   
   
   
   
   If recibo$ = "254678" Then
     '  Stop
   End If
   
   
   If (UCase$(LTrim(concepto$)) = "INVOICE") Then
   '  Stop
   End If
   
   
   
     ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
     
    If (UCase$(LTrim(concepto$)) = "INVOICE") Or (UCase$(LTrim(concepto$)) = "LATE FEE") Then
    
     user_original$ = user$
     csr_original$ = csr$
   
     X = agentes_de_invoice(recibo$)
   
     'userx$ = user_original$
     'csr$ = csr_original$   ' este estaba anulado
     
     If agente$ = UCase(csr_original$) Or agente$ = UCase(user_original$) Then
        csr$ = UCase(csr_original$)
        user$ = UCase(user_original$)
        'GoTo hayado_en_la_factura
     End If
     
    End If
        
   
   ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
   
   
   
   
   If agente$ = UCase(csr$) Or agente$ = UCase(user$) Then
   Else
        GoTo No_encontrado
   End If
   
   
hayado_en_la_factura:
   
        ' verifica si CSR2 es manager comercial
        
        
    If agente$ = csr$ Or agente$ = user$ Then
    
             existe = 0
             manager_csr2$ = ""
             Agent_Oficinacommercial$ = ""
     
             For z = 0 To lista_managers.ListCount - 1
                    a$ = RTrim(UCase(Left$(lista_managers.List(z), 20)))
                    b$ = LTrim(RTrim(Right(UCase(lista_managers.List(z)), Len(lista_managers.List(z)) - 20)))
        
                    If a$ = csr2$ Then '
                           If b$ = "MANAGER_COMMERCIAL" Then
                                    manager_csr2$ = a$
                                    existe = 1
                           End If
          
          
                           If b$ = "AGENT_COMMERCIAL" Then
                                    Agent_Oficinacommercial$ = a$
                                    existe = 1
                           End If
          
          
                           If b$ = "COMMERCIAL" Then
                                    manager_csr2$ = a$
                                    existe = 1
                           End If
                    
                    End If
        
        
                    If existe = 1 Then
                            Exit For
                    End If
        
        
             Next z
          
             If existe = 1 Then
                    contador_comercial = contador_comercial + 1
       
                    If Agent_Oficinacommercial$ <> "" Then
                           Commercial$(contador_comercial, 0) = user$
                           Commercial$(contador_comercial, 1) = cantidad$
                           Commercial$(contador_comercial, 2) = Agent_Oficinacommercial$
                    Else
                           Commercial$(contador_comercial, 0) = user$
                           Commercial$(contador_comercial, 1) = cantidad$
                           Commercial$(contador_comercial, 2) = manager_csr2$
                    End If
             End If
  
    End If
   
   
   
   
   
   'GoTo aaa
   
   If (UCase$(LTrim(concepto$)) = "INVOICE") And csr2$ <> "" And (agente$ = user$ Or agente$ = csr$ Or agente$ = csr2$) Then
       
        conta1 = 0
        
        If user$ = csr2$ Then
           conta1 = conta1 + 1
        End If
        
        If csr$ = csr2$ Then
           conta1 = conta1 + 1
        End If
        
        
        If agente$ = csr2$ And (user$ <> agente$ And csr$ <> agente$) Then
           conta1 = conta1 + 1
        End If
        
      
        If agente$ <> csr2$ And (user$ = agente$ And csr$ = agente$) And csr2$ <> "" Then
           conta1 = conta1 + 1
        End If
        
        
        If agente$ <> csr2$ And agente$ <> user$ And user$ <> csr2$ And csr2$ <> "" Then
           conta1 = conta1 + 1
        
        End If
        
        
        If agente$ <> csr2$ And agente$ <> csr$ And user$ = agente$ And csr2$ <> "" Then
           conta1 = conta1 + 1
        
        End If
        
      
      
      '  >>>>>>>>>>>>>>>>>>>>>>>>>>>>   VERIFICA AQUI SI ES MANAGER
     
     ' verifica si CSR o USER es manager-COMERCIAL
     existe = 0
     CSR2_Es_manager_commercial = 0
    
     For z = 0 To lista_managers.ListCount - 1
       a$ = RTrim(UCase(Left$(lista_managers.List(z), 20)))
       b$ = LTrim(RTrim(Right(UCase(lista_managers.List(z)), Len(lista_managers.List(z)) - 20)))
        
        
        If a$ = agente$ Or a$ = csr$ Or a$ = user$ Then
          If b$ = "COMMERCIAL" Or b$ = "MANAGER_COMMERCIAL" And Alta_Comercial = 0 Then
             manager_commercial$ = a$
             Alta_Comercial = 1
             existe = 1
             'Exit For
          End If
             
        End If
        
            
                
        If a$ = agente$ Or a$ = csr$ Or a$ = user$ Then
          If b$ = "AGENT_COMMERCIAL" Then
             Agent_Oficinacommercial$ = a$
             existe = 1
             'Exit For
          End If
             
        End If
        
        
        
        
         If a$ = csr2$ Then
          If b$ = "COMMERCIAL" Or b$ = "MANAGER_COMMERCIAL" Or b$ = "AGENT_COMMERCIAL" Then
             CSR2_Es_manager_commercial = 1
             'Exit For
          End If
             
        End If
        
            
        
      Next z
     
     
     '  >>>>>>>>>>>>>>>>>>>>>>>>>>>>   TERMINA AQUI SI ES MANAGER
     
      
        
      
      
       
        If conta1 >= 1 Then
        
                
            sSelect = "select IdReceiptsHDRWBalance from ReceiptsBalancePayments where IdReceiptsHDRPayBalance='" + recibo$ + "'"
            Rs.Open sSelect, base, adOpenUnspecified
            id_reciboHDRwBAL$ = Rs(0)
            Rs.Close
         
            ' GoTo aaa
            
            
           
      
            
            
            
            
        '    user1$ = ""
        '    CSR1$ = ""
        '    csr2$ = ""
            

             sSelect = "select idreceiptshdrwbalance from ReceiptsBalancePayments where IdReceiptsHDRPayBalance='" + recibo$ + "'"
             Rs.Open sSelect, base, adOpenUnspecified
             id_receiptshdrwbalance$ = Rs(0)
             Rs.Close
             
             
             user1$ = ""
              csr1$ = ""
             csr2$ = ""
             
            
             
      '       sSelect = "select idemployeeUSR, idemployeeCSR1, idemployeeCSR2 from ReceiptsBalancePayments recbalpay " & _
      '       "inner join ReceiptsHDR rechdr on recbalpay.IdReceiptsHDRWBalance=rechdr.IDReceiptHDR Where IdReceiptsHDRPayBalance='" + recibo$ + "'"
             
      '       Rs.Open sSelect, base, adOpenUnspecified
      '       user1$ = Rs(0)
      '       csr1$ = Rs(1)
      '       csr2$ = Rs(2)
      '       Rs.Close
             
             
                          
            
                             
              
              invoice_3_users = 1
              
              
                
                
          
          
         
 
           sSelect = "select idemployeeUSR, IdEmployeeCSR1, IdEmployeeCSR2  from [ReceiptsDTL] recdtl " & _
           "inner join  ReceiptsHDR  rechdr on rechdr.IDReceiptHDR =recdtl.IdReceiptHDR " & _
           "inner join InvoiceItemCatalog iitem on iitem.IdInvoiceItem=recdtl.IdInvoiceItem " & _
           "where rechdr.Active=1 and recdtl.[IdReceiptHDR] in (" + recibo$ + ")"
     
      
           Rs.Open sSelect, base, adOpenUnspecified
           user1$ = Rs(0)
           csr1$ = Rs(1)
           csr2$ = Rs(2)
           
           
           If UCase(agente$) = "JENNIFERR" Then
              If csr1$ <> "244" And csr2$ = "244" Then
                 csr1$ = "244"
                 
              End If
           End If
           
           
           Rs.Close
     
     
           sSelect = "select username from EmployeeInfo where IDEmployee='" + user1$ + "'"
           Rs.Open sSelect, base, adOpenUnspecified
           user$ = UCase(Rs(0))
           Rs.Close
      
      
           sSelect = "select username from EmployeeInfo where IDEmployee='" + csr1$ + "'"
           Rs.Open sSelect, base, adOpenUnspecified
           csr$ = UCase(Rs(0))
           Rs.Close
      
           
           NO_CAMBIAR = True
           
       
          ' +++++++++++  SE AGREGO ESTO Y SE DESHABILITO LO DE ARRIBA
            'X = agentes_de_invoice(recibo$)
   
          
            'csr$ = UCase(csr_original$)
            'user$ = UCase(user_original$)
        ' +++++++++++++++++++++++++++++++++++++++++++++++++++
      
      
      
      
                      
          
          
 
        End If
      
      
   End If
   
   
   ' == E N D ===================================================================
   
aaa:
   
   
   oficina_user$ = ""
   oficina_user2$ = ""
   oficina_csr$ = ""
   oficina_csr2$ = ""
   
   ' asigna la oficina a user y a csr y a agente,
   
   For Y = 1 To Val(ubicacion(0, 0))
      If UCase(RTrim(ubicacion(Y, 0))) = UCase(csr$) Then
         oficina_csr$ = ubicacion(Y, 1)
      End If
      
      If UCase(RTrim(ubicacion(Y, 0))) = UCase(user$) Then
         oficina_user$ = ubicacion(Y, 1)
      End If
      
      If UCase(RTrim(ubicacion(Y, 0))) = UCase(agente$) Then
         oficina_agente$ = ubicacion(Y, 1)
      End If
   Next Y
      
      
   'GoTo brincado
   
   
   
   
   ' *********************************************************************************************************************************************************
   ' ***   AGENTE = USER     ***
   ' ***************************
   
   
   If (UCase(agente$) = UCase(user$) And UCase(agente$) <> UCase(csr$)) Then
     
 '     If UCase(user$) = "clugo" Then Stop
 
     Grid2.col = 1
     recibo$ = Grid2.Text
     
     Grid2.col = 3
     concepto$ = Grid2.Text
     
     
     
    
       
     
     
     If invoice_3_users = 0 Then
       
       
       Grid2.col = 4
       userCSR$ = Grid2.Text
     
       Grid2.col = 5  ' cambie el 4 x 5
       userx$ = RTrim(Grid2.Text)
       
       
        
       
       
     
     Else
       
       
           userx$ = csr$
           userCSR$ = user$
         
       
     
     End If
     
     
     
     
     ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
   
   If (UCase$(LTrim(concepto$)) = "INVOICE") Then
   
     If NO_CAMBIAR = False Then
         X = agentes_de_invoice(recibo$)
     End If
   
     userxx$ = userx$
     
     
     If NO_CAMBIAR = False Then
       If UCase$(userxx$) = UCase$(user_original$) Or UCase(userxx$) = UCase(csr_original$) Then
     
       Else
          userx$ = user_original$
          csr$ = csr_original$   ' este estaba anulado
     
       End If
     End If
     
     
   End If
   
   
   ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        
     
     
     
     Grid2.col = 6
     cantidad$ = Grid2.Text
     
     Grid2.col = 9
     cantidad_pagada$ = Grid2.Text
     
     
     
     
            
           
            
            
            If Val(cantidad$) > Val(cantidad_pagada$) And Val(cantidad_pagada$) > 0 And (UCase$(LTrim(concepto$)) = "INVOICE") Then
                cantidad$ = cantidad_pagada$
            End If
            
     
     
     Grid2.col = 9
     cantidad_pagada$ = Grid2.Text
     
     If UCase(concepto$) <> "INVOICE" Then
     
                            r$ = Left(userx$, 2) + LCase(Right(userx$, Len(userx$) - 2))
                            
                            If UCase(RTrim(userx$)) = agente$ And csr$ <> agente$ Then
                                r$ = Left(csr$, 2) + LCase(Right(csr$, Len(csr$) - 2))
                            ElseIf UCase(RTrim(userx$)) = agente$ And csr$ = agente$ Then
                                r$ = Left(userx$, 2) + LCase(Right(userx$, Len(userx$) - 2))
                            End If
       If Val(cantidad$) > 0 Then
         linea$ = Format(concepto$, "!@@@@@@@@@@@@@@@@@@@@") + Format(recibo$, "000000") + " " + Format(r$, "!@@@@@@@@@@@@@@@@@@@@") + " " + Format(cantidad$, "000000.00")
       End If
        
        
     Else
       If Val(cantidad_pagada$) > Val(cantidad$) Then
           cantidad_pagada$ = cantidad$
       End If
       
       
                            r$ = Left(userx$, 2) + LCase(Right(userx$, Len(userx$) - 2))
                            
                            If UCase(RTrim(userx$)) = agente$ And csr$ <> agente$ Then
                                r$ = Left(csr$, 2) + LCase(Right(csr$, Len(csr$) - 2))
                            ElseIf UCase(RTrim(userx$)) = agente$ And csr$ = agente$ Then
                                r$ = Left(userx$, 2) + LCase(Right(userx$, Len(userx$) - 2))
                            End If
                            
        If Val(cantidad_pagada$) > 0 Then
           linea$ = Format(concepto$, "!@@@@@@@@@@@@@@@@@@@@") + Format(recibo$, "000000") + " " + Format(r$, "!@@@@@@@@@@@@@@@@@@@@") + " " + Format(cantidad_pagada$, "000000.00")
        End If
       
        
     End If
     
     
     
    
     '  >>>>>>>>>>>>>>>>>>>>>>>>>>>>   VERIFICA AQUI SI ES MANAGER
     
     GoSub check_manager
     
     '  >>>>>>>>>>>>>>>>>>>>>>>>>>>>   TERMINA AQUI SI ES MANAGER
     
     
     
     
     
     
     If existe = 0 Then
        If (UCase$(LTrim(concepto$)) = "INVOICE" Or UCase$(LTrim(concepto$)) = "LATE FEE") Then
        
        X = agentes_de_invoice(recibo$)
   
     
                If agente$ = UCase(csr_original$) Or agente$ = UCase(user_original$) Then
                   csr$ = UCase(csr_original$)
                   user$ = UCase(user_original$)
                ElseIf agente$ <> UCase(csr_original$) And agente$ <> UCase(user_original$) Then
                   GoTo No_encontrado
                End If
                
                existe = 0
                Es_Regio = 0
                
                For k = 0 To 500
                  If UCase(ubicacion(k, 0)) = UCase(csr$) Then
                     oficinaCSR1$ = UCase(ubicacion(k, 1))
                     oficinaCSR2$ = UCase(ubicacion(k, 2))
                     If oficinaCSR1$ = "JA - MONTERREY" Or oficinaCSR2$ = "JA - MONTERREY" Then
                        Es_Regio = 1
                     End If
                     existe = existe + 1
                  End If
                  
                  
                  If UCase(ubicacion(k, 0)) = UCase(user$) Then
                     oficinaUSER1$ = UCase(ubicacion(k, 1))
                     oficinaUSER2$ = UCase(ubicacion(k, 2))
                     If oficinaUSER1$ = "JA - MONTERREY" Or oficinaUSER2$ = "JA - MONTERREY" Then
                        Es_Regio = 1
                     End If
                     existe = existe + 1
                     
                  End If
                  
                  If existe >= 2 Then Exit For
                  
                Next k
                
                If oficinaCSR2$ = "" Then oficinaCSR2$ = "None"
                If oficinaUSER2$ = "" Then oficinaUSER2$ = "Nada"
        
         
         If Es_Monterrey = 1 Then
          
          List9.AddItem Str(Val(cantidad$))
          total_facturas_ajenas = total_facturas_ajenas + (Val(cantidad$))
         
         Else
         
         ' verifica oficinas a que pertenecen
                  For w = 1 To Val(ubicacion(0, 0))
                      If UCase(RTrim(ubicacion(w, 0))) = UCase(user$) Then
                          oficina_user$ = RTrim(ubicacion(w, 1))
                          oficina_user2$ = RTrim(ubicacion(w, 2))
                      End If
                  
                      If UCase(RTrim(ubicacion(w, 0))) = UCase(csr$) Then
                          oficina_csr$ = RTrim(ubicacion(w, 1))
                          oficina_csr2$ = RTrim(ubicacion(w, 2))
                      End If
                  Next w
                  
                  If oficina_user2$ = "None" And oficina_csr2$ = "None" Then
                      oficina_user2$ = "Nada"
                  End If
                  
                  
                  
                  
                  If oficina_user$ = oficina_csr$ Or oficina_user2$ = oficina_csr2$ Or oficina_user$ = oficina_csr2$ Or oficina_user2$ = oficina_csr$ Then
                  
                         If user$ = manager_user$ Or agente$ = manager_agente$ Or csr$ = manager_csr$ Or agente$ = manager_commercial$ Or csr$ = manager_commercial$ Then
                         
                         Else
                              List10.AddItem Str(Val(cantidad$))
                         End If
                              
                  Else                                                 ' manager y vendedor son de diferente oficina
                              List9.AddItem Str(Val(cantidad$) / 2)
                              List10.AddItem Str(Val(cantidad$) / 2)
                              total_facturas_ajenas = total_facturas_ajenas + (Val(cantidad$) / 2)
                  End If
                  
                  
         
          'List9.AddItem Str(Val(cantidad$) / 2)
          'List10.AddItem Str(Val(cantidad$) / 2)
          'total_facturas_ajenas = total_facturas_ajenas + (Val(cantidad$) / 2)
          
          
         End If
         
        End If
        
        
     Else
        If (UCase$(LTrim(concepto$)) = "INVOICE" Or UCase$(LTrim(concepto$)) = "LATE FEE") Then
        
          X = agentes_de_invoice(recibo$)
   
     
                If agente$ = UCase(csr_original$) Or agente$ = UCase(user_original$) Then
                   csr$ = UCase(csr_original$)
                   user$ = UCase(user_original$)
                ElseIf agente$ <> UCase(csr_original$) And agente$ <> UCase(user_original$) Then
                   GoTo No_encontrado
                End If
                
                existe = 0
                Es_Regio = 0
                
                For k = 0 To 500
                  If UCase(ubicacion(k, 0)) = UCase(csr$) Then
                     oficinaCSR1$ = UCase(ubicacion(k, 1))
                     oficinaCSR2$ = UCase(ubicacion(k, 2))
                     If oficinaCSR1$ = "JA - MONTERREY" Or oficinaCSR2$ = "JA - MONTERREY" Then
                        Es_Regio = 1
                     End If
                     existe = existe + 1
                  End If
                  
                  
                  If UCase(ubicacion(k, 0)) = UCase(user$) Then
                     oficinaUSER1$ = UCase(ubicacion(k, 1))
                     oficinaUSER2$ = UCase(ubicacion(k, 2))
                     If oficinaUSER1$ = "JA - MONTERREY" Or oficinaUSER2$ = "JA - MONTERREY" Then
                        Es_Regio = 1
                     End If
                     existe = existe + 1
                     
                  End If
                  
                  If existe >= 2 Then Exit For
                  
                Next k
                
                
                If oficinaCSR2$ = "" Then oficinaCSR2$ = "None"
                If oficinaUSER2$ = "" Then oficinaUSER2$ = "Nada"
                
                
           If agente$ = user$ And csr$ <> manager_csr$ And user$ <> manager_user$ And user$ <> csr$ And Es_Regio = 0 Then
                 List9.AddItem Str(Val(cantidad$) / 2)
                 List10.AddItem Str(Val(cantidad$) / 2)
                 total_facturas_ajenas = total_facturas_ajenas + (Val(cantidad$) / 2)
           
           
           ElseIf agente$ = user$ And agente$ = manager_agente$ And agente$ <> csr$ Then
                  
                  ' verifica oficinas a que pertenecen
                  For w = 1 To Val(ubicacion(0, 0))
                      If UCase(RTrim(ubicacion(w, 0))) = UCase(user$) Then
                          oficina_user$ = RTrim(ubicacion(w, 1))
                          oficina_user2$ = RTrim(ubicacion(w, 2))
                      End If
                  
                      If UCase(RTrim(ubicacion(w, 0))) = UCase(csr$) Then
                          oficina_csr$ = RTrim(ubicacion(w, 1))
                          oficina_csr2$ = RTrim(ubicacion(w, 2))
                      End If
                  Next w
                  
                  If oficina_user2$ = "None" And oficina_csr2$ = "None" Then
                      oficina_user2$ = "Nada"
                  End If
                  
                  If oficina_user$ = oficina_csr$ Or oficina_user2$ = oficina_csr2$ Or oficina_user$ = oficina_csr2$ Or oficina_user2$ = oficina_csr$ Then
                              List10.AddItem Str(Val(cantidad$))
                              
                  Else                                                 ' manager y vendedor son de diferente oficina
                              List9.AddItem Str(Val(cantidad$) / 2)
                              List10.AddItem Str(Val(cantidad$) / 2)
                              total_facturas_ajenas = total_facturas_ajenas + (Val(cantidad$) / 2)
                  End If
                                    
                  
                  
           
           
           ElseIf agente$ = user$ And csr$ = manager_csr$ And agente$ <> csr$ Then
                  
                   ' verifica oficinas a que pertenecen
                  For w = 1 To Val(ubicacion(0, 0))
                      If UCase(RTrim(ubicacion(w, 0))) = UCase(user$) Then
                          oficina_user$ = RTrim(ubicacion(w, 1))
                          oficina_user2$ = RTrim(ubicacion(w, 2))
                      End If
                  
                      If UCase(RTrim(ubicacion(w, 0))) = UCase(csr$) Then
                          oficina_csr$ = RTrim(ubicacion(w, 1))
                          oficina_csr2$ = RTrim(ubicacion(w, 2))
                      End If
                  Next w
                  
                  
                  If oficina_user$ = oficina_csr$ Or oficina_user2$ = oficina_csr2$ Or oficina_user$ = oficina_csr2$ Or oficina_user2$ = oficina_csr$ Then
                              List9.AddItem Str(Val(cantidad$))
                              total_facturas_ajenas = total_facturas_ajenas + (Val(cantidad$))
                  Else
                              List9.AddItem Str(Val(cantidad$) / 2)
                              List10.AddItem Str(Val(cantidad$) / 2)
                              total_facturas_ajenas = total_facturas_ajenas + (Val(cantidad$) / 2)
                  End If
                  
                  
           
           
           
           ElseIf agente$ = user$ And user$ = manager_user$ And csr$ = manager_csr$ And csr$ <> user$ Then
           
                  ' verifica oficinas a que pertenecen
                    ' verifica oficinas a que pertenecen
                  For w = 1 To Val(ubicacion(0, 0))
                      If UCase(RTrim(ubicacion(w, 0))) = UCase(user$) Then
                          oficina_user$ = RTrim(ubicacion(w, 1))
                          oficina_user2$ = RTrim(ubicacion(w, 2))
                      End If
                  
                      If UCase(RTrim(ubicacion(w, 0))) = UCase(csr$) Then
                          oficina_csr$ = RTrim(ubicacion(w, 1))
                          oficina_csr2$ = RTrim(ubicacion(w, 2))
                      End If
                  Next w
                  
                  
                  If oficina_user$ = oficina_csr$ Or oficina_user2$ = oficina_csr2$ Or oficina_user$ = oficina_csr2$ Or oficina_user2$ = oficina_csr$ Then
                           List9.AddItem Str(Val(cantidad$) / 2)
                           List10.AddItem Str(Val(cantidad$) / 2)
                           total_facturas_ajenas = total_facturas_ajenas + (Val(cantidad$) / 2)
                  Else
                           List9.AddItem Str(Val(cantidad$) / 2)
                           List10.AddItem Str(Val(cantidad$) / 2)
                           total_facturas_ajenas = total_facturas_ajenas + (Val(cantidad$) / 2)
                  End If
                  
                  
           
           ElseIf agente$ = csr$ And agente$ = user$ And agente$ <> manager_agente$ Then
                      
                  List9.AddItem cantidad$
                  total_facturas_propias = total_facturas_propias + Val(cantidad$)
           
           ElseIf agente$ = csr$ And agente$ = user$ And agente$ = manager_agente$ Then
                  
                  List9.AddItem cantidad$
                  total_facturas_propias = total_facturas_propias + Val(cantidad$)
           
           ElseIf (agente$ = csr$ Or agente$ = user$) And Es_Regio = 1 Then
                    
               List9.AddItem Str(Val(cantidad$))
               total_facturas_ajenas = total_facturas_ajenas + (Val(cantidad$))
              
              
           End If
           
        End If
     End If
     
     
     
              
              
     ' revisa si el Late fee es parte de un Invoice
     If LTrim(UCase(concepto$)) = "LATE FEE" Then
         existe_invoice = 0
         For w = 0 To List2.ListCount - 1
            a1$ = RTrim(LTrim(UCase(Left(List2.List(w), 20))))
            b1$ = LTrim(RTrim(Mid$(List2.List(w), 21, 6)))
            
            If a1$ = "INVOICE" And b1$ = recibo$ Then
                 cantidad_del_invoice = Val(Right(List2.List(w), 9))
                 existe_invoice = 1
                 Exit For
            End If
         Next w
         
         If existe_invoice = 1 Then
              If Format(cantidad_del_invoice, "###0.0") = Format(Val(cantidad_pagada$), "###0.0") Then
                 anula_LATEFEE = 1
              End If
         End If
         
     End If
              
     If anula_LATEFEE = 1 Then
        
     Else
        List2.AddItem linea$
        
     End If
     anula_LATEFEE = 0
   
   End If
  
  
  
 
  
brincado:
  
  
  
  
  
   ' *********************************************************************************************************************************************************
   ' ***   AGENTE = CSR     ***
   ' ***************************
  
   If UCase(agente$) = UCase(csr$) Then
     
     ' If UCase(user$) = "JMIRELES" Then Stop

     Grid2.col = 1
     recibo$ = Grid2.Text
     
     'If recibo$ = "189131" Then Stop
        
     
     Grid2.col = 3
     concepto$ = Grid2.Text
     
    
       
     
     
     If invoice_3_users = 0 Then
       Grid2.col = 4
       userx$ = RTrim(Grid2.Text)    ' USER
     Else
       userx$ = user$
     End If
     
     
     Grid2.col = 6
     cantidad$ = Grid2.Text
     
     
     Grid2.col = 9
     cantidad_pagada$ = Grid2.Text
     
     Grid2.col = 11
     balance_que_se_debe$ = Grid2.Text
     
     
     
     
     ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
     
    If (UCase$(LTrim(concepto$)) = "INVOICE") Or (UCase$(LTrim(concepto$)) = "LATE FEE") Then
   
     X = agentes_de_invoice(recibo$)
   
     userx$ = user_original$  ' este estaba anulado
     csr$ = csr_original$   ' este estaba anulado
     
    End If
        
   
   ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        
     
     
     
     
     
     
     
         cantidad_de_balance$ = ""
     
         sSelect = "select balancedue from ReceiptsBalancePayments recbalpay " & _
         "inner join ReceiptsHDR rechdr on recbalpay.IdReceiptsHDRWBalance=rechdr.IDReceiptHDR " & _
         "Where IdReceiptsHDRPayBalance='" + recibo$ + "'"
         
          ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
          Rs.Open sSelect, base, adOpenUnspecified
    
          cantidad_de_balance$ = Rs(0)
          
          Rs.Close
          
          
                  
        If (UCase$(LTrim(concepto$)) = "INVOICE") Or (UCase$(LTrim(concepto$)) = "LATE FEE") Then
          ' calcula
          carga_factura
          factura = 0
          recibo8x$ = recibo$
          poliza8x$ = ""
          suma_invoice = 0
          suma_late_fee = 0
          existe = False
          primer_recibo$ = ""
          asigna_cantidad = False
          total_invoices = 0
          
          For w = 1 To grid8.Rows - 1
             grid8.row = w
             grid8.col = 1
             recibo8$ = grid8.Text
                          
             grid8.col = 3
             concepto8$ = grid8.Text
             
             If UCase(concepto8$) = "INVOICE" Then
                total_invoices = total_invoices + 1
                cantidad_de_factura$ = cantidad_pagada$
             End If
             
             grid8.col = 6
             cant_original8 = Val(grid8.Text)
             
             grid8.col = 9
             cant_pagada8 = Val(grid8.Text)
             
             grid8.col = 11
             balance8 = Val(grid8.Text)
             
             grid8.col = 12
             custid8$ = grid8.Text
             
             grid8.col = 13
             poliza8$ = grid8.Text
             
             Grid2.col = 9
             cantidad_pagada$ = Grid2.Text
     
             
             If existe = False Then
                existe = True
                primer_recibo$ = recibo8$
                
                If balance8 < cant_original8 Then
                   total_cantidad_pagada = cant_original8 - balance8
                Else
                  
                   total_cantidad_pagada = cant_original8
                End If
             End If
             
             
             If recibo8$ <> recibo8x$ And recibo8x$ <> "" Then
                  If balance8 = 0 Then
                     asigna_cantidad = True
                  End If
                 
             ElseIf recibo8$ <> recibo8x$ And recibo8x$ = "" Then
             
                 recibo8x$ = recibo8$
                 
                 
                 
                 If UCase(concepto8$) = "INVOICE" Then
                     suma_invoice = suma_invoice + cant_original8
                 Else
                     suma_late_fee = suma_late_fee + cant_original8
                 End If
                 asigna_cantidad = True
             ElseIf recibo8$ = recibo8x$ And recibo8$ = primer_recibo$ Then
                 
                 If UCase(concepto8$) = "INVOICE" Then
                     suma_invoice = suma_invoice + cant_original8
                 Else
                     suma_late_fee = suma_late_fee + cant_original8
                 End If
                asigna_cantidad = True
             End If
             
                         
          Next w
          
          If UCase$(concepto$) = "INVOICE" Then
              cantidad$ = cantidad_pagada$
              cant_pagada8 = Val(cantidad$)
          End If
          
          
                    
          
          
          If total_invoices = 1 And asigna_cantidad = True Then
          
             
           
             '   If ((suma_invoice + suma_late_fee) - cant_pagada8) = balance8 And balance8 <> 0 And suma_late_fee > 0 Then
                      
                       diferencia = suma_invoice - cant_pagada8
                       If diferencia = 0 Then   ' se pago toda la factura
                          
                          cantidad_de_balance$ = Format(suma_late_fee, "####0.00")
                          suma_late_fee = 0
                          
                          If (UCase$(LTrim(concepto$)) = "INVOICE") Then
                            cantidad$ = cant_pagada8
                          ElseIf (UCase$(LTrim(concepto$)) = "LATE FEE") Then
                            cantidad$ = "0"
                          End If
                            
                       ElseIf diferencia > 0 Then
                       
                         cantidad_de_balance$ = Format(diferencia + suma_late_fee, "####0.00")
                         suma_invoice = cant_pagada8
                         suma_late_fee = 0
                         
                         If (UCase$(LTrim(concepto$)) = "INVOICE") Then
                            cantidad$ = cant_pagada8
                         ElseIf (UCase$(LTrim(concepto$)) = "LATE FEE") Then
                            cantidad$ = "0"
                         End If
                          
                         
                       ElseIf diferencia < 0 Then
                       
                         cantidad_de_balance$ = Format(suma_late_fee - (diferencia * -1), "####0.00")
                         suma_late_fee = diferencia * -1
                         If (UCase$(LTrim(concepto$)) = "INVOICE") Then
                            cantidad$ = cant_pagada8
                         ElseIf (UCase$(LTrim(concepto$)) = "LATE FEE") Then
                            cantidad$ = suma_late_fee
                         End If
                         
                       Else
                       
                       End If
                       
                        
                       'suma_invoice = 0
                       'suma_late_fee = cant_pagada8
                       'cantidad_de_balance$ = Format(suma_invoice, "####0.00")
                       factura = 1
              '  Else
                      ' suma_invoice = total_cantidad_pagada
               ' End If
             
          End If
          
          If asigna_cantidad = False Then
             asigna_cantidad = True
             GoTo No_encontrado    'saltado_sin_poner_nada
             
          End If
             
        End If
            
            
          If Val(Format(cantidad$, "00000.00")) <> Val(Format(cantidad_de_balance$, "00000.00")) And cantidad_de_balance$ <> "" And (UCase$(LTrim(concepto$)) = "INVOICE") Then
             cantidad$ = cantidad_de_balance$
          End If
            
            
            
            
            
          If suma_invoice > 0 And cant_original8 >= suma_invoice And Val(cantidad_de_balance$) >= 0 And (UCase$(LTrim(concepto$)) = "INVOICE") Then
             cantidad$ = Format(suma_invoice, "#####0.00")
             GoTo ve_directo
          ElseIf suma_late_fee > 0 And (UCase$(LTrim(concepto$)) = "LATE FEE") Then
             cantidad$ = Format(suma_late_fee, "#####0.00")
             GoTo ve_directo
          End If
            
            
            
          If Val(Format(cantidad_de_balance$, "00000.00")) > 0 And (UCase$(LTrim(concepto$)) = "INVOICE") Then
            cantidad$ = Format(Val(cantidad_de_balance$), "###,##0.00")
            
          End If
          
          
            
            
            If Val(cantidad$) > Val(cantidad_pagada$) And Val(cantidad_pagada$) > 0 And (UCase$(LTrim(concepto$)) = "INVOICE") Then
                cantidad$ = cantidad_pagada$
            End If
            
            
            If Val(cantidad_de_balance$) > Val(balance_que_se_debe$) And (UCase$(LTrim(concepto$)) = "INVOICE") Then
                cantidad$ = Format(Val(cantidad_de_balance$) - Val(balance_que_se_debe$), "######0.00")
            End If
            
     
     
          
     
     
     
ve_directo:
     
     If UCase(concepto$) <> "INVOICE" Then
     
                            r$ = Left(userx$, 2) + LCase(Right(userx$, Len(userx$) - 2))
                            
                            If UCase(RTrim(userx$)) = agente$ And csr$ <> agente$ Then
                                r$ = Left(csr$, 2) + LCase(Right(csr$, Len(csr$) - 2))
                            ElseIf UCase(RTrim(userx$)) = agente$ And csr$ = agente$ Then
                                r$ = Left(userx$, 2) + LCase(Right(userx$, Len(userx$) - 2))
                            End If
                            
        If Val(cantidad$) > 0 Then
         linea$ = Format(concepto$, "!@@@@@@@@@@@@@@@@@@@@") + Format(recibo$, "000000") + " " + Format(r$, "!@@@@@@@@@@@@@@@@@@@@") + " " + Format(cantidad$, "000000.00")
         linea2$ = linea$
        End If
        
       
     Else
       If Val(cantidad_pagada$) > Val(cantidad$) And factura = 0 Then
           cantidad_pagada$ = cantidad$
       End If
       
                            r$ = Left(userx$, 2) + LCase(Right(userx$, Len(userx$) - 2))
                            
                            If UCase(RTrim(userx$)) = agente$ And csr$ <> agente$ Then
                                r$ = Left(csr$, 2) + LCase(Right(csr$, Len(csr$) - 2))
                            ElseIf UCase(RTrim(userx$)) = agente$ And csr$ = agente$ Then
                                r$ = Left(userx$, 2) + LCase(Right(userx$, Len(userx$) - 2))
                            End If
       
       ' linea$ = Format(concepto$, "!@@@@@@@@@@@@@@@@@@@@") + Format(recibo$, "000000") + " " + Format(r$, "!@@@@@@@@@@@@@@@@@@@@") + " " + Format(cantidad_pagada$, "000000.00")
      If Val(cantidad$) > 0 Then
       linea$ = Format(concepto$, "!@@@@@@@@@@@@@@@@@@@@") + Format(recibo$, "000000") + " " + Format(r$, "!@@@@@@@@@@@@@@@@@@@@") + " " + Format(cantidad$, "000000.00")
      End If
        
     End If
     
saltado_sin_poner_nada:
     
     
     
     
     ' ++++++++++++++++++++++++++++++++   REVISA SI ES MANAGER
     
      GoSub check_manager
      
           
     ' +++++++++++++++++++++++++++++++   TERMINA
     
   
     
     
     
     If existe = 0 Then
     
        If (UCase$(LTrim(concepto$)) = "INVOICE" Or UCase$(LTrim(concepto$)) = "LATE FEE") Then
        
                 X = agentes_de_invoice(recibo$)
   
     
                If agente$ = UCase(csr_original$) Or agente$ = UCase(user_original$) Then
                   csr$ = UCase(csr_original$)
                   user$ = UCase(user_original$)
                ElseIf agente$ <> UCase(csr_original$) And agente$ <> UCase(user_original$) Then
                   GoTo No_encontrado
                End If
                
                
                 existe = 0
                Es_Regio = 0
                
                For k = 0 To 500
                  If UCase(ubicacion(k, 0)) = UCase(csr$) Then
                     oficinaCSR1$ = UCase(ubicacion(k, 1))
                     oficinaCSR2$ = UCase(ubicacion(k, 2))
                     If oficinaCSR1$ = "JA - MONTERREY" Or oficinaCSR2$ = "JA - MONTERREY" Then
                        Es_Regio = 1
                     End If
                     existe = existe + 1
                  End If
                  
                  
                  If UCase(ubicacion(k, 0)) = UCase(user$) Then
                     oficinaUSER1$ = UCase(ubicacion(k, 1))
                     oficinaUSER2$ = UCase(ubicacion(k, 2))
                     If oficinaUSER1$ = "JA - MONTERREY" Or oficinaUSER2$ = "JA - MONTERREY" Then
                        Es_Regio = 1
                     End If
                     existe = existe + 1
                     
                  End If
                  
                  If existe >= 2 Then Exit For
                  
                Next k
                
                
                
                If oficinaCSR2$ = "" Then oficinaCSR2$ = "None"
                If oficinaUSER2$ = "" Then oficinaUSER2$ = "Nada"
                
                
                  
                 If oficinaUSER1$ = oficinaCSR1$ Or oficinaUSER1$ = oficinaCSR2$ Or oficinaUSER2$ = oficinaCSR1$ Or oficinaUSER2$ = oficinaCSR2$ Then
                    
                    If (UCase(user$) = UCase(csr$)) Or ((manager_commercial$ <> agente$) And (manager_user$ = user$ Or manager_csr$ = user$ Or manager_agente$ = user$ Or manager_commercial$ = user$) Or (manager_user$ = csr$ Or manager_csr$ = csr$ Or manager_agente$ = csr$ Or manager_commercial$ = csr$)) Then
                      List9.AddItem cantidad$
                      If (UCase(user$) <> UCase(csr$)) Then
                        total_facturas_ajenas = total_facturas_ajenas + Val(cantidad$)
                      Else
                        total_facturas_propias = total_facturas_propias + Val(cantidad$)
                      End If
                      
                    Else
                      List9.AddItem Str(Val(cantidad$) / 2)
                      List10.AddItem Str(Val(cantidad$) / 2)
                      total_facturas_ajenas = total_facturas_ajenas + (Val(cantidad$) / 2)
                    
                    End If
                 
                 
                 Else
                 
        
                    If UCase(user$) = UCase(csr$) Then
                      List9.AddItem cantidad$
                      total_facturas_propias = total_facturas_propias + Val(cantidad$)
                    Else
                      List9.AddItem Str(Val(cantidad$) / 2)
                      List10.AddItem Str(Val(cantidad$) / 2)
                      total_facturas_ajenas = total_facturas_ajenas + (Val(cantidad$) / 2)
                    End If
               
               End If
          
        End If
      
     Else
     
        
       
         
     
     
        If (UCase$(LTrim(concepto$)) = "INVOICE") Or (UCase$(LTrim(concepto$)) = "LATE FEE") Then
        
        
               X = agentes_de_invoice(recibo$)
               
               
     
                If agente$ = UCase(csr_original$) Or agente$ = UCase(user_original$) Then
                   csr$ = UCase(csr_original$)
                   user$ = UCase(user_original$)
                ElseIf agente$ <> UCase(csr_original$) And agente$ <> UCase(user_original$) Then
                   GoTo No_encontrado
                End If
                
                
                existe = 0
                Es_Regio = 0
                
                For k = 0 To 500
                  If UCase(ubicacion(k, 0)) = UCase(csr$) Then
                     oficinaCSR1$ = UCase(ubicacion(k, 1))
                     oficinaCSR2$ = UCase(ubicacion(k, 2))
                     If oficinaCSR1$ = "JA - MONTERREY" Or oficinaCSR2$ = "JA - MONTERREY" Then
                        Es_Regio = 1
                     End If
                     existe = existe + 1
                  End If
                  
                  
                  If UCase(ubicacion(k, 0)) = UCase(user$) Then
                     oficinaUSER1$ = UCase(ubicacion(k, 1))
                     oficinaUSER2$ = UCase(ubicacion(k, 2))
                     If oficinaUSER1$ = "JA - MONTERREY" Or oficinaUSER2$ = "JA - MONTERREY" Then
                        Es_Regio = 1
                     End If
                     existe = existe + 1
                     
                  End If
                  
                  If existe >= 2 Then Exit For
                  
                Next k
                
                
                If oficinaCSR2$ = "" Then oficinaCSR2$ = "None"
                If oficinaUSER2$ = "" Then oficinaUSER2$ = "Nada"
             
                
             
              If oficinaUSER1$ = oficinaCSR1$ Or oficinaUSER1$ = oficinaCSR2$ Or oficinaUSER2$ = oficinaCSR1$ Or oficinaUSER2$ = oficinaCSR2$ Then
                           
                    
                 If UCase(user$) <> UCase(csr$) And (manager_user$ <> user$ And manager_Oficinacommercial$ <> user$) And (manager_csr$ <> csr$ And manager_Oficinacommercial$ <> csr$) And Es_Regio = 0 Then
                      
                                        
                      List9.AddItem Str(Val(cantidad$) / 2)
                      List10.AddItem Str(Val(cantidad$) / 2)
                      total_facturas_ajenas = total_facturas_ajenas + (Val(cantidad$) / 2)
                      
                 ElseIf UCase(user$) <> UCase(csr$) And (manager_Oficinacommercial$ = user$ Or manager_Oficinacommercial$ = csr$) And Es_Regio = 0 Then
                      List9.AddItem Str(Val(cantidad$))
                      total_facturas_ajenas = total_facturas_ajenas + (Val(cantidad$))
                      
                 ElseIf UCase(user$) <> UCase(csr$) And (manager_Oficinacommercial$ <> user$ And manager_Oficinacommercial$ <> csr$) And Es_Regio = 0 Then
                      
                      List9.AddItem Str(Val(cantidad$) / 2)
                      List10.AddItem Str(Val(cantidad$) / 2)
                      total_facturas_ajenas = total_facturas_ajenas + (Val(cantidad$) / 2)
                      
                 ElseIf UCase(user$) = UCase(csr$) Then
                     List9.AddItem Str(Val(cantidad$))
                      total_facturas_propias = total_facturas_propias + (Val(cantidad$))
                      
                 End If

             ElseIf agente$ = manager_agente$ And agente$ = csr$ And csr$ <> user$ Then
              
                    ' verifica oficinas a que pertenecen
                  For w = 1 To Val(ubicacion(0, 0))
                      If UCase(RTrim(ubicacion(w, 0))) = UCase(user$) Then
                          oficina_user$ = RTrim(ubicacion(w, 1))
                          oficina_user2$ = RTrim(ubicacion(w, 2))
                      End If
                  
                      If UCase(RTrim(ubicacion(w, 0))) = UCase(csr$) Then
                          oficina_csr$ = RTrim(ubicacion(w, 1))
                          oficina_csr2$ = RTrim(ubicacion(w, 2))
                      End If
                  Next w
                  
                  
                  If oficina_user$ = oficina_csr$ Or oficina_user2$ = oficina_csr2$ Or oficina_user$ = oficina_csr2$ Or oficina_user2$ = oficina_csr$ Then
                          List10.AddItem cantidad$
                  Else
                          List9.AddItem Str(Val(cantidad$) / 2)
                          List10.AddItem Str(Val(cantidad$) / 2)
                          total_facturas_ajenas = total_facturas_ajenas + (Val(cantidad$) / 2)
                  End If
                     
                     
                     
                     
             ElseIf user$ = manager_user$ And agente$ = csr$ And csr$ <> user$ Then
             
                   ' verifica oficinas a que pertenecen
                  For w = 1 To Val(ubicacion(0, 0))
                      If UCase(RTrim(ubicacion(w, 0))) = UCase(user$) Then
                          oficina_user$ = RTrim(ubicacion(w, 1))
                          oficina_user2$ = RTrim(ubicacion(w, 2))
                      End If
                  
                      If UCase(RTrim(ubicacion(w, 0))) = UCase(csr$) Then
                          oficina_csr$ = RTrim(ubicacion(w, 1))
                          oficina_csr2$ = RTrim(ubicacion(w, 2))
                      End If
                  Next w
                  
                  
                  If oficina_user$ = oficina_csr$ Or oficina_user2$ = oficina_csr2$ Or oficina_user$ = oficina_csr2$ Or oficina_user2$ = oficina_csr$ Then
                          List9.AddItem Str(Val(cantidad$))
                          total_facturas_ajenas = total_facturas_ajenas + (Val(cantidad$))
                  Else
                          List9.AddItem Str(Val(cantidad$) / 2)
                          List10.AddItem Str(Val(cantidad$) / 2)
                          total_facturas_ajenas = total_facturas_ajenas + (Val(cantidad$) / 2)
                  End If
                     
             ElseIf user$ = manager_user$ And agente$ = csr$ And csr$ <> user$ And csr$ = manager_csr$ Then
                      List9.AddItem Str(Val(cantidad$) / 2)
                      List10.AddItem Str(Val(cantidad$) / 2)
                      total_facturas_propias = total_facturas_propias + (Val(cantidad$) / 2)
                              
                              
             ElseIf user$ = agente$ And agente$ = csr$ Then
                              
                       List9.AddItem Str(Val(cantidad$))
                       total_facturas_propias = total_facturas_propias + (Val(cantidad$))
                       
                       
             ElseIf (agente$ = csr$ Or agente$ = user$) And Es_Regio = 1 Then
                    
               List9.AddItem Str(Val(cantidad$))
               total_facturas_ajenas = total_facturas_ajenas + (Val(cantidad$))
              

             Else
                     List10.AddItem cantidad$
             End If
         
         
        End If
      
     End If
     
     existe = 0
     recibo_almacenado$ = Mid$(linea$, 21, 6)
     cantidad_almacenada$ = LTrim(RTrim(Right(linea$, 11)))
     
    
     ' a$ = ID_Cliente$
   
     
     
     For vv = 0 To List2.ListCount - 1
        recibo_procesado$ = Mid$(List2.List(vv), 21, 6)
        cantidad_procesada$ = LTrim(RTrim(Right(List2.List(vv), 11)))
        concepto_procesado$ = RTrim(Left(UCase(List2.List(vv)), 20))
        
        
        
                
        If UCase(concepto_procesado$) = "INVOICE" Then
           ' busca el custID
               hayado = 0
               For J = 1 To Grid2.Rows - 1
                   Grid2.row = J
                   Grid2.col = 1
                   recib$ = Grid2.Text
            
                   If recibo_procesado$ = recib$ Then
                       Grid2.col = 12
                       id_cte_procesado$ = Grid2.Text
                       hayado = 1
                       Exit For
                   End If
            
               Next J
         
        End If
        
        
        
        
             
        If List2.List(vv) = linea$ Then
           existe = 1
           Exit For
        ElseIf ID_Cliente$ = id_cte_procesado$ And UCase(concepto$) = "INVOICE" Then
           existe = 1
           Exit For
        End If
     Next vv
     
     
     If existe = 0 Then
        List2.AddItem linea$
     Else
       'Stop
     End If
     
   End If
   

No_encontrado:
' revisa si es una factura cobrada por RECAUDACION
' ========================================================================================================
   
   If (UCase$(LTrim(concepto$)) = "INVOICE") Or (UCase$(LTrim(concepto$)) = "LATE FEE") Then
       
       X = agentes_de_invoice(recibo$)
   
       userx$ = user_original$
       csr$ = csr_original$
       
       If UCase(user_original$) <> UCase(agente$) And UCase(csr_original$) <> UCase(agente$) And UCase(agente$) = "JENNIFERR" Then
           ' agrega esta poliza
           linea$ = Format(concepto$, "!@@@@@@@@@@@@@@@@@@@@") + Format(recibo$, "000000") + " " + Format(agente$, "!@@@@@@@@@@@@@@@@@@@@") + " " + Format(cantidad$, "000000.00")
           List2.AddItem linea$
           
           List9.AddItem Str(Val(cantidad$))
           total_facturas_ajenas = total_facturas_ajenas + (Val(cantidad$))
       End If

   End If


' ========================================================================================================




Next t








grand_total = 0

For t = 0 To List2.ListCount - 1
  concepto$ = Left(List2.List(t), 20)
  cantidad$ = Right(List2.List(t), 9)
  recibo$ = Mid$(List2.List(t), 21, 6)
  userx$ = Mid$(List2.List(t), 28, 20)
  
  'If Val(cantidad$) = 118 Then Stop
   'If recibo$ = "225916" Then Stop
  'If UCase(LTrim(RTrim(concepto$))) = "INVOICE" Then Stop
  
  ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
   
   If RTrim(LTrim(UCase(concepto$))) = "INVOICE" Or RTrim(LTrim(UCase(concepto$))) = "LATE FEE" Then
   
     X = agentes_de_invoice(recibo$)
   
       
             
     
     If UCase$(agente$) <> UCase$(user_original$) And UCase(agente$) <> UCase(csr_original$) Then
     
        
       
       If UCase(agente$) = "JENNIFERR" Then
           
       Else
         GoTo saltado
           
       End If
       
       
       
     
           
     End If
     
     
     ' Si esta compartido pasalo al CSR
     
       userx$ = user_original$   ' estaba anulado
       csr$ = csr_original$      ' estaba anulado
     
     
     userxx$ = userx$
     
     
       
    
  ' ---------------------------  DETECTA SI EL AGENTE FUE EL COBRADOR DEL CHEQUE ----------------------------------------
    If RTrim(LTrim(UCase(concepto$))) = "LATE FEE" Then
      
     agente_cobrador$ = ""
     cobrador$ = ""
 
     sSelect = "select idemployeeUSR, IdEmployeeCSR1, IdEmployeeCSR2  from [ReceiptsDTL] recdtl " & _
     "inner join  ReceiptsHDR  rechdr on rechdr.IDReceiptHDR =recdtl.IdReceiptHDR " & _
     "inner join InvoiceItemCatalog iitem on iitem.IdInvoiceItem=recdtl.IdInvoiceItem " & _
     "where rechdr.Active=1 and recdtl.[IdReceiptHDR] in (" + recibo$ + ")"
     
     ' and iitem.InvoiceItemName='Late Fee'"  se quito esta parte   4/2/2024
 
     Rs.Open sSelect, base, adOpenUnspecified
     agente_cobrador$ = Rs(0)
     agente_csr1$ = Rs(1)
     agente_csr1$ = Rs(2)
     Rs.Close
     
     
      sSelect = "select username from EmployeeInfo where IDEmployee='" + agente_cobrador$ + "'"
      Rs.Open sSelect, base, adOpenUnspecified
      cobrador$ = Rs(0)
      Rs.Close
      
      
      
      
      sSelect = "select username from EmployeeInfo where IDEmployee='" + agente_csr1$ + "'"
      Rs.Open sSelect, base, adOpenUnspecified
      csr$ = Rs(0)
      Rs.Close
      
      
      
      userx$ = cobrador$   ' estaba anulado
      ' CSR$ = cobrador$      ' estaba anulado
     
     
     userxx$ = userx$
     
     
     'concepto$ = Left(List2.List(t), 20)
   'recibo$ = Mid$(List2.List(t), 21, 6)
   'usuario$ = Mid$(List2.List(t), 28, 20)
   'Cantidad$ = Right(List2.List(t), 9)
   
    If Val(cantidad$) > 0 Then
      linea$ = Format(concepto$, "!@@@@@@@@@@@@@@@@@@@@") + Format(recibo$, "000000") + " " + Format(cobrador$, "!@@@@@@@@@@@@@@@@@@@@") + " " + Format(cantidad$, "000000.00")
    End If
   List2.RemoveItem t
   List2.AddItem linea$
     
   
    End If
  ' -------------------------------------------------------------------------------------------------------------------
   
   
   
     
   End If
   
   
   ' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
       
  
    
  
  
  If List1.ListCount = 0 Then
           If RTrim(LTrim(UCase(concepto$))) = "INVOICE" Or RTrim(LTrim(UCase(concepto$))) = "LATE FEE" Then
 
                            r$ = UCase(Left(userx$, 2)) + LCase(Right(userx$, Len(userx$) - 2))
                            
                            If UCase(RTrim(userx$)) = agente$ And csr$ <> agente$ Then
                                r$ = UCase(Left(csr$, 2)) + LCase(Right(csr$, Len(csr$) - 2))
                            ElseIf UCase(RTrim(userx$)) = agente$ And csr$ = agente$ Then
                                r$ = UCase(Left(userx$, 2)) + LCase(Right(userx$, Len(userx$) - 2))
                            End If
                            
           Else
                            r$ = RTrim(UCase(Left(userx$, 2)) + LCase(Right(userx$, Len(userx$) - 2)))
           End If
                        
  
         
  
        List1.AddItem Format(r$, "!@@@@@@@@@@@@@@@@@@@@") + " " + Format(concepto$, "!@@@@@@@@@@@@@@@@@@@@") + Space(1) + cantidad$
       
    
  Else
  
        encontrado = 0
        For Y = 0 To List1.ListCount - 1
               concep$ = Mid$(List1.List(Y), 22, 20)
               user2$ = Left(List1.List(Y), 20)
               'userx$ = user2$
               
               
    
                If UCase(RTrim(concep$)) = UCase(RTrim(concepto$)) And LTrim(UCase(RTrim(userx$))) = LTrim(UCase(RTrim(user2$))) Then
               
               
               ' If (UCase(RTrim(concep$)) = UCase(RTrim(concepto$))) And (UCase(RTrim(userx$)) = UCase(RTrim(user2$)) Or UCase(RTrim(user2$)) = UCase(RTrim(csr$))) Then
                        
                        
                       ' cant$ = Right(List1.List(Y), 9)
                       ' Total = Val(cantidad$) + Val(cant$)
                       ' List1.RemoveItem Y
                        
                        
                        If RTrim(LTrim(UCase(concepto$))) = "INVOICE" Or RTrim(LTrim(UCase(concepto$))) = "LATE FEE" Then
                               
                            If UCase(RTrim(userx$)) = UCase(RTrim(agente$)) And UCase(RTrim(csr$)) <> UCase(RTrim(agente$)) Then
                               If UCase(RTrim(user2$)) = UCase(RTrim(csr$)) Then
                                   cant$ = Right(List1.List(Y), 9)
                                   Total = Val(cantidad$) + Val(cant$)
                                   List1.RemoveItem Y
                                   GoTo r1
                               
                               End If
                               
                             cant$ = cantidad$
                             Total = Val(cant$)
                             userx$ = csr$
                             r$ = UCase(Left(userx$, 2)) + LCase(Right(userx$, Len(userx$) - 2))
                             GoTo sigue_aqui
                             
                             
                            ElseIf UCase(RTrim(userx$)) = UCase(RTrim(agente$)) And UCase(RTrim(csr$)) = UCase(RTrim(agente$)) Then
                             cant$ = Right(List1.List(Y), 9)
                             Total = Val(cantidad$) + Val(cant$)
                             List1.RemoveItem Y
                             
                            ElseIf UCase(RTrim(userx$)) <> UCase(RTrim(agente$)) And UCase(RTrim(csr$)) = UCase(RTrim(agente$)) Then
                             cant$ = Right(List1.List(Y), 9)
                             Total = Val(cantidad$) + Val(cant$)
                             List1.RemoveItem Y
                             
                            ElseIf UCase(RTrim(userx$)) <> UCase(RTrim(agente$)) And UCase(RTrim(csr$)) <> UCase(RTrim(agente$)) Then  ' se agrego 2/12
                             cant$ = Right(List1.List(Y), 9)
                             Total = Val(cantidad$) + Val(cant$)
                             List1.RemoveItem Y
                             
                            End If
 
r1:
                            r$ = UCase(Left(userx$, 2)) + LCase(Right(userx$, Len(userx$) - 2))
                            
                            If UCase(RTrim(userx$)) = UCase(RTrim(agente$)) And UCase(RTrim(csr$)) <> UCase(RTrim(agente$)) Then
                                r$ = UCase(Left(csr$, 2)) + LCase(Right(csr$, Len(csr$) - 2))
                            ElseIf UCase(RTrim(userx$)) = UCase(RTrim(agente$)) And UCase(RTrim(csr$)) = UCase(RTrim(agente$)) Then
                                r$ = UCase(Left(userx$, 2)) + LCase(Right(userx$, Len(userx$) - 2))
                            End If
                            
                        Else
                            cant$ = Right(List1.List(Y), 9)
                            Total = Val(cantidad$) + Val(cant$)
                            List1.RemoveItem Y
                        
                            r$ = UCase(Left(userx$, 2)) + LCase(Right(userx$, Len(userx$) - 2))
                        End If
                        
                        
sigue_aqui:
                        
                        
                        
                        List1.AddItem Format(r$, "!@@@@@@@@@@@@@@@@@@@@") + " " + Format(concepto$, "!@@@@@@@@@@@@@@@@@@@@") + Space(1) + Format(Format(Total, "#####0.00"), "@@@@@@@@@")
                        
                        encontrado = 1
                        Exit For
               End If
         
        Next Y
   
   
        If encontrado = 0 Then
        
               r$ = UCase(Left(userx$, 2)) + LCase(Right(userx$, Len(userx$) - 2))
               
               If RTrim(LTrim(UCase(concepto$))) = "INVOICE" Or RTrim(LTrim(UCase(concepto$))) = "LATE FEE" Then
 
 
 
                            ' --------------------------------------SE AGREGO ESTO 11/2/2023

                          '  If UCase(RTrim(userx$)) = agente$ And csr$ <> agente$ Then
                          '   cant$ = cantidad$
                          '   Total = Val(cant$)
                          '   userx$ = csr$
                          '   r$ = UCase(Left(userx$, 2)) + LCase(Right(userx$, Len(userx$) - 2))
                          '   GoTo brinca_aqui
                             
                             
                          '  ElseIf UCase(RTrim(userx$)) = agente$ And csr$ = agente$ Then
                          '   cant$ = Right(List1.List(Y), 9)
                          '   Total = Val(cantidad$) + Val(cant$)
                          '   List1.RemoveItem Y
                             
                          '  ElseIf UCase(RTrim(userx$)) <> agente$ And csr$ = userx$ Then
                          '   cant$ = Right(List1.List(Y), 9)
                          '   Total = Val(cantidad$) + Val(cant$)
                          '   List1.RemoveItem Y
                             
                          '  End If
 
                            ' ----------------------------------------------------------------
                            
                            
                            
 
                            'r$ = UCase(Left(userx$, 2)) + LCase(Right(userx$, Len(userx$) - 2))
                            
                            If UCase(RTrim(userx$)) = agente$ And csr$ <> agente$ Then
                                r$ = UCase(Left(csr$, 2)) + LCase(Right(csr$, Len(csr$) - 2))
                            ElseIf UCase(RTrim(userx$)) = agente$ And csr$ = agente$ Then
                                r$ = UCase(Left(userx$, 2)) + LCase(Right(userx$, Len(userx$) - 2))
                            End If
                            

                            
                                                    
                            
                            
                            
               Else
                            r$ = RTrim(UCase(Left(userx$, 2)) + LCase(Right(userx$, Len(userx$) - 2)))
                            
                            
                        ' -------------------------------------------- SE AGREGO ESTO 11/2/2023
                          '  cant$ = Right(List1.List(Y), 9)
                          '  Total = Val(cantidad$) + Val(cant$)
                          '  List1.RemoveItem Y
                         ' ------------------------------------------------------------------------------------
                        
               End If
               
brinca_aqui:
                        
                         
              List1.AddItem Format(r$, "!@@@@@@@@@@@@@@@@@@@@") + " " + Format(concepto$, "!@@@@@@@@@@@@@@@@@@@@") + Space(1) + Format(Format(cantidad$, "#####0.00"), "@@@@@@@@@")
                       
        End If
  
  End If
  
saltado:
Next t




' ================================================================================================================================================================
' ================================================================================================================================================================
' ================================================================================================================================================================
' ================================================================================================================================================================
' ================================================================================================================================================================
' ================================================================================================================================================================
' ==================================  AQUI EMPIEZA EL CHEQUEO DE CADA CANTIDAD DEL AGENTE ========================================================================



If manager_commercial$ <> "" Then
   manager_Oficinacommercial$ = manager_commercial$
End If




For t = 0 To List1.ListCount - 1
    csr$ = RTrim(UCase(Left$(List1.List(t), 20)))
    concepto$ = RTrim(UCase(Mid$(List1.List(t), 22, 20)))
    cant_del_concepto = Val(Right(List1.List(t), 9))
    
    
    GoSub check_manager
    
    existe = -1
    
    
    'If cant_del_concepto = 265 Then Stop
    
    
    If UCase(LTrim(csr$)) <> UCase(LTrim(manager_csr$)) Then  ' se agrego 12/12
       manager_csr$ = ""
    End If
    
    
    If concepto$ = "LATE FEE" Then
    
      agente_cobrador$ = ""
 
      sSelect = "select idemployeeUSR from [ReceiptsDTL] recdtl " & _
      "inner join  ReceiptsHDR  rechdr on rechdr.IDReceiptHDR =recdtl.IdReceiptHDR " & _
      "inner join InvoiceItemCatalog iitem on iitem.IdInvoiceItem=recdtl.IdInvoiceItem " & _
      "where rechdr.Active=1 and recdtl.[IdReceiptHDR] in (" + recibo$ + ") and iitem.InvoiceItemName='Late Fee'"
 
      Rs.Open sSelect, base, adOpenUnspecified
      agente_cobrador$ = Rs(0)
      Rs.Close
   
   
      If agente_cobrador$ <> "" Then
         sSelect = "select username from EmployeeInfo where IDEmployee='" + agente_cobrador$ + "'"
         Rs.Open sSelect, base, adOpenUnspecified
         cobrador$ = Rs(0)
         Rs.Close
      
         If UCase(cobrador$) <> UCase(agente$) Then
            existe = 3
            GoTo brinca
         End If
   
      End If
      
    End If
    
    
    
    'If cant_del_concepto = 19 Then Stop
    
      
    
    If UCase$(LTrim(concepto$)) = "BF COMMERCIAL" Then
      
      
        ' verifica si CSR2 es manager comercial
        
        
      If manager_csr2$ = "" And manager_Oficinacommercial$ = "" Then
      
           encontrado = 0
           For Y = 0 To contador_comercial
               
             If Val(Commercial$(Y, 1)) = cant_del_concepto Then   ' and Commercial$(Y, 0) = csr$
               manager_csr2$ = Commercial$(Y, 2)
               Exit For
             End If
           
           Next Y
           
      End If
            
            
      If manager_csr2$ <> csr$ And manager_csr2$ <> agente$ And manager_csr2$ <> "" Then
         existe = 3
         GoTo brinca
         
     ElseIf manager_csr2$ <> csr$ And manager_csr2$ <> agente$ And manager_csr2$ = "" And manager_Oficinacommercial$ = "" Then
         'existe = 9
         'GoTo brinca
         
      ElseIf manager_Oficinacommercial$ = agente$ Then
         existe = 1
         GoTo brinca
         
      End If
     
     
    End If
    
    
    
    
    If csr$ = agente$ Then
             c = Val(Right(List1.List(t), 9))
    Else
  
    ' verifica si monterrey el AGENTE
             existe = -1
             For Y = 0 To lista_managers.ListCount - 1
                       a$ = RTrim(UCase(Left$(lista_managers.List(Y), 20)))
                       b$ = LTrim(RTrim(Right(UCase(lista_managers.List(Y)), Len(lista_managers.List(Y)) - 20)))
                
                       If a$ = agente$ Then
                               If b$ = "MONTERREY" Then
                                          existe = 5
                                          Exit For
                               End If
                       End If
        
        
             Next Y
  
             If existe = 5 Then GoTo brinca
  
             ' verifica si monterrey el CSR
             existe = -1
             For Y = 0 To lista_managers.ListCount - 1
                       a$ = RTrim(UCase(Left$(lista_managers.List(Y), 20)))
                       b$ = LTrim(RTrim(Right(UCase(lista_managers.List(Y)), Len(lista_managers.List(Y)) - 20)))
                
                       If a$ = csr$ Then
                                If b$ = "MONTERREY" Then
                                             existe = 4
                                             Exit For
                                End If
                       End If
        
        
             Next Y
     
             If existe = 4 Then GoTo brinca
  
  
  ' verifica si commercial el CSR
     
     For Y = 0 To lista_managers.ListCount - 1
        a$ = RTrim(UCase(Left$(lista_managers.List(Y), 20)))
        b$ = LTrim(RTrim(Right(UCase(lista_managers.List(Y)), Len(lista_managers.List(Y)) - 20)))
        existe = -1
        
        If b$ = "MANAGER_COMMERCIAL" Then
           If manager_Oficinacommercial$ <> a$ Then
              b$ = "MANAGER"
           End If
        End If
        
        If a$ = csr$ Then
          If b$ = "COMMERCIAL" Or b$ = "MANAGER_COMMERCIAL" Or b$ = "AGENT_COMMERCIAL" Then
          
          
                                  GoSub check_manager
          
                                  csr$ = RTrim(UCase(Left$(List1.List(t), 20)))
                                  csr_despues$ = RTrim(UCase(Left$(List1.List(t + 1), 20)))
                                  csr_2despues$ = RTrim(UCase(Left$(List1.List(t + 2), 20)))
                                  csr_antes$ = RTrim(UCase(Left$(List1.List(t - 1), 20)))
                                  csr_mas_antes$ = RTrim(UCase(Left$(List1.List(t - 2), 20)))
                                  
                                  a$ = RTrim(UCase(Left$(lista_managers.List(Y), 20)))
                                  b$ = LTrim(RTrim(Right(UCase(lista_managers.List(Y)), Len(lista_managers.List(Y)) - 20)))
                                  
                                  cant_del_concepto_mas_antes = Val(Right(List1.List(t - 2), 9))
                                  cant_del_concepto_antes = Val(Right(List1.List(t - 1), 9))
                                  cant_del_concepto_despues = Val(Right(List1.List(t + 1), 9))
                                  cant_del_concepto_2despues = Val(Right(List1.List(t + 2), 9))
                                  cant_del_concepto = Val(Right(List1.List(t), 9))
                                  
          
                                  concepto_mas_antes$ = RTrim(UCase(Mid$(List1.List(t - 2), 22, 20)))
                                  concepto_antes$ = RTrim(UCase(Mid$(List1.List(t - 1), 22, 20)))
                                  concepto$ = RTrim(UCase(Mid$(List1.List(t), 22, 20)))
                                  concepto_despues$ = RTrim(UCase(Mid$(List1.List(t + 1), 22, 20)))
                                  concepto_2despues$ = RTrim(UCase(Mid$(List1.List(t + 2), 22, 20)))
                                  
                                  
                                 
                                  
                                  
                                  
                                  
                                  
                                  ' verifica si antes es el mismo CSR, sino pasa a DESPUES del CSR
                                  If csr$ <> csr_antes$ Then
                                      If concepto_despues$ = "BF COMMERCIAL" Then
                                           If cant_del_concepto <> cant_del_concepto_despues Then
                                                 If csr$ = csr_2despues$ Then
                                                      If Format(cant_del_concepto_despues, "###0.0") = Format(cant_del_concepto_2despues, "###0.0") Then
                                                          ' NO ES DE LA POLIZA COMERCIAL
                                                          If oficinaUSER1$ = oficinaCSR1$ Or oficinaUSER1$ = oficinaCSR2$ Or oficinaUSER2$ = oficinaCSR1$ Or oficinaUSER2$ = oficinaCSR2$ Then
                                                              If Left(b$, 7) = "MANAGER" Or b$ = "COMMERCIAL" Or b$ = "BF COMMERCIAL" Then
                                                                existe = 1    ' Todo
                                                                Exit For
                                                              Else
                                                                existe = 9    ' compartido
                                                                Exit For
                                                              End If
                                                            
                                                          Else
                                                              existe = 9    ' compartido
                                                              Exit For
                                                          End If
                                                          
                                                      End If
                                                 End If
                                           
                                           End If
                                          
                                      End If
                                                                                              
                                  
                                  End If
                                  
                                  
                                  
                                          
                                  
                                  
                                  
                                  si_es_poliza_commercial = 0
                                  
                                  If concepto_antes$ = "BF COMMERCIAL" And csr_antes$ = csr$ Then
                                       If Left(concepto_mas_antes$, 2) = "BF" And csr_mas_antes$ = csr$ And Format(cant_del_concepto_mas_antes, "0000.0") = Format(cant_del_concepto_antes, "0000.0") Then
                                          si_es_poliza_commercial = 0
                                          GoTo continue_right_here
                                       Else
                                          si_es_poliza_commercial = 1
                                       End If
                                    
                                  End If
                                  
                                  
                                  If concepto$ = "BF COMMERCIAL" And concepto_antes$ = "BF" And csr_antes$ = csr$ And b$ = "AGENT_COMMERCIAL" Then
                                    si_es_poliza_commercial = 1
                                  End If
                                  
                                  
                                  If concepto$ = "BF COMMERCIAL" And Left(UCase(concepto_despues$), 2) = "BF" And csr_despues$ = csr$ And b$ = "AGENT_COMMERCIAL" Then
                                    si_es_poliza_commercial = 1
                                  End If
                                  
                                  
                                  If concepto_despues$ = "BF COMMERCIAL" And csr_despues$ = csr$ Then
                                    si_es_poliza_commercial = 1
                                  End If
                                  
                                    
                                  If concepto_antes$ = "BF CALL CENTER" And csr_antes$ = csr$ Then
                                    si_es_poliza_commercial = 2
                                  End If
                                  
                                  If concepto_despues$ = "BF CALL CENTER" And csr_despues$ = csr$ Then
                                    si_es_poliza_commercial = 2
                                  End If
                                  
                                                               
                                  If concepto$ = "BF CALL CENTER" And concepto_antes$ = "BF" And csr_antes$ = csr$ Then
                                    si_es_poliza_commercial = 2
                                  End If
                                  
                                  
continue_right_here:
                                  
                                   ' verifica oficinas a que pertenecen
                                         For w = 1 To Val(ubicacion(0, 0))
                                              If UCase(RTrim(ubicacion(w, 0))) = UCase(agente$) Then
                                                      oficina_user$ = RTrim(ubicacion(w, 1))
                                                      oficina_user2$ = RTrim(ubicacion(w, 2))
                                              End If
                  
                                              If UCase(RTrim(ubicacion(w, 0))) = UCase(csr$) Then
                                                      oficina_csr$ = RTrim(ubicacion(w, 1))
                                                      oficina_csr2$ = RTrim(ubicacion(w, 2))
                                                                                                   
                                              End If
                                              
                                              ' csr_despues$
                                              
                                              
                                              If UCase(RTrim(ubicacion(w, 0))) = UCase(csr_despues$) Then
                                                      oficina_csr_next$ = RTrim(ubicacion(w, 1))
                                                      oficina_csr_next2$ = RTrim(ubicacion(w, 2))
                                                                                                   
                                              End If
                                              
                                              
                                              
                                         Next w
                                         
                                         
                                         
                                         If oficina_user2$ = "None" And oficina_csr2$ = "None" Then
                                                      oficina_user2$ = "Nada"
                                         End If
                  
                                         If oficina_user2$ = "" Then oficina_user2$ = " "
                                         
                                         If oficina_user2$ = "Nada" And oficina_csr2$ = "Nada" Then
                                                      oficina_user2$ = "None"
                                         End If
                                  
                                  
                                  
                                  
                                  
                                    If concepto_antes$ <> "BF COMMERCIAL" And concepto$ <> "BF COMMERCIAL" And concepto_despues$ <> "BF COMMERCIAL" And concepto_2despues$ <> "BF COMMERCIAL" Then
                                          ' NO ES DE LA POLIZA COMERCIAL
                                             If si_es_poliza_commercial = 0 Then
                                                   
                                                 If manager_agente$ = "" And manager_csr$ = "" Then  ' And manager_commercial$ = "" Then
                                                     
                                                    existe = 9   ' comparten
                                                    Exit For
                                                    
                                                 ElseIf manager_agente$ <> "" And manager_csr$ = "" Then
                                                                                                              
                                                    If (oficina_agente$ = oficina_csr$) Then
                                                    
                                                          existe = 3    ' cero para el agente
                                                          Exit For
                                                          
                                                    Else
                                                          existe = 9    ' comparten
                                                          Exit For
                                                    
                                                    End If
                                                    
                                                  ElseIf manager_agente$ = "" And manager_csr$ <> "" Then
                                                  
                                                    If (oficina_agente$ = oficina_csr$) Or (oficina_agente$ = oficina_csr2$) Then
                                                    
                                                          existe = 1    ' todo para el agente
                                                          Exit For
                                                          
                                                    Else
                                                          existe = 9    ' comparten
                                                          Exit For
                                                    
                                                    End If
                                                    
                                                    
                                                  ElseIf b$ = "AGENT_COMMERCIAL" And (a$ <> manager_agente$ And a$ <> manager_csr$) Then
                                                    
                                                    If (oficina_agente$ = oficina_csr$) Then
                                                    
                                                          existe = 3    ' cero para el agente
                                                          Exit For
                                                          
                                                    Else
                                                          existe = 9    ' comparten
                                                          Exit For
                                                    
                                                    End If
                                                    
                                                  Else
                                                  
                                                    
                                                        existe = 9  ' comparten
                                                        Exit For
                                                        
                                                    
                                                  End If
                                                    
                                             End If
                                  
                                    End If
                                         
                                         
                                         

                
                                  ' anula el status de agente comercial sobre manager gral comercial
                                  If manager_Oficinacommercial$ = agente$ Then
                                    si_es_poliza_commercial = 0
                                    existe = 3
                                    Exit For
                                  End If
                                  
                                  
                                  
                                  
                                      If concepto$ = "BF CALL CENTER" And si_es_poliza_commercial = 2 Then
                                             existe = 1
                                             Exit For
                                      End If
                                      
                                      
                                                                      
                                      If concepto$ = "BF COMMERCIAL" And si_es_poliza_commercial = 1 Then
                                             existe = 3
                                             Exit For
                                      End If
                                   
                                    
                                       
                                       
                                      If Left(concepto$, 2) = "BF" And concepto$ <> "BF COMMERCIAL" And concepto$ <> "BF CALL CENTER" And si_es_poliza_commercial = 1 Then
                                         
                                      
                                      
                                             existe = 1
                                             Exit For
                                      End If
                                  
                                  
                                      If concepto$ = "BF" And concepto_despues$ = "BF CALL CENTER" And si_es_poliza_commercial = 2 Then
                                             existe = 1
                                             Exit For
                                      End If
                                  
                                  
          
                                      
          
          
          
          
          
          
           
                                         
                                         
                                         
                                                                                  
                                         
                                         
                  
                                         If (oficina_user$ = oficina_csr$ Or oficina_user2$ = oficina_csr2$ Or oficina_user$ = oficina_csr2$ Or oficina_user2$ = oficina_csr$) Then
                                                    If concepto$ = "BF COMMERCIAL" Then
                                                      existe = 3
                                                      Exit For
                                                    Else
                                                    
                                                   
                                                    
                                                     ' es manager el agente?
                                                         es_manager = 0
                                                         For z = 0 To lista_managers.ListCount - 1
                                                             a$ = RTrim(UCase(Left$(lista_managers.List(z), 20)))
                                                             bb$ = LTrim(RTrim(Right(UCase(lista_managers.List(z)), Len(lista_managers.List(z)) - 20)))
                         
                                                             If a$ = csr$ Then
                                                                  If bb$ = "MANAGER" Or bb$ = "MANAGER_COMMERCIAL" Or bb$ = "COMMERCIAL" Then
                                                                          es_manager = 1
                                                                          Exit For
                                                                  End If
                                                             End If
                                                                
                                                             If a$ = agente$ Then
                                                                  If bb$ = "MANAGER" Or bb$ = "MANAGER_COMMERCIAL" Or bb$ = "COMMERCIAL" Then
                                                                          es_manager = 2
                                                                          Exit For
                                                                  End If
                                                             End If
                                                                
                                                         Next z
                                                         
                                                         If es_manager = 1 Then
                                                            existe = 1
                                                            
                                                         ElseIf es_manager = 2 Then
                                                            existe = 3
                                                         Else
                                                            existe = 9
                                                         End If
                                                         
                                                         
                                                         
                                                    
                                                    
                                                      'existe = 1    ' tenia existe=1   8/17/2022
                                                      Exit For
                                                    End If
                                         Else
                                                     existe = 9
                                                      Exit For
                  
                                         End If
                  
                                         
            
                                         existe = 7
                                         Exit For
                                         
                                         
                     
          ElseIf b$ = concepto$ And b$ = "BF COMMERCIAL" Then
                                  
                                            existe = 3
                                            Exit For
              
          ElseIf b$ = concepto$ And b$ = "BF" Then
                                            existe = 1
                                            Exit For
                
            
          End If
         
        End If
        
       ' If existe = -1 Then
       '   existe = 10
       '   Exit For
       ' End If
        
        
     Next Y
     
     
salto_temp:
     
     If existe = 1 Or existe = 9 Or existe = 11 Or existe = 3 Then GoTo brinca
  
  ' verifica si es commercial el AGENTE
     'existe = 0
             For Y = 0 To lista_managers.ListCount - 1
                           a$ = RTrim(UCase(Left$(lista_managers.List(Y), 20)))
                           b$ = LTrim(RTrim(Right(UCase(lista_managers.List(Y)), Len(lista_managers.List(Y)) - 20)))
                
                           If a$ = agente$ Then   ' csmbie csr$ por agente$
                                  If b$ = "COMMERCIAL" Or b$ = "MANAGER_COMMERCIAL" Or b$ = "AGENT_COMMERCIAL" Or b$ = "BF COMMERCIAL" Then
                                  
                                 
                                  
                                  csr$ = RTrim(UCase(Left$(List1.List(t), 20)))
                                  csr_despues$ = RTrim(UCase(Left$(List1.List(t + 1), 20)))
                                  csr_2despues$ = RTrim(UCase(Left$(List1.List(t + 2), 20)))
                                  csr_antes$ = RTrim(UCase(Left$(List1.List(t - 1), 20)))
                                  
                                  
                                  
                                  cant_del_concepto_antes = Val(Right(List1.List(t - 1), 9))
                                  cant_del_concepto_despues = Val(Right(List1.List(t + 1), 9))
                                  cant_del_concepto_2despues = Val(Right(List1.List(t + 2), 9))
                                  cant_del_concepto = Val(Right(List1.List(t), 9))
                                  
          
                                  concepto_antes$ = RTrim(UCase(Mid$(List1.List(t - 1), 22, 20)))
                                  concepto$ = RTrim(UCase(Mid$(List1.List(t), 22, 20)))
                                  concepto_despues$ = RTrim(UCase(Mid$(List1.List(t + 1), 22, 20)))
                                  concepto_2despues$ = RTrim(UCase(Mid$(List1.List(t + 2), 22, 20)))
                                  
                                  
                                  
                                  ' verifica si antes es el mismo CSR, sino pasa a DESPUES del CSR
                                  If csr$ <> csr_antes$ Then
                                      If concepto_despues$ = "BF COMMERCIAL" Then
                                           If Format(cant_del_concepto, "###0.0") <> Format(cant_del_concepto_despues, "###0.0") Then
                                                 If csr$ = csr_2despues$ Then
                                                      If Format(cant_del_concepto_despues, "###0.0") = Format(cant_del_concepto_2despues, "###0.0") Then
                                                          ' NO ES DE LA POLIZA COMERCIAL, verifica si es manager la otra persona
                                                          
                                                          If oficinaUSER1$ = oficinaCSR1$ Or oficinaUSER1$ = oficinaCSR2$ Or oficinaUSER2$ = oficinaCSR1$ Or oficinaUSER2$ = oficinaCSR2$ Then
                                                              If Left(b$, 7) = "MANAGER" Or b$ = "COMMERCIAL" Or b$ = "BF COMMERCIAL" Then
                                                                existe = 3    ' nada porque es manager del agente
                                                                Exit For
                                                              Else
                                                                existe = 9    ' compartido
                                                                Exit For
                                                              End If
                                                            
                                                          Else
                                                              existe = 9    ' compartido
                                                              Exit For
                                                          End If
                                                      End If
                                                 End If
                                           
                                           End If
                                          
                                      End If
                                                                                              
                                  
                                  End If
                                  
                                  
                                  
                                  
                                  si_es_poliza_commercial = 0
                                  
                                  
                                 
                                  
                                  
                                  If List1.List(t + 1) <> "" Then
                                    cant_despues_concepto = Val(RTrim(UCase(Right$(List1.List(t + 1), 10))))
                                  End If
                                  
                                  
                                  
                                  If concepto_antes$ = "BF COMMERCIAL" And csr_antes$ = csr$ Then
                                    si_es_poliza_commercial = 1
                                  End If
                                  
                                  If concepto_despues$ = "BF COMMERCIAL" And csr_despues$ = csr$ Then
                                    si_es_poliza_commercial = 1
                                  End If
                                  
                                  
                                  If concepto$ = "BF COMMERCIAL" And csr_despues$ = "" Then
                                    si_es_poliza_commercial = 1
                                  End If
                                  
                                  If concepto$ = "BF COMMERCIAL" Then
                                    si_es_poliza_commercial = 1
                                  End If
                                  
                                  If (concepto_despues$ = "BF CALL CENTER" And csr_despues$ <> csr$) And (concepto$ = "BF CALL CENTER" And csr_antes$ = csr$) Then
                                    si_es_poliza_commercial = 3
                                    GoTo brinca_caso
                                  End If
                                  
                                  
                                  
                                  
                                  
                                  
                                  If concepto_antes$ = "BF CALL CENTER" And csr_antes$ = csr$ Then
                                    si_es_poliza_commercial = 2
                                  End If
                                  
                                  If concepto_despues$ = "BF CALL CENTER" And csr_despues$ = csr$ Then
                                    si_es_poliza_commercial = 2
                                  End If
                                  
                                  If concepto$ = "BF CALL CENTER" And csr_despues$ = csr$ Then
                                    si_es_poliza_commercial = 2
                                  End If
                                    
                                  If concepto$ = "BF CALL CENTER" And csr_antes$ = csr$ Then
                                    si_es_poliza_commercial = 2
                                  End If
                                  
                                  
                                                                   
                                  
brinca_caso:
                                     
                                     
                                      If concepto$ = "BF" And si_es_poliza_commercial = 2 And concepto_despues$ = "BF CALL CENTER" And csr_despues$ = csr$ Then
                                             existe = 3  '
                                             Exit For
                                      End If
                                     
                                     
                                      If concepto$ = "BF CALL CENTER" And si_es_poliza_commercial = 2 And concepto_antes$ = "BF" And csr_antes$ = csr$ Then
                                             existe = 1  '
                                             Exit For
                                      End If
                                      
                                      
                                      
                                     
                                   
                                      If concepto$ = "BF COMMERCIAL" And si_es_poliza_commercial = 1 Then
                                             existe = 1
                                             Exit For
                                      End If
                                   
                                                                                                           
                                      If concepto$ = "BF CALL CENTER" And si_es_poliza_commercial = 2 Then
                                             existe = 3
                                             Exit For
                                      End If
                                      
                                      
                                      If concepto$ = "BF CALL CENTER" And si_es_poliza_commercial = 3 Then
                                             existe = 9
                                             Exit For
                                      End If
                                      
                                      
                                      
                                      
                                      
                                       
                                       
                                     ' If concepto$ = "BF ENDO FEE" And si_es_poliza_commercial = 1 Then
                                     '        existe = 1
                                     '        Exit For
                                     ' End If
                                      
                                      
                                       
                                      If Left(concepto$, 2) = "BF" And concepto$ <> "BF COMMERCIAL" And concepto$ <> "BF CALL CENTER" And si_es_poliza_commercial = 1 Then
                                             
                                             
                                             
                                             
                                             
                                             
                                             If b$ = "AGENT_COMMERCIAL" Or b$ = "COMMERCIAL" Then
                                                
                                                If (concepto$ = "BF" Or concepto$ = "BF ENDO FEE" Or concepto$ = "BF PAYMENT FEE") And (concepto_despues$ = "BF COMMERCIAL" Or concepto_antes$ = "BF COMMERCIAL") Then
                                                
                                                      ' If cant_del_concepto <> cant_despues_concepto Then
                                                      
                                                          ' verifica si la cantidad del BF es parte de un BF COMMERCIAL
                                                          ' ***************************************************************************************
                                                                     
                                                                     cont_veces = 0
                                                                     
Obten_BF:
                                                                     Set Rs = New ADODB.Recordset

                                                                     
                                                                     If cont_veces = 0 Then
                                                                       r$ = "and iitem.InvoiceItemName='BF COMMERCIAL' "
                                                                     Else
                                                                       r$ = ""
                                                                     End If
                                                                     
                                                                     sSelect = "SELECT " & _
                                                                     "recdtl.[IdReceiptHDR] as [Receipt #] " & _
                                                                     ",rechdr.Date " & _
                                                                     ",iitem.InvoiceItemName as [Invoice Item] " & _
                                                                     ",emp.Username as [User] " & _
                                                                     ",csr.Username as [CSR] " & _
                                                                     ",recdtl.Amount " & _
                                                                     ",ofc.Office " & _
                                                                     ",rechdr.IdOffice, rechdr.AmountPaid, rechdr.IdEmployeeCSR2, rechdr.BalanceDue, rechdr.IdCustomer  " & _
                                                                     "FROM  [ReceiptsDTL] recdtl " & _
                                                                     "inner join  ReceiptsHDR  rechdr on recdtl.IdReceiptHDR=rechdr.IDReceiptHDR " & _
                                                                     "inner join PoliciesHDR polhdr on polhdr.IdPoliciesHDR=rechdr.IdPoliciesHDR " & _
                                                                     "inner join EmployeeInfo emp on emp.IDEmployee=rechdr.IdEmployeeUSR " & _
                                                                     "inner join EmployeeInfo csr on csr.IDEmployee=rechdr.IdEmployeeCSR1 " & _
                                                                     "inner join OfficesCatalog ofc on ofc.IdOffice=rechdr.IdOffice " & _
                                                                     "inner join OfficesCatalog ofcpol on ofcpol.IdOffice=rechdr.IdOffice " & _
                                                                     "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
                                                                     "inner join Customers cus on cus.IdCustomer=polhdr.IdCustomer " & _
                                                                     "inner join InvoiceItemCatalog iitem on iitem.IdInvoiceItem=recdtl.IdInvoiceItem " & _
                                                                     "inner join BankAccountCatalog bank on bank.IdBankAccount=iitem.IdBankAccount " & _
                                                                     "inner join CompanysAGICatalog ja on ja.IdCompanyAGI=polhdr.IdCompanyAGI " & _
                                                                     "where iitem.IdInvoiceItem in (2,35,3,11,15,22,23,33,9,30,33,17,13,31,32, 36, 37,39,42,43) " & _
                                                                     "and cast(rechdr.Date as Date) >= '" + txtdatefrom.Text + "' AND cast( rechdr.DATE as Date) <= '" + txtdateto.Text + "' " & _
                                                                     "and rechdr.Active=1 " + r$ + " and (csr.username='" + csr$ + "' or emp.Username='" + csr$ + "') " & _
                                                                     "order by [Receipt #], rechdr.IdCustomer"
        
    
                                                                     ' ---------------------------------------------------------------------------
    
                                                                     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
                                                                     Rs.Open sSelect, base, adOpenStatic, adLockOptimistic
    
                                                                     Rs.MoveLast

                                                                     Rs.MoveFirst
                                                                     ' Assuming that rs is your ADO recordset
                                                                     grid.Rows = Rs.RecordCount + 1

                                                                     rsVar = Rs.GetString(adClipString, Rs.RecordCount)

                                                                     grid.cols = Rs.Fields.Count + 1
    
                                                                     grid.TextMatrix(0, 0) = ""
                                                                     ' Set column names in the grid
                                                                     For i = 0 To Rs.Fields.Count - 1
                                                                         grid.TextMatrix(0, i + 1) = Rs.Fields(i).name
                                                                     Next

                                                                     grid.row = 1
                                                                     grid.col = 1

                                                                     ' Set range of cells in the grid
                                                                     grid.RowSel = grid.Rows - 1
                                                                     grid.ColSel = grid.cols - 1
                                                                     grid.clip = rsVar

                                                                     ' Reset the grid's selected range of cells
                                                                     grid.RowSel = grid.row
                                                                     grid.ColSel = grid.col

                                                                     Rs.Close

                                                                     Set Rs = Nothing
                                                                       
                                                                                                                                                                                                             
                                                                     encontrado1 = 0
                                                                     For w = 1 To grid.Rows - 1
                                                                         grid.row = w
                                                                         grid.col = 6
                                                                         cantidad_comercial = Val(grid.Text)
                                                                                
                                                                         If Format(cantidad_comercial, "###0.0") = Format(cant_del_concepto, "###0.0") Or Format((cantidad_comercial + 0.01), "###0.0") = Format(cant_del_concepto, "###0.0") Then
                                                                            encontrado1 = 1
                                                                            Exit For
                                                                         End If
                                                                     Next w
                                                                     
                                                                     If encontrado1 = 1 Then
                                                                        existe = 3
                                                          
                                                                     End If
                                                      
                                                           
                                                           'c = (cant_del_concepto - cant_despues_concepto) / 2
                                                           'GoTo suma_la_cantidad
                                                           
                                                           
                                                      'Else
                                                      '     existe = 3
                                                      'End If
                                                'ElseIf concepto$ = "BF ENDO FEE" And (concepto_antes$ = "BF COMMERCIAL" Or concepto_despues$ = "BF COMMERCIAL") Then
                                                '      existe = 3   '  cero
                                                      
                                                'ElseIf concepto$ = "BF PAYMENT FEE" And (concepto_antes$ = "BF COMMERCIAL" Or concepto_despues$ = "BF COMMERCIAL") Then
                                                '      existe = 3  ' cero
                                                      
                                                Else
                                                      existe = 9   '  mitad     tenia existe=9  9/14/2022
                                                End If
                                                
                                                
                                                
                                               
                                             Else
                                                      existe = 3    '   cero    tenia existe=1   8/17/2022
                                                      
                                             End If
                                             
                                             
                                             'existe = 3   ' existe estaba como 3   8/17/2022
                                             Exit For
                                      End If
                                       
                                                                    
                                   
                                   
                                   
                                            ' verifica oficinas a que pertenecen
                                         For w = 1 To Val(ubicacion(0, 0))
                                              If UCase(RTrim(ubicacion(w, 0))) = UCase(agente$) Then
                                                      oficina_user$ = RTrim(ubicacion(w, 1))
                                                      oficina_user2$ = RTrim(ubicacion(w, 2))
                                              End If
                  
                                              If UCase(RTrim(ubicacion(w, 0))) = UCase(csr$) Then
                                                      oficina_csr$ = RTrim(ubicacion(w, 1))
                                                      oficina_csr2$ = RTrim(ubicacion(w, 2))
                                              End If
                                         Next w
                                         
                                         
                                         
                                         If oficina_user2$ = "None" And oficina_csr2$ = "None" Then
                                                      oficina_user2$ = "Nada"
                                         End If
                  
                                         If oficina_user2$ = "" Then oficina_user2$ = " "
                                         
                                         
                                         If oficina_user2$ = "Nada" And oficina_csr2$ = "Nada" Then
                                                     oficina_user2$ = "None"
                                         End If
                  
                  
                                         If (oficina_user$ = oficina_csr$ Or oficina_user2$ = oficina_csr2$ Or oficina_user$ = oficina_csr2$ Or oficina_user2$ = oficina_csr$) Then
                                                    
                                                    If b$ = "AGENT_COMMERCIAL" Then
                                                         
                                                         ' es manager del agente
                                                         es_manager = 0
                                                         For z = 0 To lista_managers.ListCount - 1
                                                             a$ = RTrim(UCase(Left$(lista_managers.List(z), 20)))
                                                             bb$ = LTrim(RTrim(Right(UCase(lista_managers.List(z)), Len(lista_managers.List(z)) - 20)))
                         
                                                             If a$ = csr$ Then
                                                                  If bb$ = "MANAGER" Or bb$ = "MANAGER_COMMERCIAL" Then
                                                                          es_manager = 1
                                                                          Exit For
                                                                  End If
                                                             End If
                                                                
                                                         Next z
                                                         
                                                         If es_manager = 1 Then
                                                            existe = 1
                                                         Else
                                                            existe = 9
                                                         End If
                                                      
                                                    Else
                                                      existe = 3    '   8/17/2022
                                                      
                                                    End If
                                                      
                                                      ' estaba existe=3
                                                      Exit For
                                         Else
                                                      existe = 9
                                                      Exit For
                  
                                         End If
                             
                                                           
                                        
                                         
                                         
                                            existe = 6
                                            Exit For
                                            
                                            
                                            
                                 ' ElseIf b$ = concepto$ And b$ <> "BF COMMERCIAL" Then
                                  '          existe = 8
                                  '          Exit For
                                            
                                   ElseIf b$ = concepto$ And b$ = "BF COMMERCIAL" Then
                                      ' If manager_Oficinacommercial$ = a$ Then
                                            existe = 1
                                      ' Else
                                      '      existe = 11
                                      ' End If
                                            Exit For
              
              
                                  ElseIf b$ = concepto$ And b$ = "BF" Then
                                            existe = 3
                                            Exit For
                                  End If
         
                           End If
        
        
             Next Y
  
      If existe = 3 Or existe = 8 Or existe = 9 Or existe = 11 Or existe = 1 Then GoTo brinca
  
  
  
  
  
     
  
             ' verifica si es manager
             existe = -1
             For Y = 0 To lista_managers.ListCount - 1
                         a$ = RTrim(UCase(Left$(lista_managers.List(Y), 20)))
                         b$ = LTrim(RTrim(Right(UCase(lista_managers.List(Y)), Len(lista_managers.List(Y)) - 20)))
                         
                         If a$ = csr$ Then ' Or agente$ = a$ Then
                                    If b$ = "MANAGER" Or b$ = "MANAGER_COMMERCIAL" Then
                                                                                                       
                                         ' verifica oficinas a que pertenecen
                                         For w = 1 To Val(ubicacion(0, 0))
                                              If UCase(RTrim(ubicacion(w, 0))) = UCase(agente$) Then
                                                      oficina_user$ = RTrim(ubicacion(w, 1))
                                                      oficina_user2$ = RTrim(ubicacion(w, 2))
                                              End If
                  
                                              If UCase(RTrim(ubicacion(w, 0))) = UCase(csr$) Then
                                                      oficina_csr$ = RTrim(ubicacion(w, 1))
                                                      oficina_csr2$ = RTrim(ubicacion(w, 2))
                                              End If
                                         Next w
                  
                  
                                         If oficina_user2$ = "None" And oficina_csr2$ = "None" Then
                                                      oficina_user2$ = "Nada"
                                         End If
                  
                                         If oficina_user2$ = "" Then oficina_user2$ = " "
                                         
                                         If oficina_user2$ = "Nada" And oficina_csr2$ = "Nada" Then
                                                      oficina_user2$ = "None"
                                         End If
                  
                                         If (oficina_user$ = oficina_csr$ Or oficina_user2$ = oficina_csr2$ Or oficina_user$ = oficina_csr2$ Or oficina_user2$ = oficina_csr$) Then
                                                    
                                                      existe = 1
                                                      Exit For
                                         Else
                                                      existe = 8
                                                      Exit For
                  
                                         End If
                  
                                    
                                               
                    
                                    ElseIf b$ = concepto$ Then
                                                existe = 2
                                                Exit For
                                    End If
         
                         End If
        
        
        
        
                         If a$ = csr2$ Then ' Or agente$ = a$ Then
                                    If b$ = "MANAGER_COMMERCIAL" Then
                                                                                                       
                                         ' verifica oficinas a que pertenecen
                                         For w = 1 To Val(ubicacion(0, 0))
                                              If UCase(RTrim(ubicacion(w, 0))) = UCase(agente$) Then
                                                      oficina_user$ = RTrim(ubicacion(w, 1))
                                                      oficina_user2$ = RTrim(ubicacion(w, 2))
                                              End If
                  
                                              If UCase(RTrim(ubicacion(w, 0))) = UCase(csr$) Then
                                                      oficina_csr$ = RTrim(ubicacion(w, 1))
                                                      oficina_csr2$ = RTrim(ubicacion(w, 2))
                                              End If
                                         Next w
                  
                  
                                         If oficina_user2$ = "None" And oficina_csr2$ = "None" Then
                                                      oficina_user2$ = "Nada"
                                         End If
                  
                                         If oficina_user2$ = "" Then oficina_user2$ = " "
                                         
                                         If oficina_user2$ = "Nada" And oficina_csr2$ = "Nada" Then
                                                      oficina_user2$ = "None"
                                         End If
                                         
                                         
                                         
                                          
                  
                                         If (oficina_user$ = oficina_csr$ Or oficina_user2$ = oficina_csr2$ Or oficina_user$ = oficina_csr2$ Or oficina_user2$ = oficina_csr$) Then
                                                    
                                                      existe = 1
                                                      Exit For
                                         Else
                                                      existe = 8
                                                      Exit For
                  
                                         End If
                  
                                    
                                               
                    
                                    ElseIf b$ = concepto$ Then
                                                existe = 2
                                                Exit For
                                    End If
         
                         End If
        
        
        
                         If a$ = agente$ And csr$ <> agente$ Then
          
                                    If b$ = "MANAGER" Or b$ = "MANAGER_COMMERCIAL" Then
                                    
                                    ' verifica oficinas a que pertenecen
                                          For w = 1 To Val(ubicacion(0, 0))
                                                If UCase(RTrim(ubicacion(w, 0))) = UCase(agente$) Then
                                                            oficina_user$ = RTrim(ubicacion(w, 1))
                                                            oficina_user2$ = RTrim(ubicacion(w, 2))
                                                End If
                  
                                                If UCase(RTrim(ubicacion(w, 0))) = UCase(csr$) Then
                                                            oficina_csr$ = RTrim(ubicacion(w, 1))
                                                            oficina_csr2$ = RTrim(ubicacion(w, 2))
                                                End If
                                          Next w
                  
                                          If oficina_user2$ = "None" And oficina_csr2$ = "None" Then
                                              oficina_user2$ = "Nada"
                                          End If
                                          
                                          If oficina_user2$ = "" Then oficina_user2$ = " "
                                          
                                          
                                          If oficina_user2$ = "Nada" And oficina_csr2$ = "Nada" Then
                                              oficina_user2$ = "None"
                                          End If
                                          
                  
                  
                                          If oficina_user$ = oficina_csr$ Or (oficina_user2$ = oficina_csr2$ And oficina_user2$ <> "") Or oficina_user$ = oficina_csr2$ Or oficina_user2$ = oficina_csr$ Then
                                                            existe = 3     'c=0
                                                            Exit For
                                          Else
                                                            existe = 8     'c=c/2
                                                            Exit For
                                          End If
                  
                  
                                                
                    
                                    ElseIf b$ = concepto$ Then
                                                existe = 2   'c=c
                                                Exit For
                                    End If
        
        
        
           
                         End If
        
                         If a$ = agente$ Then
        
                         End If
        
        
             Next Y
     
'            If existe = 1 Or existe = 3 Then GoTo brinca
  
     
  
     
brinca:

     If BF_CALL = 1 Then
         For k = 0 To List2.ListCount - 1
              concepto$ = Left(List2.List(k), 20)
              recibo$ = Mid$(List2.List(k), 21, 6)
              usuario$ = Mid$(List2.List(k), 28, 20)
              cantidad$ = Right(List2.List(k), 9)
              
              If Format(Val(cant_guardada_CALL_CENTER$), "###0.0") = Format(Val(cantidad$), "###0.0") And LTrim(RTrim(UCase(concepto$))) <> "BF CALL CENTER" Then
                existe = 1   ' asigna todo
                If par_encontrado = 1 Then
                   par_encontrado = par_encontrado + 1
                End If
                Exit For
              End If
              
           Next k

           BF_CALL = 0
           
     End If

  
     If UCase(concepto$) = "BF CALL CENTER" Then
       If (UCase(oficina_user$) = "JA - PHONE SALES") Then
          
       Else
           existe = 5  ' asigna CERO
           BF_CALL = 1
           par_encontrado = par_encontrado + 1
           cant_guardada_CALL_CENTER$ = LTrim(Str(cant_del_concepto))  'cantidad$
           USUARIO_DE_PS$ = csr$
       End If
       
       
     End If
     
     
     If par_encontrado = 2 Then
       par_encontrado = 0
       
     End If
     
     
     ' detecta si el invoice con JENNIFERR esta en la lista de facturas ya pasadas de tiempo
     
     
     
     
  
     If agente$ = "JENNIFERR" And (UCase(concepto$) = "INVOICE" Or UCase(concepto$) = "LATE FEE") Then   ' And UCase(cobrador$) <> "JENNIFERR" Then
       
       If UCase(concepto$) = "INVOICE" Then
       
        hayado = 0
        For q = 0 To lista_invoices30.ListCount - 1
            
            cant_invoice = Val(Format(Mid$(lista_invoices30.List(q), 8, 9), "00000.00"))
            
            
            If cant_del_concepto = cant_invoice Then
            
             n$ = LTrim(RTrim(lista_invoices30.List(q)))
             For Y = Len(n$) To 1 Step -1
                If Mid$(n$, Y, 1) <> Space(1) Then
                   conta = conta + 1
                   csr$ = Right$(n$, conta)
                Else
                   Exit For
                End If
             Next Y
 
             agente1$ = csr$
  
             r$ = RTrim(Left(n$, Len(n$) - (conta)))
             conta = 0
             For Y = Len(r$) To 1 Step -1
               If Mid$(r$, Y, 1) <> Space(1) Then
                     conta = conta + 1
                     csr$ = Right$(r$, conta)
               Else
                     Exit For
               End If
             Next Y
             csr1$ = csr$
            
            
             If UCase(csr$) = UCase(csr1$) Or UCase(csr$) = UCase(agente1$) Then
               existe = 1
               hayado = 1
               Exit For
             End If
            
            End If
            
            
        Next q
       
        If hayado = 1 Then
           GoTo revisa
        End If
       
        existe = 9
        GoTo revisa
        
       End If
       
        existe = 1
        
     ElseIf csr$ = "JENNIFERR" And (UCase(concepto$) = "INVOICE" Or UCase(concepto$) = "LATE FEE") Then
        
        If UCase(cobrador$) = "JENNIFERR" And UCase(concepto$) = "LATE FEE" Then
           existe = 3
        Else
           existe = 9
        End If
                        
        
     End If
     
     
  
    ' If Alta_Comercial = 1 Then
    '    existe = 5  ' asigna CERO
    ' End If
  
     
     'If Es_Monterrey = 1 Then
     '  existe = 1
    ' End If
  
revisa:
  
  
  
     If existe = 1 Then
       c = (Val(Right(List1.List(t), 9)))
     ElseIf existe = 2 Then
       c = (Val(Right(List1.List(t), 9))) '/ 2
     ElseIf existe = 3 Then
       c = 0
     ElseIf existe = 4 Then
       c = (Val(Right(List1.List(t), 9)))
     ElseIf existe = 5 Then
       c = 0
     ElseIf existe = 6 Then
       c = (Val(Right(List1.List(t), 9))) / 2
       ' c = 0
     ElseIf existe = 8 Then
        c = (Val(Right(List1.List(t), 9))) / 2
       
     ElseIf existe = 7 Then
       c = (Val(Right(List1.List(t), 9))) / 2   ' tenia c = 0
    ElseIf existe = 9 Then
       c = (Val(Right(List1.List(t), 9))) / 2
       
       
     ElseIf existe = 10 Then
       c = cant_del_concepto / 2
       'c = 0
       
     ElseIf existe = 11 Then   ' BF COMMERCIAL
       c = 0
       
     Else
       c = (Val(Right(List1.List(t), 9))) / 2
     End If
  End If
  
suma_la_cantidad:
    
  gtotal = gtotal + c
  Alta_Comercial = 0

Next t



If par_encontrado = 1 Then
   gtotal = gtotal + (Val(cant_guardada_CALL_CENTER$) / 2)

End If
   
   
   
If UCase(lblagent.Caption) = "JENNIFERR" Then
    GoTo brinca_por_recaudacion
End If


' descuenta el invoice de mas de 30 dias si existe

For t = 0 To lista_invoices30.ListCount - 1
  csr$ = ""
  conta = 0
  
  
  
  
  n$ = LTrim(RTrim(lista_invoices30.List(t)))
  For Y = Len(n$) To 1 Step -1
    If Mid$(n$, Y, 1) <> Space(1) Then
         conta = conta + 1
         csr$ = Right$(n$, conta)
    Else
         Exit For
    End If
  Next Y
 
  agente1$ = csr$
  
  
  
  
  r$ = RTrim(Left(n$, Len(n$) - (conta)))
  
  conta = 0
  For Y = Len(r$) To 1 Step -1
    If Mid$(r$, Y, 1) <> Space(1) Then
         conta = conta + 1
         csr$ = Right$(r$, conta)
    Else
         Exit For
    End If
  Next Y
  csr1$ = csr$
  
  
  ' verifica si esta asignado
  existe = 0
  For k = 0 To List1.ListCount - 1
    '  List1.AddItem Format(r$, "!@@@@@@@@@@@@@@@@@@@@") + " " + Format(concepto$, "!@@@@@@@@@@@@@@@@@@@@") + Space(1) + cantidad$
    nom_fact$ = UCase(RTrim(Left(List1.List(k), 20)))
    If UCase(nom_fact$) = UCase(agente1$) Or UCase(nom_fact$) = UCase(csr1$) Then
       existe = 1
       Exit For
    Else
      
    
    End If
    
    
  Next k
  
  
  If existe = 0 Then
     GoTo no_hagas
  End If
  
  
  
  ' asigna la oficina a user y a csr y a agente
   
   For Y = 1 To Val(ubicacion(0, 0))
      If UCase(RTrim(ubicacion(Y, 0))) = UCase(csr1$) Then
         oficina_csr$ = ubicacion(Y, 1)
         Exit For
      End If
   Next Y
   
   
   For Y = 1 To Val(ubicacion(0, 0))
      If UCase(RTrim(ubicacion(Y, 0))) = UCase(user1$) Then
         oficina_user$ = ubicacion(Y, 1)
         Exit For
      End If
   Next Y
      
      
   For Y = 1 To Val(ubicacion(0, 0))
      If UCase(RTrim(ubicacion(Y, 0))) = UCase(agente1$) Then
         oficina_agente$ = ubicacion(Y, 1)
         Exit For
      End If
   Next Y
    
    
    
    
    
    
  
  
  
  
  r$ = RTrim(Left(r$, Len(r$) - conta - 1))
  
  
  For w = 0 To List2.ListCount - 1
    '          linea$ = Format(concepto$, "!@@@@@@@@@@@@@@@@@@@@") + Format(recibo$, "000000") + " " + Format(r$, "!@@@@@@@@@@@@@@@@@@@@") + " " + Format(cantidad$, "000000.00")
        recibox$ = Mid$(List2.List(w), 21, 6)
        If recibox$ = Left(r$, 6) Then
          cant = Val(Right(List2.List(w), 9))
          Exit For
        End If
        
  Next w
  
  pos = InStr(1, r$, "$")
  'cant = Val(Right(r$, Len(r$) - (pos)))
  
  existe = 0
  
  If RTrim(UCase(lblagent.Caption)) = RTrim(UCase(agente1$)) And RTrim(UCase(lblagent.Caption)) = RTrim(UCase(csr1$)) And Right(oficina_csr$, 9) <> "MONTERREY" Then
    
      
        gtotal = gtotal - cant
        existe = 1
      
  End If
  
  
  If RTrim(UCase(lblagent.Caption)) <> RTrim(UCase(csr1$)) And RTrim(UCase(lblagent.Caption)) = RTrim(UCase(agente1$)) And Right(oficina_csr$, 9) <> "MONTERREY" Then
      
         gtotal = gtotal - (cant / 2)
         existe = 1
    
  End If
  
     
   
  
   If RTrim(UCase(lblagent.Caption)) <> RTrim(UCase(agente1$)) And RTrim(UCase(lblagent.Caption)) = RTrim(UCase(csr1$)) And Right(oficina_csr$, 9) <> "MONTERREY" Then
       
    
        gtotal = gtotal - (cant / 2)
        If gtotal < 0 Then gtotal = 0
        existe = 1
   
       
      
  
  End If
  
  
no_hagas:
  
  
Next t



brinca_por_recaudacion:


X = Redondear(gtotal, 2)  ' redondea a 2 decimales
If X < 0 Then X = 0
lbltotal_invoices.Caption = Format(X, "$###,##0.00")




' realiza total de cada agente
Erase tabla
    csr$ = ""
    For t = 0 To List1.ListCount - 1
      nombre$ = Left(List1.List(t), 20)
      existe = 0
      For Y = 0 To 19
         If tabla(Y, 0) = "" Then
            Exit For
         End If
      
         If tabla(Y, 0) = nombre$ Then
            tabla(Y, 1) = tabla(Y, 1) + Val(Right(List1.List(t), 9))
            existe = 1
            Exit For
         End If
         
      Next Y
      If existe = 0 Then
         tabla(Y, 0) = nombre$
         tabla(Y, 1) = tabla(Y, 1) + Val(Right(List1.List(t), 9))
      End If
        
        
        
       
    Next t
    
    

carga_treeview

  Exit Sub
  
  
  
check_manager:


     ' verifica si CSR es manager
     existe = 0
     Es_Monterrey = 0
     
     For z = 0 To lista_managers.ListCount - 1
       a$ = RTrim(UCase(Left$(lista_managers.List(z), 20)))
       b$ = LTrim(RTrim(Right(UCase(lista_managers.List(z)), Len(lista_managers.List(z)) - 20)))
       
       
       
       If a$ = UCase(userx$) Then '
          If b$ = "MANAGER" Or b$ = "MANAGER_COMMERCIAL" Or b$ = "COMMERCIAL" Then
             manager_csr$ = a$
             existe = 1
            ' Exit For
          End If
        End If
        
        
        If a$ = csr$ Then '
          If b$ = "MANAGER" Or b$ = "MANAGER_COMMERCIAL" Or b$ = "COMMERCIAL" Then
             manager_csr$ = a$
             existe = 1
             'Exit For
          End If
          
        End If
        
        
        If a$ = user$ Then
            If b$ = "MANAGER" Or b$ = "MANAGER_COMMERCIAL" Or b$ = "COMMERCIAL" Then
             manager_user$ = a$
             existe = 1
             'Exit For
            End If
             
        End If
          
          
          
        If a$ = agente$ Then
          If b$ = "MANAGER" Or b$ = "MANAGER_COMMERCIAL" Or b$ = "COMMERCIAL" Then
             manager_agente$ = a$
             existe = 1
             'Exit For
          End If
             
        End If
        
        
        If a$ = agente$ Or a$ = csr$ Or a$ = user$ Then
          If b$ = "COMMERCIAL" And Alta_Comercial = 0 Then
             manager_commercial$ = a$
             
             Alta_Comercial = 1
             existe = 1
             'Exit For
          End If
             
        End If
        
        
        
        If a$ = agente$ Or a$ = csr$ Or a$ = user$ Then
          If b$ = "MANAGER_COMMERCIAL" Then
             manager_Oficinacommercial$ = a$
             existe = 1
             'Exit For
          End If
             
        End If
        
        
        If a$ = agente$ Or a$ = csr$ Or a$ = user$ Then
          If b$ = "AGENT_COMMERCIAL" Then
             Agent_Oficinacommercial$ = a$
             existe = 1
             'Exit For
          End If
             
        End If
        
        
        If a$ = agente$ Or a$ = csr$ Or a$ = user$ Then
            If b$ = "MONTERREY" Then
               Es_Monterrey = 1
              ' Exit For
            End If
        End If
        
        
        If existe = 1 Or Es_Monterrey = 1 Then
            ' Exit For
        End If
        
        
      Next z
      
      Return
      
End Sub
