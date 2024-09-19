VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form HR 
   BackColor       =   &H80000016&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HR"
   ClientHeight    =   15030
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   28710
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   15030
   ScaleWidth      =   28710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.lvButtons_H btntransfiere 
      Height          =   375
      Left            =   6360
      TabIndex        =   23
      Top             =   6840
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
      Mode            =   0
      Value           =   0   'False
      Image           =   "HR_goals.frx":0000
      ImgSize         =   40
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btncvs 
      Height          =   855
      Left            =   9600
      TabIndex        =   20
      Top             =   7080
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1508
      Caption         =   "&Export Commisions (CSV)"
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
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   1
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "HR_goals.frx":0452
      ImgSize         =   32
      Enabled         =   0   'False
      cBack           =   12632256
   End
   Begin MSComctlLib.ProgressBar barra 
      Height          =   105
      Left            =   6720
      TabIndex        =   25
      Top             =   7200
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
      Max             =   2
      Scrolling       =   1
   End
   Begin Project1.lvButtons_H btnlimpia 
      Height          =   255
      Left            =   9600
      TabIndex        =   24
      Top             =   6600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      CapAlign        =   2
      Shape           =   5
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
      Image           =   "HR_goals.frx":0C06
      ImgSize         =   40
      cBack           =   -2147483633
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   6720
      Sorted          =   -1  'True
      TabIndex        =   21
      Top             =   6720
      Width           =   3015
   End
   Begin Project1.lvButtons_H btnloadinfo 
      Height          =   375
      Left            =   18240
      TabIndex        =   18
      Top             =   240
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "Load info"
      CapAlign        =   2
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
      cBack           =   -2147483633
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   16440
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Project1.lvButtons_H op_tipo 
      Height          =   255
      Index           =   0
      Left            =   12240
      TabIndex        =   14
      Top             =   6960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      Caption         =   "Only &Agents"
      CapAlign        =   2
      BackStyle       =   7
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
      Mode            =   2
      Value           =   -1  'True
      cBack           =   12632256
   End
   Begin VB.PictureBox msg 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   5040
      ScaleHeight     =   825
      ScaleWidth      =   10665
      TabIndex        =   12
      Top             =   3120
      Visible         =   0   'False
      Width           =   10695
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
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   10455
      End
   End
   Begin VB.ComboBox cboimpre 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   12120
      TabIndex        =   10
      Top             =   6240
      Width           =   3615
   End
   Begin VB.ComboBox cboyear 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3880
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   5960
      Width           =   1095
   End
   Begin Project1.lvButtons_H btnload 
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   7560
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      Caption         =   "&Load"
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
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   1
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   32
      cBack           =   12632256
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   6440
      Width           =   2775
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   3120
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   6720
      Width           =   3135
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   1815
      Left            =   18600
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3201
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin Project1.lvButtons_H btnok 
      Height          =   495
      Left            =   18960
      TabIndex        =   1
      Top             =   7440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      Caption         =   "&OK"
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
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   5535
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   19455
      _ExtentX        =   34316
      _ExtentY        =   9763
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   12632256
      BackColorSel    =   16761024
      GridColorFixed  =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin Project1.lvButtons_H btnprint 
      Height          =   375
      Left            =   14760
      TabIndex        =   11
      Top             =   6720
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   661
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
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   1
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   32
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H op_tipo 
      Height          =   255
      Index           =   1
      Left            =   13440
      TabIndex        =   15
      Top             =   6960
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      Caption         =   "All &Users"
      CapAlign        =   2
      BackStyle       =   7
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
      Mode            =   2
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin Project1.lvButtons_H btnexcel 
      Height          =   855
      Left            =   17160
      TabIndex        =   17
      Top             =   6000
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1508
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
      cBhover         =   8421504
      cGradient       =   8421504
      Gradient        =   1
      CapStyle        =   1
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "HR_goals.frx":17EB
      ImgSize         =   32
      cBack           =   12632256
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Two weeks"
      Height          =   255
      Left            =   7680
      TabIndex        =   26
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   " From                        To"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   6840
      TabIndex        =   22
      Top             =   6480
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Year:"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   3360
      TabIndex        =   19
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   12240
      TabIndex        =   16
      Top             =   6600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   615
      Left            =   12120
      Shape           =   4  'Rounded Rectangle
      Top             =   6720
      Width           =   2175
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   12240
      TabIndex        =   9
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " From                        To"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   6480
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select the payday:"
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
      Left            =   480
      TabIndex        =   4
      Top             =   6120
      Width           =   2175
   End
End
Attribute VB_Name = "HR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim payday$, fromdate$, todate$

Dim DesignX As Integer
      Dim DesignY As Integer
Dim primeravez As Integer

Dim Tipo As Integer






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

Private Sub btncvs_Click()
On Error Resume Next
Dim sData As String

cd1.DialogTitle = "Save File"
    cd1.InitDir = "c:\goals"
    cd1.Filter = "CSV Files (*.csv)|*.CSV "
    cd1.FilterIndex = 1
    cd1.flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNExplorer
    cd1.CancelError = True

  cd1.ShowSave
  n$ = cd1.FileName
  
  
If n$ = "" Then Exit Sub


lblmsg.Caption = "Please, wait a moment... transferring data to Excel CSV"

msg.Visible = True
msg.Refresh



' carga primera semana
payday$ = List1.List(List1.ListIndex)



col = InStr(1, List3.List(0), ",")
fromdate$ = Left(List3.List(0), col - 1)

todate$ = Right(List3.List(0), Len(List3.List(0)) - col)
cargar_payroll





nf = FreeFile
Open "c:\goals\week1" For Output Shared As #nf


sData = "Employee ID,EIN Tax ID,Pay Statement Type,Start Date,End Date,E OT Bonus Amount,E Commission Amount" + vbNullString


Lock #nf
Print #nf, sData
Unlock #nf


tax_id$ = "26-1432097"

Grid1.row = 1
Grid1.col = 3
fecha_inicio$ = Format(Grid1.Text, "mm/dd/yyyy")

Grid1.col = 4
fecha_final$ = Format(Grid1.Text, "mm/dd/yyyy")


For t = 1 To Grid1.Rows - 1
   
   Grid1.row = t
   Grid1.col = 19
   id_empleado$ = Grid1.Text
   
   Grid1.col = 16
   commision$ = Format(Grid1.Text, "####0.00")
   
   sData = id_empleado$ + "," + tax_id$ + ",Regular," + fecha_inicio$ + "," + fecha_final$ + ", 0, " + commision$ + vbNullString  ' + vbCr + vbLf
   
   
   Lock #nf
   Print #nf, sData
   Unlock #nf


Next t

Close #nf


' ------------------------------------------------------------------------------------------------------
' carga semana DOS


col = InStr(1, List3.List(1), ",")
fromdate$ = Left(List3.List(1), col - 1)

todate$ = Right(List3.List(1), Len(List3.List(1)) - col)
cargar_payroll





nf = FreeFile
Open "c:\goals\week2" For Output Shared As #nf


'sData = "Employee ID,EIN Tax ID,Pay Statement Type,Start Date,End Date,E OT Bonus Amount,E Commission Amount" + vbNullString


'Lock #nf
'Print #nf, sData
'Unlock #nf


tax_id$ = "26-1432097"

Grid1.row = 1
Grid1.col = 3
fecha_inicio$ = Format(Grid1.Text, "mm/dd/yyyy")

Grid1.col = 4
fecha_final$ = Format(Grid1.Text, "mm/dd/yyyy")


For t = 1 To Grid1.Rows - 1
   
   Grid1.row = t
   Grid1.col = 19
   id_empleado$ = Grid1.Text
   
   Grid1.col = 16
   commision$ = Format(Grid1.Text, "####0.00")
   
   sData = id_empleado$ + "," + tax_id$ + ",Regular," + fecha_inicio$ + "," + fecha_final$ + ", 0, " + commision$ + vbNullString  ' + vbCr + vbLf
   
   
   Lock #nf
   Print #nf, sData
   Unlock #nf


Next t

Close #nf



' fusiona ambos archivos
' ******************************************



nf = FreeFile
Open "c:\goals\week1" For Append Shared As #nf

nf2 = FreeFile
Open "c:\goals\week2" For Input Shared As #nf2


Do Until EOF(nf2)
   Lock #nf2
   Line Input #nf2, r$
   Unlock #nf2
   
   Lock #nf
   Print #nf, r$
   Unlock #nf
Loop

Close nf, nf2

FileCopy "c:\goals\week1", n$

Kill "c:\goals\week2"
Kill "c:\goals\week1"








msg.Visible = False
msg.Refresh


End Sub

Private Sub btnexcel_Click()
On Error Resume Next

 Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset
    
    
'If lista_agentes.ListCount = 0 Then Exit Sub
If Grid1.Rows < 3 Then Exit Sub

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
Dim DataArray(1 To 100, 1 To 22) As Variant
Dim r As Integer



grandtotal = 0
num = 0
For t = 1 To Grid1.Rows - 1
   
   Grid1.row = t
   Grid1.col = 1
   id_emp = Val(Grid1.Text)
   
   Grid1.col = 2
   n$ = Grid1.Text
   
     
   
   
   
   Grid1.col = 1
   idempLAE$ = Grid1.Text
   sSelect = "select idemppayrolllink, initials from payrollconfig where idemployeelae='" + idempLAE$ + "'"
   
   Rs.Open sSelect, base, adOpenUnspecified
    
   id_emppayrolllink$ = Rs(0)
   iniciales$ = Rs(1)
                         
   Rs.Close
 
   
   
   
   If Tipo = 0 Then
   
   ' carga el usuario del agente del LAE
    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT username From employeeinfo where firstname='" + nombre$ + "' and lastname1='" + apellido$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    agente$ = Rs(0)
                         
    Rs.Close
    
    
   
   
   ' detecta si es manager y lo marca
    existe = 0
    For Y = 0 To form1.lista_managers.ListCount - 1
      n2$ = UCase(RTrim(Left(form1.lista_managers.List(Y), 20)))
      If n2$ = UCase(agente$) Then
          puesto$ = Right(form1.lista_managers.List(Y), Len(form1.lista_managers.List(Y)) - 20)
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

   Else
     existe = 0
   
   End If

   If existe = 0 Then
   
    Grid1.col = 10
    cant = Val(Grid1.Text)
   
    grandtotal = grandtotal + cant

    num = num + 1
    DataArray(num, 1) = " "
    
    DataArray(num, 2) = n$    'Left(UCase(iniciales$), 3) + " -" + n$
    
    Grid1.col = 3
    DataArray(num, 3) = Format(Grid1.Text, "mm/dd/yyyy")
    
    Grid1.col = 4
    DataArray(num, 4) = Format(Grid1.Text, "mm/dd/yyyy")
    
    Grid1.col = 5
    DataArray(num, 5) = Format(Grid1.Text, "mm/dd/yyyy")
    
    Grid1.col = 6   ' rhour
    DataArray(num, 6) = Format(Grid1.Text, "#0.00")
    
    Grid1.col = 7    ' OT
    DataArray(num, 7) = Format(Grid1.Text, "#0.00")
    
    Grid1.col = 8    ' Ahour
    DataArray(num, 8) = Format(Grid1.Text, "$###,##0.00")
    
    Grid1.col = 9    ' NB
    DataArray(num, 9) = Grid1.Text   ' NB
        
    Grid1.col = 10
    DataArray(num, 10) = Format(Grid1.Text, "$###,##0.00")   'BF
    
    Grid1.col = 11
    DataArray(num, 11) = Format(Grid1.Text, "$###,##0.00")   ' INVOICE
    
    Grid1.col = 12
    DataArray(num, 12) = Format(Grid1.Text, "$###,##0.00")   ' DED
    
    Grid1.col = 13
    DataArray(num, 13) = Format(Grid1.Text, "0")   'NB DED
    
    Grid1.col = 14
    DataArray(num, 14) = Format("0", "$###,##0.00") ' bonus
    
    Grid1.col = 15
    If Val(Grid1.Text) = 0 Then Grid1.Text = ""
    DataArray(num, 15) = Grid1.Text ' %
        
    Grid1.col = 16
    DataArray(num, 16) = Format(Grid1.Text, "$###,##0.00")  ' COMISION
    
    Grid1.col = 17
    DataArray(num, 17) = ""  ' OT bounus
    
    
    Grid1.col = 18
    DataArray(num, 18) = Format(Grid1.Text, "$###,##0.00")  ' TOTAL
    
    Grid1.col = 19
    DataArray(num, 19) = id_emppayrolllink$
    
    Grid1.col = 20
    DataArray(num, 20) = Grid1.Text
    
   
   End If



Next t


'Add headers to the worksheet on row 1
Set oSheet = oBook.Worksheets(1)
oSheet.range("A1:T1").Value = Array(" ", "Employee", "From", "To", "Pay Day", "R. Hour", "OT", "A. Hour", "NB", "BF", "Invoice", "Deduction", "NB Deduc.", "Bonus", "%", "Comission", "OT Bonus", "Total", "ID", "Notes")

'Transfer the array to the worksheet starting at cell A2, -- I changed A2 by A1
oSheet.range("A2").Resize(100, 20).Value = DataArray


'lblgrandtotal.Caption = Format(grandtotal, "$###,##0.00")


cd1.DialogTitle = "Save File"
    cd1.InitDir = "c:\goals"
    cd1.Filter = "Excel Files (*.xlsx)|*.xlsx|All " & _
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
  
  
If n$ = "" Then Exit Sub





'Save the Workbook and Quit Excel
oBook.SaveAs n$
oExcel.Quit


msg.Visible = False
msg.Refresh


End Sub

Private Sub btnlimpia_Click()
On Error Resume Next
List3.Clear
barra.Value = 0
btncvs.Enabled = False
End Sub

Private Sub btnload_Click()
On Error Resume Next
If List1.ListIndex = -1 Then Exit Sub

payday$ = List1.List(List1.ListIndex)
col = InStr(1, List2.List(List2.ListIndex), ",")
fromdate$ = Left(List2.List(List2.ListIndex), col - 1)

todate$ = Right(List2.List(List2.ListIndex), Len(List2.List(List2.ListIndex)) - col)
cargar_payroll

End Sub

Private Sub btnloadinfo_Click()
On Error Resume Next
Dim xlApp2 As Excel.Application
Dim xlLibro2 As Excel.Workbook
Dim xlHoja2 As Excel.Worksheet
Dim varMatriz As Variant
Dim lngUltimaFila As Long
lblmsg.Caption = "Please, wait a moment... loading hourly wages"
msg.Visible = True
msg.Refresh

btn_dmv.Visible = False
contador = 0

inicio:
n$ = "c:\goals\wages.xlsx" '.csv"
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
grid6.cols = 19


cont = 0
For t = 1 To grid6.Rows - 2
  grid6.row = t - 1
  If t > 1 And t <= lngUltimaFila And varMatriz(t, 1) <> "" Then
     cont = cont + 1
     grid6.col = 0
     grid6.Text = cont
     
  End If
  For Y = 1 To 19
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


lin = 0
For t = 1 To grid6.Rows - 1
   grid6.row = t
   grid6.col = 1
   numrow$ = grid6.Text
   
   If Val(numrow$) <= 0 Or numrow$ = "" Then
      For z = 0 To grid6.cols - 1
          grid6.col = z
          grid6.Text = ""
      Next z
      lin = lin + 1
   End If
   
   
Next t

grid6.Rows = grid6.Rows - lin

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


semana1$ = ""
semana2$ = ""

grid6.col = 1
grid6.Sort = flexSortGenericAscending

For t = 1 To grid6.Rows - 1

 
   grid6.row = t
   grid6.col = 0
   
   grid6.Text = t
   
   
   grid6.col = 4
   grid6.Text = Format(grid6.Text, "mm/dd/yyyy")
   
   If semana1$ = "" Then
      semana1$ = grid6.Text
   End If
   
   
   If semana2$ = "" And grid6.Text <> semana1$ Then
      semana2$ = grid6.Text
   End If
   
   
 
  
  
   grid6.col = 5
   grid6.Text = Format(grid6.Text, "mm/dd/yyyy")
   
     
   
Next t


conta = 1
For t = 1 To grid6.Rows - 1

   existe = 0
   grid6.row = t
     
   
   grid6.col = 4
   f$ = Format(grid6.Text, "mm/dd/yyyy")
   
    
   
   If grid6.Text = semana1$ And existe = 0 Then
     grid6.col = 1
     grid6.Text = Format(grid6.Text, "00000") + f$
     conta = conta + 1
     existe = 1
   ElseIf grid6.Text = semana2$ And existe = 0 Then
     grid6.col = 1
     grid6.Text = Format(grid6.Text, "00000") + f$
     existe = 1
     conta = conta + 1
   End If
     
   
      
Next t


grid6.col = 1
grid6.Sort = flexSortGenericAscending



For t = 1 To grid6.Rows - 1
   grid6.row = t
     grid6.col = 1
     grid6.Text = Format(Left(grid6.Text, 5), "####0")
Next t





' agrega las columnas faltantes





msg.Visible = False
msg.Refresh


' separa_campos
btn_dmv.Visible = True
End Sub

Private Sub btnOK_Click()
Unload Me
End Sub

Private Sub btnprint_Click()
On Error Resume Next
If Grid1.Rows < 3 Then Exit Sub


X$ = MsgBox("Do you want to print the report?", 4, "Attention")
If X$ = "7" Then Exit Sub


lblmsg.Caption = "Please, wait a moment... preparing data for printing"

msg.Visible = True
msg.Refresh

'Create an array with 20 columns and 100 rows
Dim DataArray(100, 20)
Dim r As Integer

Dim sSelect As String
Dim Rs As ADODB.Recordset
    

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
For t = 1 To Grid1.Rows - 1
   
   Grid1.row = t
   Grid1.col = 19
   id_emp = Val(Grid1.Text)
   
   Grid1.col = 1
   nombre$ = Grid1.Text
   Grid1.col = 2
   apellido$ = Grid1.Text
   
   
   n$ = UCase$(nombre$ + " " + apellido$)
   
   
   If Tipo = 0 Then
   
   ' carga el usuario del agente del LAE
    
    Set Rs = New ADODB.Recordset
    sSelect = "SELECT username From employeeinfo where firstname='" + nombre$ + "' and lastname1='" + apellido$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    agente$ = Rs(0)
                         
    Rs.Close
    
    
   
   
   ' detecta si es manager y lo marca
    existe = 0
    For Y = 0 To form1.lista_managers.ListCount - 1
      n2$ = UCase(RTrim(Left(form1.lista_managers.List(Y), 20)))
      If n2$ = UCase(agente$) Then
          puesto$ = Right(form1.lista_managers.List(Y), Len(form1.lista_managers.List(Y)) - 20)
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

   Else
     existe = 0
   
   End If

   If existe = 0 Then
   
    Grid1.col = 10
    cant = Val(Grid1.Text)
   
    grandtotal = grandtotal + cant

    num = num + 1
    DataArray(num, 1) = " "
    DataArray(num, 2) = n$
    
    Grid1.col = 3
    DataArray(num, 3) = Format(Grid1.Text, "mm/dd/yyyy")
    
    Grid1.col = 4
    DataArray(num, 4) = Format(Grid1.Text, "mm/dd/yyyy")
    
    Grid1.col = 5
    DataArray(num, 5) = Format(Grid1.Text, "mm/dd/yyyy")
    
    Grid1.col = 6
    DataArray(num, 6) = Format(Grid1.Text, "$###,##0.00")
    
    Grid1.col = 7
    DataArray(num, 7) = Grid1.Text
    
    Grid1.col = 8
    DataArray(num, 8) = Format(Grid1.Text, "$###,##0.00")
    
    Grid1.col = 9
    DataArray(num, 9) = Grid1.Text
        
    Grid1.col = 10
    DataArray(num, 10) = Format(Grid1.Text, "$###,##0.00")
    
    Grid1.col = 11
    DataArray(num, 11) = Format(Grid1.Text, "$###,##0.00")
    
    Grid1.col = 12
    DataArray(num, 12) = Format(Grid1.Text, "$###,##0.00")
    
    Grid1.col = 13
    DataArray(num, 13) = Format(Grid1.Text, "#0")
    
    Grid1.col = 14
    DataArray(num, 14) = Grid1.Text
    
    Grid1.col = 15
    DataArray(num, 15) = Format(Grid1.Text, "$###,##0.00")
        
    Grid1.col = 16
    DataArray(num, 16) = Grid1.Text
    
    Grid1.col = 17
    DataArray(num, 17) = Format(Grid1.Text, "$###,##0.00")
    
    Grid1.col = 18
    DataArray(num, 18) = Format(Grid1.Text, "$###,##0.00")
    
    Grid1.col = 19
    DataArray(num, 19) = Grid1.Text
    
    Grid1.col = 20
    DataArray(num, 20) = Grid1.Text
    
   
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

Printer.Print " Employee              From       To         Pay-Day    R.Hour   OT  A.Hour   NB      BF      Invoice Deduction NBded TotNB  TotBF      %  Commision     Total    ID   "
Printer.Print Space(1)

For t = 1 To num
 'For Y = 1 To 19
 Printer.Print Format(DataArray(t, 1), "@");
 Printer.Print Format(Left(DataArray(t, 2), 20), "!@@@@@@@@@@@@@@@@@@@@") + Space(1); ' employee
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
 
 'If DataArray(t, 16) = "$50.00" Then
 '  DataArray(t, 16) = "$50"
 'End If
 
 Printer.Print Format(DataArray(t, 17), "@@@@@@@@") + Space(2); ' Commision
 Printer.Print Format(DataArray(t, 18), "@@@@@@@@") + Space(3); ' Total
 Printer.Print Format(DataArray(t, 19), "@@@@") + Space(1); 'ID
 
 Printer.Print " "
 
 
 Printer.FontBold = Not Printer.FontBold
 
Next t

 Printer.Print " "
 Printer.Print Space(60) + "TOTAL: " + Format(grandtotal, "$###,##0.00")
 

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
Printer.FontName = "Arial"
Printer.FontSize = 8
Printer.Print Format(Now, "mm/dd/yyyy") + " " + Format(Now, "hh:mm:ss am/pm")

Printer.EndDoc

msg.Visible = False
msg.Refresh
MsgBox "The report has been printed", 64, "Attention"
End Sub

Private Sub btntransfiere_Click()
On Error Resume Next
If List3.ListCount > 2 Or List2.ListCount = 0 Then
  Exit Sub
End If

existe = 0
For t = 0 To List3.ListCount - 1
  If List2.List(List2.ListIndex) = List3.List(t) Then
     existe = 1
     Exit For
  End If
Next t

If existe = 0 Then
  If List3.ListCount = 2 Then
    Exit Sub
  End If
  
  List3.AddItem List2.List(List2.ListIndex)
  barra.Value = List3.ListCount
End If


If List3.ListCount = 2 Then
  btncvs.Enabled = True
End If


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
 
End Sub


Private Sub cboyear_Click()
On Error Resume Next
carga_fechas_payday
End Sub


Private Sub Form_Load()
On Error Resume Next
Top = 0
Left = (Screen.Width - Width) / 2
payday$ = form1.txtdatepayday.Text
fromdate$ = form1.txtdatefrom.Text
todate$ = form1.txtdateto.Text
carga_impresoras

cboyear.Clear
ano_actual = Val(Format(Now, "yyyy"))
Y = ano_actual - 5
cboyear.AddItem Y

Y = ano_actual - 4
cboyear.AddItem Y
Y = ano_actual - 3
cboyear.AddItem Y
Y = ano_actual - 2
cboyear.AddItem Y
Y = ano_actual - 1
cboyear.AddItem Y
Y = ano_actual
cboyear.AddItem Y
Y = ano_actual + 1
cboyear.AddItem Y

cboyear.ListIndex = 5
Tipo = 0



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




carga_fechas_payday

If payday$ <> "" Then cargar_payroll


End Sub



Public Sub carga_fechas_payday()
On Error Resume Next
List2.Clear
 Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

   grid.Clear
   ' sSelect = "SELECT distinct payday from payrollgoals"
    
    sSelect = "SELECT distinct payday from employeePayee"
    
   
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    grid.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid.DataSource = Rs
         
    Rs.Close
    
    
    List1.Clear
    
    For t = 1 To grid.Rows - 1
      grid.row = t
      grid.col = 1
      n = Val(Left(grid.Text, 4))
      If n = cboyear.Text Then
         List1.AddItem grid.Text
      End If
    Next t
  


End Sub

Public Sub cargar_payroll()
On Error Resume Next

 Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset
   
   Grid1.Clear
   
   
 '  sSelect = "SELECT employeeinfo.FirstName, employeeinfo.LastName1, payrollgoals.fromdate,payrollgoals.todate, payrollgoals.payday,payrollgoals.rhour, payrollgoals.ot," & _
 '    "payrollgoals.ahour,payrollgoals.nb,payrollgoals.bf,payrollgoals.invoice,payrollgoals.deduction,payrollgoals.nbdeduction,payrollgoals.totalnb,payrollgoals.totalbf," & _
 '    "payrollgoals.percentage,payrollgoals.commision,payrollgoals.total,payrollgoals.idemployee,payrollgoals.notes " & _
 '    "from PayrollGoals inner join EmployeeInfo on employeeinfo.IDEmployee=PayrollGoals.IdEmployee where " & _
 '    "payday='" + payday$ + "' and fromdate='" + fromdate$ + "' and todate='" + todate$ + "'"
 
   
   
 sSelect = "select emppay.IdEmployee, CONCAT( emp.FirstName, ' ', emp.LastName1) as Employee, emppay.DateFrom, emppay.DateTo, emppay.PayDay, " & _
           "emppay.RegularHours, emppay.OvertimeHours, emppay.TotalHoursAmount, emppay.NB, emppay.BF, emppay.Invoice, " & _
           "emppay.BFDeduction, emppay.NBDeduction, emppay.bonus, emppay.Percentage, emppay.Comm, emppay.otbonus, emppay.totalpaymentempl, " & _
           "emp.PayrollLinkID, emppay.Notes from EmployeePayee emppay " & _
           "inner join EmployeeInfo emp on Emppay.IdEmployee=emp.IDEmployee " & _
           "inner join PayrollConfig payconf on emppay.IdEmployee=payconf.IdEmployeeLAE " & _
           "where payday='" + payday$ + "' and datefrom='" + fromdate$ + "' and dateto='" + todate$ + "' and payconf.export='1'"
           
   
   
   
   
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    Grid1.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid1.DataSource = Rs
         
    Rs.Close
    
    
    
    
    For t = 1 To Grid1.Rows - 1
       ' sSelect = "select payrolllinkid from employeeinfo where idemployee='"
       Grid1.row = t
       Grid1.col = 0
       Grid1.Text = t
    Next t
    
    
    
    
    
    Grid1.row = 0
    Grid1.ColWidth(0) = 500     ' num
    
    Grid1.ColWidth(1) = 900    ' LAe#    emppay.IdEmployee
    Grid1.col = 1
    Grid1.Text = "LAE#"
    
    
    
    Grid1.ColWidth(2) = 2600    ' nombre    CONCAT( emp.FirstName, ' ', emp.LastName1) as Employee
    Grid1.col = 2
    Grid1.Text = "Name"
    
    Grid1.ColWidth(3) = 1600     ' from     emppay.DateFrom
    Grid1.col = 3
    Grid1.Text = "From"
    
    Grid1.ColWidth(4) = 1600  ' To     emppay.DateTo
    Grid1.col = 4
    Grid1.Text = "To"
    
    Grid1.ColWidth(5) = 1600    ' payday     emppay.PayDay
    Grid1.col = 6
    Grid1.Text = "Payday"
    
    Grid1.ColWidth(6) = 900    'rhour      emppay.RegularHours
    Grid1.col = 6
    Grid1.Text = "Rhour"
    
    Grid1.ColWidth(7) = 900    ' OT       emppay.OvertimeHours
    Grid1.col = 7
    Grid1.Text = "OT"
    
    Grid1.ColWidth(8) = 900    ' ahour     emppay.TotalHoursAmount
    Grid1.col = 8
    Grid1.Text = "Ahour"
    
    Grid1.ColWidth(9) = 1200    ' NB    emppay.NB
    Grid1.col = 9
    Grid1.Text = "NB"
    
    Grid1.ColWidth(10) = 1600    ' BF     emppay.BF
    Grid1.col = 10
    Grid1.Text = "BF"
    
    Grid1.ColWidth(11) = 1200    ' invoice   emppay.Invoice
    Grid1.col = 11
    Grid1.Text = "Invoice"
    
    Grid1.ColWidth(12) = 1000    ' deduction       emppay.BFDeduction
    Grid1.col = 12
    Grid1.Text = "Deduction"
    
    Grid1.ColWidth(13) = 1000    ' nbdeduction       emppay.NBDeduction
    Grid1.col = 13
    Grid1.Text = "NB ded."
    
    Grid1.ColWidth(14) = 1000    ' BONUS       emppay.bonus
    Grid1.col = 14
    Grid1.Text = "Bonus"
    
    Grid1.ColWidth(15) = 1000    ' percentage     emppay.Percentage
    Grid1.col = 15
    Grid1.Text = "Percentage"
    
    Grid1.ColWidth(16) = 1100    ' Commision       emppay.Comm
    Grid1.col = 16
    Grid1.Text = "Commision"
    
    Grid1.ColWidth(17) = 900    ' OT Bonus      emppay.otbonus
    Grid1.col = 17
    Grid1.Text = "OT Bonus"
    
    Grid1.ColWidth(18) = 1100    ' total        emppay.TotalHoursAmount
    Grid1.col = 18
    Grid1.Text = "Total"
    
    Grid1.ColWidth(19) = 900    ' id        emp.PayrollLinkID
    Grid1.col = 19
    Grid1.Text = "ID"
    
    Grid1.ColWidth(20) = 3500    ' notes     emppay.Notes
    Grid1.col = 20
    Grid1.Text = "Notes"
    
    
    For t = 1 To 20
      Grid1.RowHeight(t) = 420
    Next t
    
    
    
End Sub

Public Sub carga_fechas_week()
On Error Resume Next

 Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

   grid.Clear
   'sSelect = "SELECT distinct fromdate, todate from payrollgoals where payday='" + List1.List(List1.ListIndex) + "'"
   sSelect = "SELECT distinct datefrom, dateto from employeePayee where payday='" + List1.List(List1.ListIndex) + "'"
    
   
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    grid.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid.DataSource = Rs
         
    Rs.Close
    
     List2.Clear
    
    For t = 1 To grid.Rows - 1
      f$ = ""
      grid.row = t
      grid.col = 1
      f$ = grid.Text
      grid.col = 2
      f$ = f$ + "," + grid.Text
      List2.AddItem f$
    Next t
    
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

Private Sub Grid1_EnterCell()
On Error Resume Next
'Clipboard.Clear
'Clipboard.SetText grid1.
End Sub


Private Sub List1_Click()
On Error Resume Next
btnlimpia_Click
carga_fechas_week
End Sub


Private Sub op_tipo_Click(Index As Integer)
Tipo = Index
End Sub


