VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form forma_DMV 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DMV monthly Goals"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6285
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox msg 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   240
      ScaleHeight     =   585
      ScaleWidth      =   5745
      TabIndex        =   17
      Top             =   3840
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label lblmsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Please, wait a moment... loading all data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   120
         Width           =   5295
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   5415
      Begin VB.ComboBox cboyear 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin Project1.lvButtons_H btnmes 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "January"
         CapAlign        =   2
         BackStyle       =   1
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
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnmes 
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   5
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "February"
         CapAlign        =   2
         BackStyle       =   1
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
      Begin Project1.lvButtons_H btnmes 
         Height          =   375
         Index           =   2
         Left            =   1800
         TabIndex        =   6
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "March"
         CapAlign        =   2
         BackStyle       =   1
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
      Begin Project1.lvButtons_H btnmes 
         Height          =   375
         Index           =   3
         Left            =   2640
         TabIndex        =   7
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "April"
         CapAlign        =   2
         BackStyle       =   1
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
      Begin Project1.lvButtons_H btnmes 
         Height          =   375
         Index           =   4
         Left            =   3480
         TabIndex        =   8
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "May"
         CapAlign        =   2
         BackStyle       =   1
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
      Begin Project1.lvButtons_H btnmes 
         Height          =   375
         Index           =   5
         Left            =   4320
         TabIndex        =   9
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "June"
         CapAlign        =   2
         BackStyle       =   1
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
      Begin Project1.lvButtons_H btnmes 
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "July"
         CapAlign        =   2
         BackStyle       =   1
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
      Begin Project1.lvButtons_H btnmes 
         Height          =   375
         Index           =   7
         Left            =   960
         TabIndex        =   11
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "August"
         CapAlign        =   2
         BackStyle       =   1
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
      Begin Project1.lvButtons_H btnmes 
         Height          =   375
         Index           =   8
         Left            =   1800
         TabIndex        =   12
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "September"
         CapAlign        =   2
         BackStyle       =   1
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
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnmes 
         Height          =   375
         Index           =   9
         Left            =   2640
         TabIndex        =   13
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "October"
         CapAlign        =   2
         BackStyle       =   1
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
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnmes 
         Height          =   375
         Index           =   10
         Left            =   3480
         TabIndex        =   14
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "November"
         CapAlign        =   2
         BackStyle       =   1
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
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin Project1.lvButtons_H btnmes 
         Height          =   375
         Index           =   11
         Left            =   4320
         TabIndex        =   15
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "December"
         CapAlign        =   2
         BackStyle       =   1
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
         Mode            =   2
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Year:"
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
         Left            =   195
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
   End
   Begin Project1.lvButtons_H btnok 
      Height          =   735
      Left            =   5160
      TabIndex        =   0
      Top             =   8520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      Caption         =   "&OK"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   6975
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   12303
      _Version        =   393216
      BackColor       =   -2147483633
      BackColorFixed  =   12632256
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5520
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Project1.lvButtons_H btn_excel 
      Height          =   1335
      Left            =   5160
      TabIndex        =   19
      Top             =   2160
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   2355
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
      Image           =   "forma_DMV_goals.frx":0000
      ImgSize         =   32
      cBack           =   12632256
   End
End
Attribute VB_Name = "forma_DMV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mes_actual As Integer, ano_actual As Integer, dias_actual As Integer

Private Sub btnload1_Click()
End Sub

Private Sub btn_excel_Click()
On Error Resume Next

 Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset
    
    
'If lista_agentes.ListCount = 0 Then Exit Sub
If grid.Rows < 3 Then Exit Sub

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
Dim DataArray(1 To 100, 1 To 4) As Variant
Dim r As Integer



grandtotal = 0
num = 0
For t = 1 To grid.Rows - 1
   
   grid.row = t
   grid.col = 1
   UserName$ = grid.Text
   
   grid.col = 2
   bf$ = grid.Text
   
   
   
    num = num + 1
    DataArray(num, 1) = Str$(t)
    
    DataArray(num, 2) = UserName$    'Left(UCase(iniciales$), 3) + " -" + n$
    
    Grid1.col = 3
    DataArray(num, 3) = bf$
    
   

Next t


'Add headers to the worksheet on row 1
Set oSheet = oBook.Worksheets(1)
oSheet.range("A1:C1").Value = Array(" ", "AGENT", "BF")

'Transfer the array to the worksheet starting at cell A2, -- I changed A2 by A1
oSheet.range("A2").Resize(100, 4).Value = DataArray


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
  
  
If n$ = "" Then
   Exit Sub
End If





'Save the Workbook and Quit Excel
oBook.SaveAs n$
oExcel.Quit


msg.Visible = False
msg.Refresh


End Sub

Private Sub btnmes_Click(Index As Integer)
On Error Resume Next
mes_actual = Index + 1



cargar_BF_DMV

End Sub

Private Sub btnOK_Click()
On Error Resume Next
Unload Me

End Sub

Public Sub cargar_BF_DMV()
On Error Resume Next
Dim sSelect As String
   
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset

Select Case mes_actual
Case 1, 3, 5, 7, 8, 10, 12
  dias_actual = 31
Case 4, 6, 9, 11
  dias_actual = 30
Case 2
   cant = (ano_actual / 4)
   residuo = cant - Int(cant)
   If residuo = 0 Then
      dias_actual = 29
   Else
      dias_actual = 28
   End If
End Select




    fecha1$ = Format(mes_actual, "00") + "/01/" + Format(ano_actual, "0000")
    fecha2$ = Format(mes_actual, "00") + "/" + Format(dias_actual, "00") + "/" + Format(ano_actual, "0000")
    
    sSelect = "select emp.Username,sum (recdtl.Amount) as 'BF DMV' from ReceiptsHDR rechdr " & _
    "inner join ReceiptsDTL recdtl on recdtl.IdReceiptHDR = rechdr.IDReceiptHDR " & _
    "inner join EmployeeInfo emp on emp.IDEmployee = rechdr.IdEmployeeUSR " & _
    "inner join InvoiceItemCatalog ii on ii.IdInvoiceItem = recdtl.IdInvoiceItem " & _
    "Where recdtl.IdInvoiceItem = 15 and  cast(rechdr.date as Date) >= '" + fecha1$ + "' " & _
    "AND cast( rechdr.date as Date) <= '" + fecha2$ + "' group by emp.Username"


     

    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
        
     ' Permitir redimensionar las columnas
    grid.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid.DataSource = Rs
                         
    Rs.Close
    
    
    
   grid.ColWidth(0) = 500
   grid.ColWidth(1) = 1700  ' Username
   grid.ColWidth(2) = 1000 ' BF

For t = 1 To grid.Rows - 1
    grid.row = t
    grid.col = 0
    grid.Text = t
Next t
   
    
End Sub

Private Sub cboyear_Click()
On Error Resume Next
ano_actual = cboyear.List(cboyear.ListIndex)

cargar_BF_DMV

End Sub


Private Sub Form_Load()
On Error Resume Next

Top = 0
Left = (Screen.Width - Width) / 2


cboyear.Clear
ano_actual = Format(Now, "yyyy")

cboyear.AddItem ano_actual - 2
cboyear.AddItem ano_actual - 1
cboyear.AddItem ano_actual
cboyear.AddItem ano_actual + 1

' asigna el año actual
For t = 0 To cboyear.ListCount - 1
  If ano_actual = cboyear.List(t) Then
     cboyear.ListIndex = t
     Exit For
  End If
Next t

mes_actual = Format(Now, "mm")
btnmes(mes_actual - 1).Value = True
' btnmes_Click (mes_actual2 - 1)

btnmes_Click (mes_actual - 1)


End Sub


