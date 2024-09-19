VERSION 5.00
Begin VB.Form forma_inicial 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4710
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8235
   ControlBox      =   0   'False
   Icon            =   "forma_inicial_goals.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin Project1.lvButtons_H btncancel 
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   4080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "&Cancel"
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
   Begin Project1.lvButtons_H btnOK 
      Height          =   615
      Left            =   7320
      TabIndex        =   3
      Top             =   3960
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      Caption         =   "&Ok"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   4
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
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtpass 
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   5640
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7200
      Top             =   3000
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Updated: September/2024"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 3.28"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   3810
      Left            =   0
      Picture         =   "forma_inicial_goals.frx":16B92
      Top             =   0
      Width           =   8235
   End
End
Attribute VB_Name = "forma_inicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim seg As Integer

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const GWL_STYLE = (-16)


Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Private Const MAX_COMPUTERNAME_LENGTH As Long = 31

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Dim pId As String
Dim OSInfo As OSVERSIONINFO


Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
Private Const TOKEN_QUERY As Long = &H8
Private Const SE_PRIVILEGE_ENABLED As Long = &H2

Private Const EWX_LOGOFF As Long = &H0
Private Const EWX_SHUTDOWN As Long = &H1
Private Const EWX_REBOOT As Long = &H2
Private Const EWX_FORCE As Long = &H4
Private Const EWX_POWEROFF As Long = &H8
Private Const EWX_FORCEIFHUNG As Long = &H10 '2000/XP only

Private Const VER_PLATFORM_WIN32_NT As Long = 2



' ********************************************************************
' aqui empieza
Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4
  
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
  
Dim OReg As Registro

  
  
' aqui acaba
' **********************************************************************


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
  ByVal lpParameters As String, ByVal lpDirectory As String, _
  ByVal nShowCmd As Long) As Long



Private Type LUID
   dwLowPart As Long
   dwHighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
   udtLUID As LUID
   dwAttributes As Long
End Type

Private Type TOKEN_PRIVILEGES
   PrivilegeCount As Long
   laa As LUID_AND_ATTRIBUTES
End Type
      
Private Declare Function ExitWindowsEx Lib "user32" _
   (ByVal dwOptions As Long, _
   ByVal dwReserved As Long) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function OpenProcessToken Lib "advapi32" _
  (ByVal ProcessHandle As Long, _
   ByVal DesiredAccess As Long, _
   TokenHandle As Long) As Long

Private Declare Function LookupPrivilegeValue Lib "advapi32" _
   Alias "LookupPrivilegeValueA" _
  (ByVal lpSystemName As String, _
   ByVal lpName As String, _
   lpLuid As LUID) As Long

Private Declare Function AdjustTokenPrivileges Lib "advapi32" _
  (ByVal TokenHandle As Long, _
   ByVal DisableAllPrivileges As Long, _
   NewState As TOKEN_PRIVILEGES, _
   ByVal BufferLength As Long, _
   PreviousState As Any, _
   ReturnLength As Long) As Long


Public Function GetIPHostName() As String
On Error Resume Next
    Dim sHostName As String * 256
    
    If Not SocketsInitialize() Then
        GetIPHostName = ""
        Exit Function
    End If
    
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPHostName = ""
        MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
                " has occurred.  Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    
    GetIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
    SocketsCleanup

End Function


Public Function HiByte(ByVal wParam As Integer) As Byte
  On Error Resume Next
  'note: VB4-32 users should declare this function As Integer
   HiByte = (wParam And &HFF00&) \ (&H100)
 
End Function

Public Function LoByte(ByVal wParam As Integer) As Byte
On Error Resume Next
  'note: VB4-32 users should declare this function As Integer
   LoByte = wParam And &HFF&

End Function

Public Sub SocketsCleanup()
On Error Resume Next
    If WSACleanup() <> ERROR_SUCCESS Then
        MsgBox "Socket error occurred in Cleanup."
    End If
    
End Sub
Public Function SocketsInitialize() As Boolean
On Error Resume Next

   Dim WSAD As WSADATA
   Dim sLoByte As String
   Dim sHiByte As String
   
   If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
      MsgBox "The 32-bit Windows Socket is not responding."
      SocketsInitialize = False
      Exit Function
   End If
   
   
   If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        MsgBox "This application requires a minimum of " & _
                CStr(MIN_SOCKETS_REQD) & " supported sockets."
        
        SocketsInitialize = False
        Exit Function
    End If
   
   
   If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
     (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
      HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
      
      sHiByte = CStr(HiByte(WSAD.wVersion))
      sLoByte = CStr(LoByte(WSAD.wVersion))
      
      MsgBox "Sockets version " & sLoByte & "." & sHiByte & _
             " is not supported by 32-bit Windows Sockets."
      
      SocketsInitialize = False
      Exit Function
      
   End If
    
    
  'must be OK, so lets do it
   SocketsInitialize = True
        
End Function

Private Sub btncancel_Click()
End
End Sub

Private Sub btnOK_Click()
On Error Resume Next
If txtpass.Text = "Ja456!" Then
   user_sistema$ = "FULL ACCESS"
   forma_inicial.Hide
   Load form1
   form1.Show
   Unload forma_inicial
   
ElseIf txtpass.Text = "JA123456!" Then
   
   user_sistema$ = "ONLY READ"
   forma_inicial.Hide
   Load form1
   form1.Show
   Unload forma_inicial
   
   
Else
   MsgBox "Password is invalid. Access denied.", 16, "Attention"
End If


End Sub

Private Sub Form_Load()
On Error Resume Next
forma_inicial.Top = ((Screen.Height - Height) / 2) - 3000
forma_inicial.Left = ((Screen.Width - Width) / 2)


  a$ = GetIPHostName()

  nf = FreeFile
  Open "\\192.168.84.215\Goals_update\" + a$ + "-in" For Output Shared As #nf
  Lock #nf
  Print #nf, Format(Now, "mm/dd/yyyy  hh:mm am/pm")
  Unlock #nf
  Close #nf
  
  
  

 actualiza = 0
  nf = FreeFile
  Open "\\192.168.84.215\Goals_update\version.txt" For Input Shared As #nf
  Lock #nf
  Line Input #nf, version_actual$
  Unlock #nf
  Close #nf
  
  nf = FreeFile
  Open "c:\goals\version.txt" For Input Shared As #nf
  Lock #nf
  Line Input #nf, version_programa$
  Unlock #nf
  Close #nf
  
  If Val(version_programa$) < Val(version_actual$) Then
     actualiza = 1
     r$ = Shell("\\192.168.84.215\goals_update\actualizador.exe", vbNormalFocus)
     
     Hide
     Refresh
     End
     
  End If
  
End Sub

Private Sub Timer1_Timer()
seg = seg + 1
If seg >= 4 Then
 
End If
End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
  btnOK_Click
End If
End Sub


