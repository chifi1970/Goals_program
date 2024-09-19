Attribute VB_Name = "Module1"
Global AI(50, 6)
Global JA(50, 6)
Global ruta$, ruta_pdf$
Global tipomsg
Global user_sistema$

' esto es para obtener el IP

Public Const MAX_WSADescription As Long = 256
Public Const MAX_WSASYSStatus As Long = 128
' Public Const ERROR_SUCCESS       As Long = 0
Public Const WS_VERSION_REQD     As Long = &H101
Public Const WS_VERSION_MAJOR    As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR    As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD    As Long = 1
Public Const SOCKET_ERROR        As Long = -1

Public Type HOSTENT
   hName      As Long
   hAliases   As Long
   hAddrType  As Integer
   hLen       As Integer
   hAddrList  As Long
End Type

Public Type WSADATA
   wVersion      As Integer
   wHighVersion  As Integer
   szDescription(0 To MAX_WSADescription)   As Byte
   szSystemStatus(0 To MAX_WSASYSStatus)    As Byte
   wMaxSockets   As Integer
   wMaxUDPDG     As Integer
   dwVendorInfo  As Long
End Type

Public Declare Function WSAGetLastError Lib "wsock32" () As Long

Public Declare Function WSAStartup Lib "wsock32" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Public Declare Function WSACleanup Lib "wsock32" () As Long

Public Declare Function gethostname Lib "wsock32" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
   
Public Declare Function gethostbyname Lib "wsock32" _
  (ByVal szHost As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" (hpvDest As Any, _
   ByVal hpvSource As Long, _
   ByVal cbCopy As Long)
   
   
   



Type oficinx
  office As String * 30
  location As String * 40
  subnet As String * 15
  public_ip As String * 15
  router_type As String * 30
  kyocera As String * 15
  haw As String * 15
  nota As String * 18
  empresa As Integer
  telefono As String * 15
End Type

Global Const tam_reg = 195
Global oficina As oficinx




'Type oficinx2
'  office As String * 30
'  location As String * 40
'  subnet As String * 15
'  public_ip As String * 15
'  router_type As String * 30
'  kyocera As String * 15
'  haw As String * 15
'  nota As String * 18
'  empresa As Integer
'  telefono As String * 15
'End Type

'Global Const tam_reg2 = 195
'Global oficina2 As oficinx2





Type attributos_agente

   nombre As String * 20
   manager As Integer
   monterrey As Integer
   Commercial As Integer
   excepcion1 As Integer
   tipo_excep1 As Integer
   excepcion2 As Integer
   tipo_excep2 As Integer
   excepcion3 As Integer
   tipo_excep3 As Integer
   excepcion4 As Integer
   tipo_excep4 As Integer
   excepcion5 As Integer
   tipo_excep5 As Integer
   iniciales As String * 3
   phonesales As Integer
   id As Integer
   id_LAE As Integer
   oficina As Integer
   oficina2 As Integer
   nombre_completo As String * 30
   ded_mod As Single
   nb_ded_mod As Integer
   porc_mod As Integer
   autorizado As String * 20
   categoria_pago As Integer
   
End Type

Global Const tam_atributo = 119
Global attr As attributos_agente





'Type attributos_agente2
'   nombre As String * 20
'   manager As Integer
'   monterrey As Integer
'   Commercial As Integer
'   excepcion1 As Integer
'   tipo_excep1 As Integer
'   excepcion2 As Integer
'   tipo_excep2 As Integer
'   excepcion3 As Integer
'   tipo_excep3 As Integer
'   excepcion4 As Integer
'   tipo_excep4 As Integer
'   excepcion5 As Integer
'   tipo_excep5 As Integer
'   iniciales As String * 3
'   phonesales As Integer
'   id As Integer
'   id_LAE As Integer
'   oficina As Integer
'   oficina2 As Integer
   
'   nombre_completo As String * 30
'   ded_mod As Single
'   nb_ded_mod As Integer
'   porc_mod As Integer
'   autorizado As String * 20
'   categoria_pago As Integer
   
'End Type

'Global Const tam_atributo2 = 119
'Global attr2 As attributos_agente2




Type extens1
  num As Integer
  name As String * 18
  tel As Integer
End Type

Global Const tam_ext = 22
Global ext As extens1
  
Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer ' As Long para que funcione en Windows XP con VB6
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type


Public Const FO_DELETE = &H3
Public Const FOF_ALLOWUNDO = &H40
Public Const FOF_NOCONFIRMATION = &H10
Public Const FOF_SILENT = &H4



Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
    (lpFileOp As SHFILEOPSTRUCT) As Long


Public Function ShellDelete(ParamArray vntFileName() As Variant) As Long

    Dim i As Integer
    Dim sFileNames As String
    Dim SHFileOp As SHFILEOPSTRUCT

    For i = LBound(vntFileName) To UBound(vntFileName)
    sFileNames = sFileNames & vntFileName(i) & vbNullChar
    Next
    sFileNames = sFileNames & vbNullChar

    With SHFileOp
        .wFunc = FO_DELETE
        .pFrom = sFileNames
        .fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION

 'FOF_ALLOWUNDO
    End With

    ShellDelete = SHFileOperation(SHFileOp)

End Function

Sub Resize_For_Resolution(ByVal SFX As Single, _
       ByVal SFY As Single, MyForm As Form)
       On Error Resume Next
       
      Dim i As Integer
      Dim SFFont As Single

      SFFont = (SFX + SFY) / 2  ' average scale
      ' Size the Controls for the new resolution
      On Error Resume Next  ' for read-only or nonexistent properties
      With MyForm
        For i = 0 To .Count - 1
         If TypeOf .Controls(i) Is ComboBox Then   ' cannot change Height
           .Controls(i).Left = .Controls(i).Left * SFX
           .Controls(i).Top = .Controls(i).Top * SFY
           .Controls(i).Width = .Controls(i).Width * SFX
         Else
           .Controls(i).Move .Controls(i).Left * SFX, _
            .Controls(i).Top * SFY, _
            .Controls(i).Width * SFX, _
            .Controls(i).Height * SFY
         End If
           .Controls(i).FontSize = .Controls(i).FontSize * SFFont
        Next i
        If RePosForm Then
          ' Now size the Form
          .Move .Left * SFX, .Top * SFY, .Width * SFX, .Height * SFY
        End If
      End With
End Sub

