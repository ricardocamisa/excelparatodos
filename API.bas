Attribute VB_Name = "API"
Option Explicit

Public Const IDC_HAND = 32649&

'--------------------------------------------------------------
#If VBA7 Then
    Public Declare PtrSafe Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Public Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare PtrSafe Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
    Public Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
    Public Declare PtrSafe Function IniciaJanela Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Public Declare PtrSafe Function MoveJanela Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    'API DO ICON
    Public Declare PtrSafe Function IconApp Lib "user32" Alias "SendMessageA" ( _
                ByVal hwnd As Long, _
                ByVal wMsg As Long, _
                ByVal wParam As Long, _
                lParam As Any) As Long
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwNilliseconds As Long)
#Else
    Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
    Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
    Public Declare Function IniciaJanela Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Public Declare Function MoveJanela Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    'API DO ICON
    Public Declare Function IconApp Lib "user32" Alias "SendMessageA" ( _
                ByVal hwnd As Long, _
                ByVal wMsg As Long, _
                ByVal wParam As Long, _
                lParam As Any) As Long
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwNilliseconds As Long)
#End If

Public MeuForm As Long
Public hIcone As Long
'// Contantes do Ícone
Public Const FOCO_ICONE = &H80
Public Const ICONE = 0&
Public Const GRANDE_ICONE = 1&

Private Const WS_EX_DLGMODALFRAME As Long = &H1
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public ESTILO As Long
Public ESTILO2 As Long
Public STYLE_FORM As Long
Public Const GWL_STYLE As Long = (-16)
Public Const ESTILO_ATUAL As Long = (-16)
Public Const WS_CAPTION = 55000000 '&HCCCCC0 '55000000 '&HFDF2FB         '&HDC47BE          '&HFCCFCF ' &HCCCCC0                             '&HE0E0E0   'vbWhite '55000000
Public mdOriginX As Double
Public mdOriginY As Double
Public hWndForm As Long

Public Function MouseCursor(CursorType As Long)
  Dim lngRet As Long
  lngRet = LoadCursorBynum(0&, CursorType)
  lngRet = SetCursor(lngRet)
End Function

Public Function Maozinha()
    Call MouseCursor(IDC_HAND)
End Function
Public Sub removeTudo(objForm As Object)
    MeuForm = FindWindowA(vbNullString, objForm.Caption)
    ESTILO = ESTILO Or WS_CAPTION
    MoveJanela MeuForm, ESTILO_ATUAL, (ESTILO)
End Sub

Public Sub removeCaption(objForm As Object)
    'Hide title bar and border around userform
    Dim lngWindow As Long
    Dim lFrmHdl As Long
    lFrmHdl = FindWindowA(vbNullString, objForm.Caption)
'    'Build window and set window until you remove the caption, title bar and frame around the window
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow Or WS_CAPTION
    SetWindowLong lFrmHdl, GWL_STYLE, lngWindow
    lngWindow = GetWindowLong(lFrmHdl, ESTILO_ATUAL)
    lngWindow = lngWindow Or WS_EX_DLGMODALFRAME   'MAX MIN CLOSE SEM BORDA AO REDIMENSIONAR
'    lngWindow = lngWindow And WS_EX_DLGMODALFRAME
    SetWindowLong lFrmHdl, ESTILO_ATUAL, lngWindow
    DrawMenuBar lFrmHdl
End Sub

