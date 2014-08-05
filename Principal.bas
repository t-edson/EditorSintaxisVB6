Attribute VB_Name = "Principal"
'                                     Editor con Visual Basic
'
' Demuestra el uso del control ctlEdit.ctl, que incluye coloreado de sintaxis, y ayuda
' contextual.
'
' Está basado en el control usado por el Preprocesador de SQL PreSQL.
' Este código fuente puede ser usado, modificado y redistribuido libremente de acuerdo
' a su libre criterio, con solo indicar como referencia al autor.
'
'                                      Creado por Tito Hinostroza 04/03/2010 - Lima - Perú
'

Option Explicit

Public ed_ver_est As Boolean    'bandera de Editor - Ver estado
Public ed_ver_hor As Boolean    'bandera de Editor - Ver Barra Horizontal
Public ed_ver_ver As Boolean    'bandera de Editor - Ver Barra Vertical
Public ed_ver_num As Boolean    'bandera de Editor - Ver Número de línea

Public ActAyudCon As Boolean    'Activa ayuda contextual

'Constantes de Identificación del programa
Public Const NOMB_PROG = "VBEditor"

Const MSJ_ABRIR_ARCH = "Abrir Archivo"   'Mensaje del cuadro de diálogo
Const MSJ_GUARD_ARCH = "Guardar Archivo" 'Mensaje del cuadro de diálogo

'API's para cuadro de diálogo de abrir archivo. No se usa controles de VB para hacer más transportable el programa
Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

'------------------------------------------------------------------------
Public Function Posit(num As Single) As Single
'Devuelve siempre un número positivo o cero
    If num < 0 Then Posit = 0 Else Posit = num
End Function

Public Function DialogoAbrir(hWnd As Long, _
    tipArch As String, extArchiv As String, _
    Optional InitialDir As String = "C:\", _
    Optional titulo As String = MSJ_ABRIR_ARCH) As String
'Si se canceló devuelve "". Si se seleccionó un archivo, devuelve archivo.
Dim OpDlg As OPENFILENAME
Dim tmp As String
'Inicia estructura de Nombre de archivo
    OpDlg.lStructSize = Len(OpDlg)
    OpDlg.hwndOwner = hWnd              'Set the parent window
    OpDlg.hInstance = App.hInstance     'Set the application's instance
    'Select a filter
    OpDlg.lpstrFilter = tipArch + Chr$(0) + extArchiv + Chr$(0) + "Todos (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    OpDlg.lpstrFile = Space$(254)       'create a buffer for the file
    OpDlg.nMaxFile = 255                'set the maximum length of a returned file
    OpDlg.lpstrFileTitle = Space$(254)  'Create a buffer for the file title
    OpDlg.nMaxFileTitle = 255           'Set the maximum length of a returned file title
    OpDlg.lpstrInitialDir = InitialDir  'Set the initial directory
    OpDlg.lpstrTitle = titulo           'Set the title
    OpDlg.flags = 0                     'No flags

    If GetOpenFileName(OpDlg) <> 0 Then
        tmp = OpDlg.lpstrFile
    End If
    'Quita caracter nulo final y espacios
    tmp = RTrim(tmp)
    If tmp = "" Then Exit Function
    If Asc(Right$(tmp, 1)) = 0 Then
        DialogoAbrir = Left$(tmp, Len(tmp) - 1)
    End If
End Function

Public Function DialogoGuardar(hWnd As Long, _
    tipArch As String, extArchiv As String, _
    Optional InitialDir As String = "C:\", _
    Optional titulo As String = MSJ_GUARD_ARCH) As String
Dim OFName As OPENFILENAME
Dim arch As String
'Inicia estructura de Nombre de archivo
    OFName.lStructSize = Len(OFName)    'Set the structure size
    OFName.hwndOwner = hWnd          'Set the owner window
    OFName.hInstance = App.hInstance    'Set the application's instance
    'Set the filter
    OFName.lpstrFilter = tipArch + Chr$(0) + extArchiv + Chr$(0) + "Todos (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    OFName.lpstrFile = Space$(254)      'Create a buffer
    OFName.nMaxFile = 255               'Set the maximum number of chars
    OFName.lpstrFileTitle = Space$(254)     'Create a buffer
    OFName.nMaxFileTitle = 255          'Set the maximum number of chars
    OFName.lpstrInitialDir = InitialDir       'Set the initial directory
    OFName.lpstrTitle = titulo      'Set the dialog title
    OFName.flags = 0                    'no extra flags
    'Show the 'Save File'-dialog
    If GetSaveFileName(OFName) Then
        arch = Trim$(OFName.lpstrFile)
        'quita chr(0) al final
        If Asc(Right$(arch, 1)) = 0 Then arch = Left$(arch, Len(arch) - 1)
        DialogoGuardar = arch
    Else
        DialogoGuardar = ""
    End If
End Function



