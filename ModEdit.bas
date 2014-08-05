Attribute VB_Name = "ModEdit"
'Declaraciones y funciones necesarias para implementar el editor.
Option Explicit




'Ambitos del alcance para búsquedas
Public Const AMB_TODO = 0      'Todo el texto
Public Const AMB_SELE = 1      'Sólo la selección

'Direcciones de desplazamiento para "DesplazaCursor()"
Public Const DIR_NUL = 0
Public Const DIR_DER = 1   'Hacia la derecha
Public Const DIR_IZQ = 2   'Hacia la izquierda
Public Const DIR_ARR = 3   'Hacia arriba
Public Const DIR_ABA = 4   'Hacia abajo
Public Const DIR_PARR = 5   'Pagina arriba
Public Const DIR_PABA = 6   'Pagina abajo
Public Const DIR_INI = 7   'Inicio de líneas
Public Const DIR_FIN = 8   'Fin de líneas
Public Const DIR_HOM = 9   'Inicio de texto
Public Const DIR_END = 10   'Fin de texto
Public Const DIR_DERPAL = 11   'Dirección a la derecha por palabra
Public Const DIR_IZQPAL = 12  'Dirección a la izquierda por palabra
Public Const DIR_ARRPAR = 13  'Hacia arriba por párrafo
Public Const DIR_ABAPAR = 14  'Hacia abajo por párrafo

Public Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" ( _
        ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
  
' Estructura TEXTMETRIC
Public Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type


Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Type SIZE
    cx As Long
    cy As Long
End Type

Type POINTAPI
   x As Long
   y As Long
End Type

'Posición en texto. Coordenadas de texto
Type Tpostex
    xt As Integer
    yt As Long
End Type

'Descripción de segmentos
Type TDesSeg
    'txt As String   'No guarda un campo de tipo texto para no ocupar mucho espacio
    col As Long     'color de identificador
    tam As Integer  'tamaño de segmento. Se guarda tamaño en lugar de posicion porque
                    'es más fácil desplazar. Se usa Integer para economizar espacio
    tip As Integer  'tipo de segmento
End Type

'Tipos de segmentos
Public Const TSEG_DES = 0      'Segmento desconocido
Public Const TSEG_NOR = 1      'Segmento normal
Public Const TSEG_COM = 2      'Segmento comentario
Public Const TSEG_CAD = 3      'Segmento cadena
Public Const TSEG_PRS = 4      'Segmento palabra reservada
Public Const TSEG_PRS2 = 5     'Segmento palabra reservada 2
Public Const TSEG_FUN = 6      'Segmento función

'Descriptor de línea. Se usa para coloreado de sintaxis
Type TdesLin
    tip   As Long       'tipo de línea
    seg() As TDesSeg    'Vector de descripción de segmentos
End Type

'Tipos de línea para TdesLin
Public Const TLIN_DES = 0      'Línea desconocida, no analizado aún
Public Const TLIN_NOR = 1      'Línea de texto normal (1 solo segmento y color)
Public Const TLIN_MIX = 2      'Línea con segmentos mixtos
Public Const TLIN_COM = 3      'Línea de comentario (sin delimitador)
Public Const TLIN_CAD = 4      'Línea de cadena (sin delimitador)

'Tipo de acción para dehacer
Type Tundo
    acc As Integer  'acción realizada
    pos As Tpostex  'posición donde se ejecutó la acción
    cad As String   'cadena afectada
End Type

'Tipos de acciones, con su acción opuesta
Public Const TU_INS = 1
Public Const TU_ELI = -1
Public Const TU_INSn = 2    'Insertar modo normal
Public Const TU_ELIn = -2   'Eliminar modo normal
Public Const TU_SCOL = 3    'Pasa a modo columna
Public Const TU_SNOR = -3    'Pasa a modo normal

'Estilos de lapiz
Public Const PS_SOLID = 0
Public Const PS_DASH = 1
Public Const PS_DOT = 2
Public Const PS_DASHDOT = 3
Public Const PS_DASHDOTDOT = 4
Public Const PS_NULL = 5
Public Const PS_INSIDEFRAME = 6

'Fin de declaraciones

'*****************************************************************************
'********************* Funciones generales para el editor ********************
Public Sub InsertarCad(lin As String, pos As Integer, cad As String)
'Inserta una cadena a la línea indicada
    If pos > Len(lin) + 1 Then Exit Sub
    lin = Left$(lin, pos - 1) & cad & Mid$(lin, pos)
End Sub

Public Function EliminarCad(lin As String, pos As Integer, Optional tam As Integer = -1000)
'Elimina caracteres de la línea indicada. Devuelve la cadena eliminada
    If pos > Len(lin) Then Exit Function
    If pos < 1 Then Exit Function
    If tam = -1000 Then     'elimina hasta el final
        EliminarCad = Mid$(lin, pos)    'toma cadena
        lin = Left$(lin, pos - 1)
    Else
        EliminarCad = Mid$(lin, pos, tam)    'toma cadena
        lin = Left$(lin, pos - 1) & Mid$(lin, pos + tam)
    End If
End Function

Private Sub IntercPostex(p1 As Tpostex, p2 As Tpostex)
'Intercambia las posiciones p1 y p2
Dim p As Tpostex
    p = p1
    p1 = p2
    p2 = p
End Sub

Public Function MenorPos(p1 As Tpostex, p2 As Tpostex) As Boolean
'Compara dos posiciones del texto para ver si la primera está antes
    If p1.yt < p2.yt Then
        'No hay más que comparar, está antes
        MenorPos = True
        Exit Function
    ElseIf p1.yt > p2.yt Then
        'No hay más que comparar, está después
        MenorPos = False
        Exit Function
    Else
        'Están en la misma línea
        If p1.xt < p2.xt Then MenorPos = True Else MenorPos = False
    End If
End Function

Public Function MayorPos(p1 As Tpostex, p2 As Tpostex) As Boolean
'Compara dos posiciones del texto para ver si la primera está después
    If p1.yt > p2.yt Then
        'No hay más que comparar, está después
        MayorPos = True
        Exit Function
    ElseIf p1.yt < p2.yt Then
        'No hay más que comparar, está antes
        MayorPos = False
        Exit Function
    Else
        'Están en la misma línea
        If p1.xt > p2.xt Then MayorPos = True Else MayorPos = False
    End If
End Function
