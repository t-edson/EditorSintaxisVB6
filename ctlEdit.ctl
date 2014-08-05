VERSION 5.00
Begin VB.UserControl ctlEdit 
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2925
   ScaleHeight     =   1725
   ScaleWidth      =   2925
   Begin VB.ListBox lstCont 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   45
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   45
      Width           =   1095
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Width           =   1215
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   1920
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   1320
      Top             =   360
   End
   Begin VB.PictureBox pic 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "ctlEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'                            Control editor de texto.
' Este código fuente puede ser usado, modificado y redistribuido libremente de acuerdo
' a su libre criterio, con solo indicar como referencia al autor.
'
'Implementa las principales funciones de un editor de texto plano. Sólo
'trabaja con tipos de letra de ancho de caracter constante.
'No utiliza ningún control "TextBox" o "RichTextBox". Esta optimizado en velocidad
'y maneja bien archivos de hasta 50000 líneas.
'
'Las principales propiedades son:
'
'"verNumLin".- Permite mostrar el número de línea activando la propiedad
'"verDesHor".- Permite activar la barra de desplazamiento horizontal.
'"bloqText".- Bloquea el contenido de modificación. Lo convierte en un visor.
'"TipoSelec".- Define el tipo de selección 0=Normal; 1=Por columnas

'En el desarrollo del programa se han considerado dos tipos de coordenadas:
' *Coordenadas del cursor.- Refrentes a la posición relativa del cursor con respecto
'  a la pantalla. La esquina superior izquierda es siempre la posición (1,1). Las
'  variables que se refieren a coordenadas del cursor, son de tipo xc, yc
' *Coordenadas del texto.- Referentes a la posición relativa de un caracter con
'  respecto al texto completo como si estuviera impreso en una hoja suficientemente
'  grande. El texto se ve con las tabulaciones expandidas. La posición (1,1) es
'  siempre el primer caracter del texto. Las variables que se refieren a coordenadas
'  del texto, son de tipo xt, yt
'
'Se debe mejorar:
'* Verificar la administración de memoria dinámica. No se ha revisado detalladamente.
'* El desplazamiento con la barra vertical para más de 32000 líneas
'* Agregar la opción de edición con reemplazo.
'* Mejorarse las opciones de "deshacer" en modo columna.
'* Agregar la opción "Rehacer".
'* Mejorar la opción de búsqueda. Agregar Reemplazo.
'
'                                           Iniciado por Tito Hinostroza 27/10/2008
'                                         Continuado por Tito Hinostroza 28/10/2009
'                                         Modificado por Tito Hinostroza 25/11/2009
'                                         Modificado por Tito Hinostroza 11/12/2009
'                                         Modificado por Tito Hinostroza 06/01/2010
'                                         Modificado por Tito Hinostroza 19/01/2010
'
'                       Modificado por Tito Hinostroza 12/02/2010 Lima - Perú
'
'Se ha agregado opciones de dibujo de texto con color.
'Se ha mejorado el manejo de la opción de deshacer en modo columna. Aún hay varios
'aspectos por mejorar.
'Se ha agregado la acción de cambio de modo de selección como acción para deshacer
'
'                       Modificado por Tito Hinostroza 19/04/2010
'                       Modificado por Tito Hinostroza 02/06/2010
'Se mejoró la función de búsqueda, corrigiendo las rutinas.
'Se optimizó la rutina Dibujar(), controlando el evento Paint() del "Picture"
'                       Modificado por Tito Hinostroza 14/06/2010
'Se mejoró el modo de inserción en modo columna para poder insertar varios
'caracteres sin seleccionar varias veces.
'Se creó el evento CambiaModo para gestionar mejor el cambio de modo del editor
'con la combinación ALt-C, ya que los menús no responden a esta combinación de
'teclas.

Option Explicit

Const MAX_ANC_LIN = 32766    'Ancho máximo de línea en caracteres. < 32767
Const TAM_MAX_UNDO = 1000000 'Tamaño máximo de bytes usados en "Deshacer"
Const NAC_MAX_UNDO = 60      'Número máximo de acciones "Deshacer"

'Direcciones de ajuste horizontal para posicionar el cursor
Const A_NULO = 0
Const A_IZQ_TAB = 1
Const A_DER_TAB = 2

'Caracteres de desplazamiento por palabra
Const CAR_DESP_PAL = "[a-zA-Z0-9_$'.áéíóúÁÉÍÓÚ-]"
Const CAR_IDEN_VALM = "[A-Z0-9_$ÁÉÍÓÚÑ]"    'caracteres válidos para identificador (mayúscula)

Private Const GMEM_MOVEABLE = 2
Private Const GMEM_DDESHARE = &H2000

Public Event ArchivoSoltado(arc As String)  'Evento para archivo soltado
'Public Event TeclaEscape()                 'Indica la pulsación de la tecla escape
Public Event CambiaModo()     'Indica una solicitud de cambio de modo. El editor
                              'lo pide con la tecla ALt-C.
Public Event KeyDown(KeyCode As Integer, Shift As Integer)  'Tecla pulsada


Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long

Const CF_TEXT = 1
Const GHND = &H42

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" ( _
    ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
    ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, _
    ByVal wFlags As Long) As Long
  
Const TA_LEFT = 0
Const TA_RIGHT = 2
Const TA_CENTER = 6
Const TA_TOP = 0
Const TA_BOTTOM = 8
Const TA_BASELINE = 24

'Variables generales
Public archivo As String    'archivo a cargar con el método "CargarArch"



Public nEspTab As Integer   'número de espacios por tabulación
Public verNumLin As Boolean 'Muestra número de línea
Public verDesHor As Boolean 'Muestra barra de desplazamiento horizontal
Public verDesVer As Boolean 'Muestra barra de desplazamiento vertical

Public bloqText As Boolean   'bloquea el texto para modificación
Private tipSelec As Integer  'tipo de selección de texto
Public tipArch As Integer   'tipo de archivo: 0->DOS 1->UNIX
Public menuContext As Menu  'referencia al menú contextual del Editor

Private ancNLinP As Integer 'ancho de la columna de número de líneas en pixeles
'Variables de colores de texto
Private mColFonEdi As Long  'color de fondo del control de edición
Private mColFonSel As Long  'color del fondo de la selección
Private mColTxtNor As Long  'color de texto normal
Private mColTxtSel As Long  'color del texto de la selección
Private mColFonNli As Long  'color de fondo para número de línea

Private mColTxtCom As Long  'color de comentarios
Private mColTxtCad As Long  'color de constantes cadenas
Private mColPalRes As Long  'color de palabras Reservadas
Private mColPalRes2 As Long  'color de palabras Reservadas
Private mColTxtFun As Long  'color de palabras Reservadas

Public nlin As Long         'número de líneas del texto

Private fil1 As Long        'número de fila inicial en el control
Private fil2 As Long        'número de fila final en el control
Private nfilFin As Long     'número de fila final a mostrar en el control
Private maxLinVis As Integer  'máximo número de líneas visibles (que caben en la ventana)

Private col1 As Integer       'número de columna inicial en el control
Private col2 As Integer       'número de columna final en el control
Private maxColVis As Integer  'máximo número de columnas visibles (que caben en la ventana)
Private maxTamLin As Integer  'ancho máximo de las líneas

Private pintando As Boolean  'bandera usada por el evento "Paint"
Private Redibujar As Boolean 'bandera que indica que se debe redibujar el control
Private msjError As String   'Mensaje de error

'Variables de manejo de la selección
Public haysel As Boolean    'Indica si hay un bloque seleccionado
Private sel1 As Tpostex     'Posición inicial de la selección
Private sel2 As Tpostex     'Posición final de la selección
Private sel0 As Tpostex     'Posición donde se empieza a marcar la selección
Private sel1ant As Tpostex  'sel1 anterior
Private sel2ant As Tpostex  'sel2 anterior

'Variables de manejo del cursor
Private curXt As Integer    'número de caracter que apunta el cursor en coord. del texto
Private curYt As Long       'número de línea que apunta el cursor en coord. del texto
Private curXd As Integer    'coordenada X de cursor deseada. Se usa para desplazarse por líneas
Private curXt_ant As Integer 'posición anterior X del cursor con respecto a la ventana
Private curYt_ant As Long    'posición anterior Y del cursor con respecto a la ventana
Private curBorrado As Boolean   'Para controlar el parpadeo del cursor

Private xt0 As Integer      'coordenada X de cursor inicial. Usado para selección con ratón
Private yt0 As Long         'coordenada Y de cursor inicial. Usado para selección con ratón
Private curtmp As Integer   'contador para temporizar el parpadeo del cursor
Private linact As String    'línea actual completa apuntada por el cursor
Private cursorOn As Boolean 'bandera para activar o desactivar el cursor

Private anccarP As Single    'ancho de caracter en pixeles
Private altcarP As Single    'ancho de caracter en pixeles

'Variables para control de pantalla
Const MAX_LIN_EDI = 100000      'Máxima cantidad de líneas que soporta el editor
Const MAX_LIN_COL = 1000        'Máxima cantidad de líneas que se soportan con coloreado de sintaxis
Private linrea() As String      'texto cargado en el editor en líneas
Private lincol() As TdesLin     'Descripción de colores de Líneas para gráficar
Private deslin() As TDesSeg     'Para manejo de descripción de línea

Private pulsadoI As Boolean     'bandera de botón Izquierdo pulsado
Private ultBotPul As Integer    'útlimo botón pulsado
Private facdesV  As Single      'factor de desplazamiento

'Variables de los métodos gráficos
Private tR As RECT, tTR As RECT

Private hPen As Long            'lapíz nuevo
Private hBrush As Long          'brocha para el PIC
Private hFont As Long           'fuente
Private ret_pt As POINTAPI      'punto POINTAPI para uso temporal

'Variables para manejo de tabulaciones
Private ptabexp() As Integer    'posiciones de inicio de las tabulaciones en el texto expandido
Private ptabrea() As Integer    'posiciones real de las tabulaciones en el texto

'Variable para el control de la opción DESHACER
Private Undos() As Tundo       'Comandos deshacer
Public nUndo As Integer        'Número de acciones en Undos
Private nTxtModif As Integer   'Indice al nUndo que da el texto sin modificar
Private Deshaciendo As Boolean

'Variables para el control de la búsqueda
Private PosBus1 As Tpostex  'Posición inicial del texto buscado.
Private PosBus2 As Tpostex  'Posición final del texto buscado.
Private CadBus As String    'Cadena de búsqueda
Private CajBus As Boolean   'Bandera de caja para búsqueda
Private PalCBus As Boolean  'Bandera de palabra completa para búsqueda
Private DirBus As Integer   'Dirección de búsqueda
Private PosEnc As Tpostex   'Posición del texto encontrado. Usado para búsquedas

'Variables para ayuda contextual
Private HayAyudC As Boolean 'bandera de menú de Ayuda Contextual activa
Private xCont0 As Long      'Coordenada inicial de menú contextual
Private yCont0 As Long      'Coordenada final de menú contextual
Private xtIniIden As Integer    'Columna de inicio de identificador
Private ytIniIden As Long       'Fila de inicio de identificador
Private ancMenCon As Long       'Ancho de menú contextual
Private altMenCon As Long
Private IdentAyudC() As String  'Guarda los identificadores de la ayuda contextual
Private nFilAyudC As Integer    'Número de filas a mostrar en el menú
Private ListandoTab As Boolean
Private ArcListaTab As String   'Archivo de lista de Tablas
Private ListaTablas() As String 'Guarda los nombres de las tablas

'*************************************************************************************
'********************************FUNCIONES DE BAJO NIVEL******************************
'*************************************************************************************
Private Function Posit(num As Single) As Single
'Devuelve siempre un número positivo o cero
    If num < 0 Then Posit = 0 Else Posit = num
End Function

Public Property Get TipoSelec() As Integer
    TipoSelec = tipSelec
End Property

Public Property Let TipoSelec(ByVal vNewValue As Integer)
    If tipSelec <> vNewValue Then
        If vNewValue = 0 Then  'Pasa a modo normal
            GuarAcc TU_SNOR, LeePosCur(), ""     'para deshacer
        Else                'Pasa a modo por columna
            GuarAcc TU_SCOL, LeePosCur(), ""     'para deshacer
        End If
        tipSelec = vNewValue
    End If
End Property
'------------Lectura de colores--------------
Public Property Get ColFonEdi() As Long: ColFonEdi = mColFonEdi: End Property
Public Property Get ColFonSel() As Long: ColFonSel = mColFonSel: End Property
Public Property Get ColTxtNor() As Long: ColTxtNor = mColTxtNor: End Property
Public Property Get ColTxtSel() As Long: ColTxtSel = mColTxtSel: End Property
Public Property Get ColFonNli() As Long: ColFonNli = mColFonNli: End Property

Public Property Get ColTxtCom() As Long: ColTxtCom = mColTxtCom: End Property
Public Property Get ColTxtCad() As Long: ColTxtCad = mColTxtCad: End Property
Public Property Get ColPalRes() As Long: ColPalRes = mColPalRes: End Property
Public Property Get ColPalRes2() As Long: ColPalRes2 = mColPalRes2: End Property
Public Property Get ColTxtFun() As Long: ColTxtFun = mColTxtFun: End Property

'------------Escritura de colores--------------
Public Property Let ColFonEdi(ByVal vNewValue As Long)
    If mColFonEdi <> vNewValue Then
        mColFonEdi = vNewValue  'lee nuevo valor
        pic.BackColor = mColFonEdi  'Actualiza color de fondo
    End If
End Property

Public Property Let ColFonSel(ByVal vNewValue As Long): mColFonSel = vNewValue: End Property
Public Property Let ColTxtSel(ByVal vNewValue As Long): mColTxtSel = vNewValue: End Property
Public Property Let ColTxtNor(ByVal vNewValue As Long): mColTxtNor = vNewValue: End Property
Public Property Let ColFonNli(ByVal vNewValue As Long): mColFonNli = vNewValue: End Property

Public Property Let ColTxtCom(ByVal vNewValue As Long): mColTxtCom = vNewValue: End Property
Public Property Let ColTxtCad(ByVal vNewValue As Long): mColTxtCad = vNewValue: End Property
Public Property Let ColPalRes(ByVal vNewValue As Long): mColPalRes = vNewValue: End Property
Public Property Let ColPalRes2(ByVal vNewValue As Long): mColPalRes2 = vNewValue: End Property
Public Property Let ColTxtFun(ByVal vNewValue As Long): mColTxtFun = vNewValue: End Property
'------------------------------------------------
Private Function MaxCol1() As Integer
'Devuelve el máximo valor que puede tomar "col1"
    'Debería ser maxTamLin - maxColVis + 1, pero el cursor
    'se puede mover hasta len(linea)+1.
    MaxCol1 = maxTamLin - maxColVis + 2
    If MaxCol1 < 1 Then MaxCol1 = 1 'protección
End Function

Private Function CarXY(xt As Integer, ByVal yt As Long) As String
'Devuelve el caracter de la posición X,Y de la ventana visible.
Dim lin As String
    'Verifica si escapa de la pantalla
    If xt < col1 Or xt > col2 Then CarXY = " ": Exit Function
    If yt < 1 Or yt > nlin Then CarXY = " ": Exit Function
    'toma línea afectada
    lin = linexp(yt)
    'Verifica si escapa de la línea mostrada
    If xt > Len(lin) Then CarXY = " ": Exit Function
    'Toma caracter
    CarXY = Mid$(lin, xt, 1)
End Function

Private Function CarPos(p As Tpostex) As String
'Devuelve el caracter actual de una posición. Coordenadas de texto
Dim lin As String
Dim xt As Integer
    lin = linrea(p.yt)    'lee primera línea
    xt = PosXTreal(p.xt, p.yt)
    CarPos = Mid$(lin, xt, 1)
End Function

Private Function CarPosAnt(p As Tpostex) As String
'Devuelve el caracter anterior a una posición. Coordenadas de texto
'Si el caracter es el primero de una línea, se devuelve cadena vacía
Dim lin As String
Dim xt As Integer
    lin = linrea(p.yt)    'lee primera línea
    xt = PosXTreal(p.xt, p.yt)
    If xt <= 1 Then Exit Function    'protección
    CarPosAnt = Mid$(lin, xt - 1, 1)
End Function

Private Function CarPosSig(p As Tpostex) As String
'Devuelve el caracter siguiente a una posición. Coordenadas de texto
'Si el caracter es el último de una línea, se devuelve cadena vacía
Dim lin As String
Dim xt As Integer
    lin = linrea(p.yt)    'lee primera línea
    xt = PosXTreal(p.xt, p.yt)
    If xt >= Len(lin) + 1 Then Exit Function    'protección
    CarPosSig = Mid$(lin, xt + 1, 1)
End Function

Private Function ColXY(xt As Integer, yt As Long) As Long
'Devuelve el color de la posición X,Y en la ventana visible
'Coordenadas del cursor
Dim lin() As TDesSeg
Dim i As Integer
Dim pos As Integer
Dim txt As String
    'Verifica si escapa de la pantalla
    If xt < col1 Or xt > col2 Then Exit Function
    If yt < 1 Or yt > nlin Then Exit Function
    'Lee o actualiza el vector de descripción
    If lincol(yt).tip = TLIN_DES Then
        'Auqnue no es trabajo de ColXY() ver el color
        'Pueda que a veces tenga que hacerlo
        txt = linexp(yt)                'lee toda la línea
        ActualDescColFil txt, lin()
        lincol(yt).tip = TLIN_MIX
        lincol(yt).seg = lin    'copia descripción de línea
    Else    'ya tiene descripción de segmento
        lin = lincol(yt).seg
    End If
    
    'Explora el vector de descripción
    '-------Busca segmento de inicio "i"-------
    i = 1   'Se asume que empieza en el segmento 1
    pos = 1
    For i = 1 To UBound(lin)
        pos = pos + lin(i).tam
        If pos > xt Then
            'Este es el segmento que contiene al caracter.
            ColXY = lin(i).col  'toma color
            Exit Function
        End If
    Next
    'Está al final de la línea, o no hay descripción
    ColXY = vbBlack
End Function

Private Sub BorCursorXY()
'Borra el cursor de la posición (curXt_ant, curYt_ant), redibujando
'el caracter XY (coordenadas de cursor) en la pantalla.
Dim car As String
Dim hdc As Long
Dim posX As Long, posY As Long
    'Verifica si cae fuera de la pantalla
    If curYt_ant < fil1 Or curYt_ant > fil2 Then Exit Sub
    hdc = pic.hdc
    posX = anccarP * (curXt_ant - col1)     '+ ancNLinP
    posY = altcarP * (curYt_ant - fil1)
    'primero dibuja fondo para borrar las marcas anteriores
    FijaRelleno mColFonEdi
    FijaLapiz PS_SOLID, 1, mColFonEdi
    Rectangle hdc, posX, posY, posX + anccarP, posY + altcarP
    'redibuja caracter
    car = CarXY(curXt_ant, curYt_ant)
    pic.ForeColor = ColXY(curXt_ant, curYt_ant)
    SetTextAlign hdc, TA_LEFT
    TextOut hdc, posX, posY, car, Len(car)
    curBorrado = True   'marca bandera
End Sub

Private Sub DibCursorXY()
'Dibuja el símbolo del cursor en la posición (curXt, curYt)
'en la pantalla. Actualiza (curXt_ant, curYt_ant)
Dim car As String
Dim hdc As Long
Dim posX As Long, posY As Long
Dim posXfin As Long, posYfin As Long
    posX = anccarP * (curXt - col1)     '+ ancNLinP
    posY = altcarP * (curYt - fil1)
'    posXfin = posx + (anccarP - 1)
    posYfin = posY + (altcarP - 1)
    hdc = pic.hdc
    If tipSelec = 1 Then        'Selección por columnas
        pic.Line (posX, posY)-(posX + 1, posYfin), vbRed, B
    Else    'Selección normal
        pic.Line (posX, posY)-(posX + 1, posYfin), mColTxtNor, B
    End If
    curBorrado = False   'marca bandera
    'Actualiza posición donde se dibujó el cursor
    curXt_ant = curXt
    curYt_ant = curYt
End Sub

Public Sub XYCursor(x As Long, y As Long)
'Devuelve coordenadas del cursor en Twips.
    x = anccarP * (curXt - col1 + 1) * Screen.TwipsPerPixelX
    y = altcarP * (curYt - fil1 + 1) * Screen.TwipsPerPixelY
End Sub

Private Sub ActivarCursor()
    cursorOn = True
End Sub

Private Sub DesactivarCursor()
'Desactiva el parpadeo del cursor
    'No lo borra, para no interferir con el método de selección
    cursorOn = False
End Sub

Private Sub ApagCursor()
'Apaga el cursor en la posición actual y lo desactiva para parpadeo.
    Call DesactivarCursor
    Call BorCursorXY
End Sub

Public Sub EncenCursor()
'Fuerza a que el cursor aparezca encendido. Activa el parpadeo si es que
'estuviera desactivado
    ActivarCursor       'activa temporización de cursor
    curtmp = 0  'para mantener encendico el cursor por 500mseg
    'borra el cursor de la posición anterior
    Call BorCursorXY
    'Dibuja en la nueva posición
    Call DibCursorXY
End Sub

Private Sub Timer1_Timer()
    'Temporiza el parpadeo del cursor
    If Not cursorOn Then Exit Sub
    curtmp = curtmp + 1     'Lleva la cuenta de 250 mseg
    If curtmp >= 2 Then
        curtmp = 0      'para iniciar otra cuenta de 2 pasos
        If curBorrado Then DibCursorXY Else BorCursorXY
    End If
End Sub

Private Function CursorFueraPantalla() As Boolean
'Verifica si el cursor se encuentra fuera de la pantalla visible
    If curYt < fil1 Or curYt > fil2 Or curXt < col1 Or curXt > col2 Then
        CursorFueraPantalla = True
    End If
End Function

Private Sub FijaCol1(valor As Integer)
'Fija un nuevo valor para "col1" y actualiza "col2". Sólo debería cambiarse
'"col1" desde este procedimiento. Equivale a un desplazamiento horizontal sin
'mover el cursor. 'No redibuja
    'Verificaciones de validez
    If valor < 1 Then valor = 1
    If valor > MaxCol1() Then valor = MaxCol1()
    'Asignación final
    col1 = valor
    col2 = col1 + maxColVis - 1
End Sub

Private Sub ActualizaNLinFin()
'Recalcula la variable "nFilFin".
'Debería llamarse después de cualquier cambio a "fil1" o si se
'eliminan o aumentan lineas visibles en pantalla
    'valor normal de nfilFin
    nfilFin = fil2
    'valor real de nfilFin si hay menos líneas
    If nfilFin > nlin Then nfilFin = nlin
    'Protección por exceso. Limita hasta donde se mueve la pantalla
    If fil1 > nfilFin Then
        fil1 = nfilFin              'Ajusta pantalla
        If fil1 < 1 Then fil1 = 1   'protección por debajo
        'ha habido modificación de fil1, se debe actualizar fil2
        fil2 = fil1 + maxLinVis - 1
    End If
End Sub

Private Sub FijaFil1(valor As Long)
'Fija un nuevo valor para "fil1" y actualiza "fil2". Sólo debería cambiarse
'"fil1" desde este procedimiento Equivale a un desplazamiento vertical sin
'mover el cursor. No redibuja
'Desplaza sólo en variables. No redibuja
    'verifica validez de desplazamiento
    If valor < 1 Then valor = 1     'limita por arriba
    If valor > nlin Then valor = nlin 'limita
    'Asignación final
    fil1 = valor
    fil2 = fil1 + maxLinVis - 1
    Call ActualizaNLinFin    'aquí puede cambiar fil1 y fil2
End Sub

Private Sub FijarSel0()
'Fija el punto base de la selección, en el punto actual del cursor
    sel0 = LeePosCur()
    sel1 = sel0         'Selección sigue al cursor
    sel2 = sel0         'Selección sigue al cursor
    sel1ant = sel0      'actualiza anterior
    sel2ant = sel0      'actualiza anterior
End Sub

Private Function LineaEnSel(yt As Long) As Boolean
'Devuelve Verdadero si la linea "yt" está en el bloque de selección
    If haysel And sel1.yt <= yt And yt <= sel2.yt Then
        LineaEnSel = True
    Else
        LineaEnSel = False
    End If
End Function

'*****************************************************************************
'*************** FUNCIONES PARA MANEJO DE LA MATRIZ DE LÍNEAS ****************
'*****************************************************************************
Public Sub LimpiarLineas()
    nlin = 0
    ReDim linrea(0)
    ReDim lincol(0)
    maxTamLin = 1   'Valor inicial
End Sub

Public Property Get linea(i As Long) As String
'Propiedad para devolver el texto de una línea
    linea = linrea(i)
End Property

Public Property Get nLinActual() As Long
    nLinActual = curYt
End Property

Public Sub EliminarLineas(pos As Long, nelim As Long)
'Elimina líneas en la matriz de cadenas, desplazando los elementos
Dim i As Long
    'protecciones
    If pos > nlin Then Exit Sub     'a eliminar desde más alla del texto
    If nelim > nlin Then Exit Sub   'a eliminar más de lo que hay
    If nelim < 1 Then Exit Sub      'a eliminar negativamente
    'Desplaza elementos
    For i = pos To nlin - nelim
        linrea(i) = linrea(i + nelim)
        lincol(i) = lincol(i + nelim)
    Next
    'Actualiza tamaño
    nlin = nlin - nelim
    ReDim Preserve linrea(nlin)
    ReDim Preserve lincol(nlin)
End Sub

Private Sub InsertarLineas(pos As Long, nins As Long)
'Inserta líneas, desplazando los elementos. Las líneas se insertan en blanco.
Dim i As Long
    'protecciones
    If pos > nlin Then Exit Sub
    If nins < 1 Then Exit Sub
    nlin = nlin + nins
    ReDim Preserve linrea(nlin)
    ReDim Preserve lincol(nlin)
    'Desplaza elementos
    For i = nlin To pos + nins Step -1
        linrea(i) = linrea(i - nins)
        lincol(i) = lincol(i - nins)
    Next
End Sub

Private Sub AgregaLinea(lin As String)
'Agrega una línea al final en el control de texto
Dim tamexp As Integer
    nlin = UBound(linrea) + 1   'actualiza nuevo número de líneas
    ReDim Preserve linrea(nlin) 'crea espacio
    ReDim Preserve lincol(nlin) 'crea espacio
    linrea(nlin) = lin      'escribe la nueva línea
    'actualiza ancho máximo de línea
    tamexp = Len(LineaFin(lin))
    If tamexp > maxTamLin Then maxTamLin = tamexp
    Call ActualizaNLinFin   'Para dibujar correctamente
    Call ActLimitesBarDesp  'actualiza límites de Scroll Bar's
End Sub

Public Sub ReinicColEdi()
'Reinicia todas las líneas para que se evalúe de nuevo el color
Dim i As Long
    For i = 1 To UBound(lincol)
        lincol(i).tip = TLIN_DES
    Next
End Sub
'*****************************************************************************
'***************** FUNCIONES PARA MANEJO DE LAS TABULACIONES *****************
'*****************************************************************************

Private Function linexp(i As Long) As String
'Devuelve la línea expandida como es vísible, reemplazando las tabulaciones
    linexp = LineaFin(linrea(i))      'expande los tabs
End Function

Private Function LineaFin(txt As String) As String
'Reemplaza la cadena "txt" a como se debe mostrar (reemplaza tabulaciones)
Dim tmp As String
Dim lin() As String
Dim nesp As Integer
Dim i As Integer
    lin = Split(txt, vbTab)       'corta por tabulaciónes
    For i = 0 To UBound(lin)
        'completa con espacios
        tmp = tmp & lin(i)           'agrega línea
        If i < UBound(lin) Then 'si no es el último se agrega
            nesp = nEspTab - (Len(lin(i)) Mod nEspTab)  'de 1 a "nEspTab" espacios
            tmp = tmp & String(nesp, " ")   'completa con espacios
        End If
    Next
    LineaFin = tmp
End Function

Private Function posFinTab(x As Integer) As Integer
'Devuelve la posición inicial (en caracteres) del siguiente caracter
'que sigue a un caracter de tabulación ubicado en la posición x
    posFinTab = nEspTab * ((x - 1) \ nEspTab + 1) + 1
End Function

Private Sub ExplorarTabs(yt As Long)
'Explora la línea "yt", y actualiza las matrices:
' * ptabexp() con las posiciones de inicio de las tabulaciones en el texto expandido.
' * ptabrea() con las posiciones de las tabulaciones en el texto real
' Los índices de ptabexp() y ptabrea() empiezan en 1
Dim texp As String      'texto expandido
Dim trea As String      'texto real
Dim lin() As String
Dim nesp As Integer
Dim i As Integer
    lin = Split(linrea(yt), vbTab)             'corta por tabulaciones
    If UBound(lin) = -1 Then
        ReDim ptabexp(0)
        ReDim ptabrea(0)
    Else
        ReDim ptabexp(UBound(lin))
        ReDim ptabrea(UBound(lin))
    End If
    For i = 0 To UBound(lin)
        'completa con espacios
        texp = texp & lin(i)                'actualiza línea
        trea = trea & lin(i) & " "          'actualiza línea
        If i < UBound(lin) Then             'si no es el último se agrega
            nesp = nEspTab - (Len(lin(i)) Mod nEspTab)  'de 1 a "nEspTab" espacios
            ptabexp(i + 1) = Len(texp) + 1  'guarda posición de inicio del tab
            texp = texp & String(nesp, " ") 'completa con espacios
            
            ptabrea(i + 1) = Len(trea)
        End If
    Next
End Sub

Private Function PosXTreal(xt As Integer, yt As Long) As Integer
'Devuelve la posición horizontal real xt (en la cadena sin expandir)
'para la fila "yt" , caracter "xt" en el texto expandido
Dim i As Integer
Dim posIni As Integer
Dim posFin As Integer
Dim posSigCar As Integer
Dim distab As Integer
    '--------------Verifica si hay tabulaciones---------------
    PosXTreal = xt    'valor por defecto
    ExplorarTabs yt   'lee posiciones de tabulación
    If UBound(ptabexp) = 0 Then 'Si no hay tabulaciones
        Exit Function   'sale con la misma posición
    End If
    '----------------------------------------------------------
    'Hay al menos una tabulación. Verifica en que zona
    'de la cadena cae para buscar su posición real
    '----------------------------------------------------------
    'Verifica para la primera zona
    If xt < ptabexp(1) Then
        PosXTreal = xt    'Ocupa la misma posición en la cadena real
        Exit Function
    End If
    'Debe estar en las otras zonas
    For i = 1 To UBound(ptabexp)
        'Toma inicio de zona
        posIni = ptabexp(i)
        'Toma fin de zona
        If i < UBound(ptabexp) Then
            posFin = ptabexp(i + 1)
        Else    'Es el tab final
            posFin = Len(linexp(yt)) + 1
        End If
        posSigCar = posFinTab(posIni)
        'Ver si cae en la zon aprohibida de un "tab" en la cadena expandida
        If xt >= posIni And xt < posSigCar Then
            'Está en la zona de la tabulación expandida
            If xt = posIni Then
                'El cursor está bien ubicado, se inserta antes
                'de la tabulación
                PosXTreal = ptabrea(i)
                Exit Function
            Else
                'El cursor está en una posición prohibida
                'Puede ser que se esté insertando en modo columnas
                msjError = "Error de ubicación de cursor"
                PosXTreal = ptabrea(i)   'por ahora se ajusta a la izquierda
                Exit Function
            End If
        'Verifica si cae en el texto después del tab
        ElseIf xt >= posSigCar And xt < posFin Then
            distab = xt - posSigCar 'distancia a la tabulación
            PosXTreal = ptabrea(i) + distab + 1
            Exit Function
        End If
    Next
    'Sólo debería debería llegar aquí si está al final
    PosXTreal = Len(linrea(yt)) + 1
End Function

Private Function PosXTexp(xt As Integer, yt As Long) As Integer
'Devuelve la posición horizontal xt en la cadena expandida
'para la fila "yt" , caracter "xt" en el texto real
Dim a() As String
Dim lin As String
    '--------------Verifica si hay tabulaciones---------------
    PosXTexp = xt    'valor por defecto
    If xt <= 1 Then Exit Function   'No puede ser de otra forma
    lin = Left$(linrea(yt), xt - 1) 'lee parte afectada de la línea
    'La posición expandida equivale al largo de la cadena
    'afectada expandida.
    PosXTexp = Len(LineaFin(lin)) + 1
End Function

Public Sub TabToSpaces()
'Convierte tabulaciones a espacios en el editor.
Dim i As Long
    For i = 1 To nlin
        linrea(i) = linexp(i)
    Next
    Call InicDeshacer   'Porque esta acción no se puede deshacer
End Sub

'*****************************************************************************
'****************** FUNCIONES PARA DESPLAZAMIENTO DEL CURSOR *****************
'*****************************************************************************

Private Function curSigPal() As Long
'Devuelve la posición del cursor "xt" (en coordenadas de texto) de la siguiente
'palabra en la línea actual
Dim x As Long
    x = curXt
    If x < 1 Then Exit Function
    'Termina la palabra actual
    While x <= Len(linact) And Mid$(linact, x, 1) Like CAR_DESP_PAL
        x = x + 1
    Wend
    'busca siguiente
    While x <= Len(linact) And Not (Mid$(linact, x, 1) Like CAR_DESP_PAL)
        x = x + 1
    Wend
    curSigPal = x
End Function

Private Function curIniPal() As Long
'Devuelve la posición del cursor "xt" de la
'anterior palabra en la línea actual
Dim x As Long
    If curXt <= 1 Then Exit Function
    x = curXt
    'busca fin de palabra
    While x > 1 And Not (Mid$(linact, x, 1) Like CAR_DESP_PAL)
        x = x - 1
    Wend
    If x = curXt Then   'ya estaba al inicio de palabra
        x = x - 1   'retrocede
        If x = 0 Then curIniPal = x: Exit Function
        'Busca fin de palabra anterior
        While x > 1 And Not (Mid$(linact, x, 1) Like CAR_DESP_PAL)
            x = x - 1
        Wend
    End If
    'Busca inicio de palabra actual
    While x > 1 And Mid$(linact, x, 1) Like CAR_DESP_PAL
        x = x - 1
    Wend
    If x = 1 Then curIniPal = 1 Else curIniPal = x + 1
End Function

Private Function curIniPal2() As Long
'Devuelve la posición del cursor "xt" del
'inicio de palabra en la posición actual del cursor
Dim x As Long
    If curXt < 1 Then Exit Function
    x = curXt
    'busca fin de palabra
    If Mid$(linact, x, 1) Like CAR_DESP_PAL Then
        'ya está en medio de palabra
    Else
        'Hay que buscar el inicio
        If x = 1 Then Exit Function 'es el primero
        x = x - 1   'retrocede
        'Busca fin de palabra anterior
        While x > 1 And Not (Mid$(linact, x, 1) Like CAR_DESP_PAL)
            x = x - 1
        Wend
    End If
    'Busca inicio de palabra actual
    While x > 1 And Mid$(linact, x, 1) Like CAR_DESP_PAL
        x = x - 1
    Wend
    If x = 1 Then curIniPal2 = 1 Else curIniPal2 = x + 1
End Function

Private Function curFinPal() As Long
'Devuelve la posición del cursor "xt" del caracter que sigue al
'fin del identificador que se encuentra bajo el cursor
Dim x As Long
    If curXt < 1 Then Exit Function
    x = curXt
    'Busca fin de palabra actual
    While x < Len(linact) + 1 And Mid$(linact, x, 1) Like CAR_DESP_PAL
        x = x + 1
    Wend
    curFinPal = x   'funciona inclusive en el caso límite
End Function

Private Function InicioParrafo(yt As Long, xt As Integer) As Boolean
'Indica si una línea está al inicio de un párrafo para un valor de xt
    If Len(linexp(yt)) >= xt And Len(linexp(yt - 1)) < xt Then
        InicioParrafo = True
    Else
        InicioParrafo = False
    End If
End Function

Private Function FinalParrafo(yt As Long, xt As Integer) As Boolean
'Indica si una línea está al final de un párrafo para un valor de xt
    If yt = nlin Then   'No hay línea siguiente
        If Len(linexp(yt)) >= xt Then
            FinalParrafo = True
        Else
            FinalParrafo = False
        End If
    Else                'Hay línea siguiente
        If Len(linexp(yt)) >= xt And Len(linexp(yt + 1)) < xt Then
            FinalParrafo = True
        Else
            FinalParrafo = False
        End If
    End If
End Function

Private Function EntreParrafos(yt As Long, xt As Integer) As Boolean
'Indica si una línea está entre dos párrafos para un valor de xt
    If Len(linexp(yt)) < xt Then
        EntreParrafos = True
    Else
        EntreParrafos = False
    End If
End Function

Private Function curIniPar() As Long
'Devuelve la posición del cursor "yT" (en coordenadas del texto) del inicio
'del parrafo en la posición actual del cursor o del fin del párrafo anterior
Dim y As Long
    If curYt = 1 Then Exit Function
    y = curYt
    'Verifica si está al inicio de un párrafo
    If InicioParrafo(y, curXt) Then
        'Buscará el fin del párrafo anterior
        y = y - 1
        'Busca fin de párrafo anterior
        While y > 1 And Len(linexp(y)) < curXt
            y = y - 1
        Wend
        curIniPar = y
        Exit Function   'sale
    End If
    'Si está entre párrafos se mueve hasta el fin del anterior
    If EntreParrafos(y, curXt) Then
        While y > 1 And Len(linexp(y)) < curXt
            y = y - 1
        Wend
        curIniPar = y
        Exit Function   'sale
    End If
    'Busca inicio de párrafo actual
    While y > 1 And Len(linexp(y)) >= curXt
        y = y - 1
    Wend
    If y = 1 Then        'llegó al inicio
        curIniPar = y
    Else    'encontró límite
        y = y + 1
        curIniPar = y
    End If
End Function

Private Function curFinPar() As Long
'Devuelve la posición del cursor "yt" (en coordenadas del texto) del fin
'del parrafo en la posición actual del cursor o del inicio del párrafo siguiente
Dim y As Long
    If curYt = nlin Then curFinPar = curYt: Exit Function
    y = curYt
    'Verifica si está al final de un párrafo
    If FinalParrafo(y, curXt) Then
        'Buscará el inicio del párrafo siguiente
        y = y + 1
        'Busca fin de párrafo anterior
        While y < nlin And Len(linexp(y)) < curXt
            y = y + 1
        Wend
        curFinPar = y
        Exit Function   'sale
    End If
    'Si está entre párrafos se mueve hasta el inicio del siguiente
    If EntreParrafos(y, curXt) Then
        While y < nlin And Len(linexp(y)) < curXt
            y = y + 1
        Wend
        curFinPar = y
        Exit Function   'sale
    End If
    'Busca fin de párrafo actual
    While y < nlin And Len(linexp(y)) >= curXt
        y = y + 1
    Wend
    If Not FinalParrafo(y, curXt) Then        'al inicio
        y = y - 1
    End If
    curFinPar = y
End Function

Public Property Let Text(ByRef txt As String)
'Actualiza el contenido del CONTROL
Dim a() As String
Dim i As Integer
    'inicia contenido en 0 líneas
    Call LimpiarLineas
    If txt <> "" Then
        'asignar texto
        a = Split(txt, vbCrLf)
        For i = 0 To UBound(a)
            AgregaLinea a(i)
        Next
    End If
    Call ActualizaNLinFin   'actualiza líneas visibles
    Call ActLimitesBarDesp  'actualiza límites de Scroll Bar's
    'posición inicial del cursor
    tCursorA2 maxTamLin, nlin
    Call Dibujar
End Property

Public Property Get Text() As String
'Devuelve el contenido del control
'Devuelve el texto seleccionado.
Dim tip0 As Integer
    tip0 = tipSelec 'guarda tipo de seleccion
    tipSelec = 0    'pone en modo normal para copiar todo
    Text = TextBlo(MinPos, MaxPos)
    tipSelec = tip0 'Restaura tipo de seleccion
End Property

'*****************************************************************************
'******************** FUNCIONES PARA MANEJO DE BLOQUES ***********************
'*****************************************************************************
Private Function LeePosCur() As Tpostex
'Lee la posición actual de cursor en la variable de posición pos
    LeePosCur.xt = curXt
    LeePosCur.yt = curYt
End Function

Private Function PosNulo(p As Tpostex) As Boolean
'Indica si una posición no ha sido iniciada
    PosNulo = (p.xt = 0)
End Function

Private Function MinPos() As Tpostex
'Devuelve la posición menor del texto en el editor
    If nlin > 0 Then
        MinPos.xt = 1
        MinPos.yt = 1
    End If
End Function

Private Function MaxPos() As Tpostex
'Devuelve la posición mayor del texto en el editor
    MaxPos.yt = nlin
    If nlin > 0 Then
        MaxPos.xt = Len(linexp(nlin)) + 1
    End If
End Function

Private Function TextPosIni(p1 As Tpostex) As String
'Devuelve la primera línea del bloque que empieza en la posición p1.
'Sólo debe usarse cuando el bloque tiene más de una línea
'Devuelve texto sin expandir
Dim x1r As Integer, y1 As Long
    y1 = p1.yt
    x1r = PosXTreal(p1.xt, y1)
    TextPosIni = Mid$(linrea(y1), x1r)       'copia primera línea
End Function

Private Function TextPosFin(p2 As Tpostex) As String
'Devuelve la última línea del bloque que termina en la posición p2.
'Sólo debe usarse cuando el bloque tiene más de una línea
'Devuelve texto sin expandir
Dim x2r As Integer, y2 As Long
    y2 = p2.yt
    x2r = PosXTreal(p2.xt, y2)
    TextPosFin = Mid$(linrea(y2), 1, x2r - 1)  'copia última línea
End Function

Private Function TextPosLin(p1 As Tpostex, p2 As Tpostex) As String
'Devuelve el texto seleccionado del bloque p1-p2.
'Sólo debe usarse cuando el bloque está en una sola línea.
'Devuelve texto sin expandir
Dim y1 As Long
Dim x1r As Integer, x2r As Integer
    y1 = p1.yt      'igual a p2.yt
    x1r = PosXTreal(p1.xt, y1)
    x2r = PosXTreal(p2.xt, y1)
    TextPosLin = Mid$(linrea(y1), x1r, x2r - x1r)  'copia línea
End Function

Private Function TextPosCol(yt As Long, xt1 As Integer, xt2 As Integer) As String
'Devuelve el texto intermedio de una línea en el modo columna
'Para extraer el texto intermedio, se expanden primero las tabulaciones
'y se completan con expacios si la línea es muy pequeña.
'xt1 DEBE SER SIEMPRE menor o igual a xt2
Dim x1 As Integer
Dim x2 As Integer
Dim txt As String
    txt = linexp(yt)    'expande las tabulaciones
    If Len(txt) < xt1 Then   'línea muy pequeña
        TextPosCol = String(xt2 - xt1, " ")  'devuelve espacios
    ElseIf Len(txt) < xt2 - 1 Then 'sólo se ve parcialmente
        TextPosCol = Mid$(txt & String(xt2 - 1 - Len(txt), " "), xt1)
    Else    'La línea es suficientemente grande
        TextPosCol = Mid$(txt, xt1, xt2 - xt1)
    End If
End Function

Private Function TamPosBlo(p1 As Tpostex, p2 As Tpostex) As Long
'Devuelve el tamaño de un bloque en bytes. Es eficiente aún con bloques grandes
'Toma en cuenta el tipo de selección actual
'Para el caso de selección por columnas, se considera el texto expandido.
'Para el caso de selección normal se da el tamaño del texto sin expandir
Dim tam As Long
Dim i As Long
    If UBound(linrea) = 0 Then tam = 0: Exit Function
    If tipSelec = 1 Then        'Selección por columnas
        'Para el cálculo, se fija un tamaño uniforme de selección
        tam = Abs(p1.xt - p2.xt) + 2  'caracteres por línea más el salto de línea
        i = (p2.yt - p1.yt + 1)   'número de líneas
        TamPosBlo = tam * i - 2        'quita el salto final
        Exit Function
    End If
    'Selección por filas
    If p1.yt = p2.yt Then   'Bloque en una sola línea
        tam = Len(TextPosLin(p1, p2))
    Else                    'Bloque en varias líneas
        tam = Len(TextPosIni(p1))            'tamaño primera línea
        For i = p1.yt + 1 To p2.yt - 1
            tam = tam + 2 + Len(linrea(i))   'considera 2 caracteres del salto
        Next
        tam = tam + 2 + Len(TextPosFin(p2))  'tamaño última línea
    End If
    TamPosBlo = tam   'Devuelve
End Function

Private Function BloAMem(tam As Long, p1 As Tpostex, p2 As Tpostex) As Long
'Funcion estrella del programa. Se usa para las opciones del portapapeles y
'para la funcionalidad de "deshacer".
'Hace un volcado rápido de un bloque de texto a Memoria.
'Puede trabajar éficientemente con varios miles de líneas.
'Si hubo error en asignar memoria devuelve 0, de otra forma devuelve el
'manejador del bloque de memoria.
'Actualiza la variable "tam" con el tamaño del bloque reservado
Dim nbytes As Long  'Número de bytes escritos
Dim f As Long
Dim lpMemory As Long
Dim retval As Long
Dim hData As Long   'Manejador de memoria
Dim cad As String   'Cadena a escribir
Dim xt1 As Integer, xt2 As Integer
Dim xtmp As Integer
    '-----------Copia datos a memoria - Metodo1----------
    'Calcula el tamaño total del bloque
    tam = TamPosBlo(p1, p2)
    'Asigna espacio en memoria. Es tam+1 porque se le agregará un NULL a la cadena
    hData = GlobalAlloc(GMEM_MOVEABLE Or GMEM_DDESHARE, tam + 1)
    If hData = 0 Then   'No se pudo encontrar memoria
        BloAMem = 0
        Exit Function
    End If
    'Copiamos la cadena al espacio de memoria reservada
    lpMemory = GlobalLock(hData)    'bloquea mientras copia y obtenemos dirección
    If tipSelec = 1 Then        'Selección por columnas
        'Lee coordenadas horizontales
        xt1 = p1.xt: xt2 = p2.xt
        If xt1 > xt2 Then   'Verifica si hay que invertir
            xtmp = xt1: xt1 = xt2: xt2 = xtmp
        End If
        cad = TextPosCol(p1.yt, xt1, xt2) 'Lee parte inicial del bloque
        nbytes = Len(cad)       'tamaño sin incluir el NULL final
        retval = lstrcpy(lpMemory, cad)     'Copia incluyendo el NULL final
        lpMemory = lpMemory + nbytes        'apunta al NULL final escrito
        For f = p1.yt + 1 To p2.yt
            cad = vbCrLf & TextPosCol(f, xt1, xt2)
            nbytes = Len(cad)       'tamaño sin incluir el NULL final
            retval = lstrcpy(lpMemory, cad) 'Copia incluyendo el NULL final
            lpMemory = lpMemory + nbytes    'apunta al NULL final escrito
        Next
    Else    'Selección normal
        If p1.yt = p2.yt Then   'Selección de una sola línea
            'Copia incluyendo el NULL final
            retval = lstrcpy(lpMemory, TextPosLin(p1, p2))
        Else
            'Selección de varias líneas
            cad = TextPosIni(p1)        'Lee parte inicial del bloque
            nbytes = Len(cad)           'tamaño sin incluir el NULL final
            retval = lstrcpy(lpMemory, cad)     'Copia incluyendo el NULL final
            lpMemory = lpMemory + nbytes        'apunta al NULL final escrito
            For f = p1.yt + 1 To p2.yt - 1
                cad = vbCrLf & linrea(f)
                nbytes = Len(cad)       'tamaño sin incluir el NULL final
                retval = lstrcpy(lpMemory, cad) 'Copia incluyendo el NULL final
                lpMemory = lpMemory + nbytes    'apunta al NULL final escrito
            Next
            cad = vbCrLf & TextPosFin(p2)       'Lee parte final del bloque
            retval = lstrcpy(lpMemory, cad)     'Copia incluyendo el NULL final
        End If
    End If
    Call GlobalUnlock(hData)            'desbloquea
    BloAMem = hData     'Devuelve manejador
End Function

Private Property Get TextBlo(p1 As Tpostex, p2 As Tpostex) As String
'Devuelve el texto de un bloque.
Dim hData As Long   'Manejador de memoria
Dim lpMemory As Long
Dim nbytes As Long
    'Protección
    If p1.xt = p2.xt And p1.yt = p2.yt Then Exit Sub
    'Copia bloque a memoria
    hData = BloAMem(nbytes, p1, p2) 'Copia selección a memoria
    If hData = 0 Then
        MsgBox "No se puede obtener selección. Error asignando memoria", vbCritical
        Exit Property
    End If
    TextBlo = Space(nbytes + 1)     'crea espacio para contener a la cadena y el NULL
    lpMemory = GlobalLock(hData)    'bloquea mientras copia y obtenemos dirección
    lstrcpy TextBlo, lpMemory        'Copia rápidamente a cadena
    Call GlobalUnlock(hData)        'desbloquea
    TextBlo = Left(TextBlo, Len(TextBlo) - 1)    'Quita NULL final
    GlobalFree hData
    Exit Property

'No se usa el método:
'        For i = .yt To  .yt
'            tmp = tmp & vbCrLf & linrea(i)
'        Next
'Porque es muy lento cuando hay muchas líneas
End Property

Public Property Get TextSel() As String
'Devuelve el texto seleccionado.
    If Not haysel Then Exit Property
    TextSel = TextBlo(sel1, sel2)
End Property

'*****************************************************************************
'******************** FUNCIONES DE POSICIONAMIENTO DE CURSOR *****************
'*****************************************************************************

Private Sub AjustaPantalla()
'Ajusta las coordenadas de la pantalla para que sea visible el cursor
'No modifica nada en el caso que el cursor sea visible.
'Actualiza las variables globales "curXt", "curYt"
    'Ubica horizontal
    If curXt < col1 Then
        FijaCol1 curXt
        curXt = col1
    End If
    If curXt > col2 Then
        FijaCol1 curXt - maxColVis + 1
        curXt = col2
    End If
    'Ubica vertical
    If curYt < fil1 Then
        FijaFil1 curYt
        curYt = fil1
    End If
    If curYt > fil2 Then
        FijaFil1 curYt - maxLinVis + 1   'desplaza
        curYt = fil2    'debe actualizarse también
    End If
End Sub

Private Sub FijaCursor(ByVal xt As Integer, ByVal yt As Long, _
            Optional ajus_hor As Integer = A_NULO)
'Fija el cursor en la posición xc, yc.
'Sólo permite poner el cursor en una zona válida del texto (aunque no sea
'visible en la pantalla actual). Si cae fuera del texto lo ajusta a caer dentro.
'No hace desplazamiento de pantalla.
'Actualiza las variables globales "curXt", "curYt" y "linact"
''"ajus_hor" es el ajuste que se desea realizar, puede ser:
'A_NULO ->     Sin ajuste. Puede aparecer en medio de la zona del tab
'A_IZQ_TAB ->  A izquierda cuando hay tabulación
'A_DER_TAB ->  A derecha cuando hay tabulación
Dim i As Integer
Dim posSigCar As Integer
    '-----Verifica si hay líneas en el control
    If nlin = 0 Then    'no hace nada, sólo posiciona
        curXt = 1: curYt = 1
        linact = ""
        Exit Sub
    End If
    '-----valida que caiga en una zona válida del texto
    'calcula y valida "curYt"
    curYt = yt                     'valor inicial
    If curYt < 1 Then curYt = 1    'valida por abajo
    If curYt > nlin Then curYt = nlin  'valida por arriba
    linact = linexp(curYt)          'actualiza línea actual
    If Len(linact) > maxTamLin Then
        maxTamLin = Len(linact)     'es la línea mayor
        Call ActLimitesBarDesp      'actualiza barra
    End If
    'calcula y valida "curXt"
    If CLng(xt) > 32767 Then xt = col1  'protección
    curXt = xt                      'valor inicial
    If curXt < 1 Then curXt = 1     'valida por abajo
    If curXt > Len(linact) + 1 Then curXt = Len(linact) + 1 'valida por arriba
    '----Realiza los ajustes si se han pedido.
    If ajus_hor <> A_NULO Then  'Hay que realizar ajustes
        'Verifica si cae en posición prohibida por haber tabulación
        ExplorarTabs curYt  'lee posiciones de tabulación
        For i = 1 To UBound(ptabexp)
            'Verifica si cae en la zona prohibida definida por un "tab"
            posSigCar = posFinTab(ptabexp(i))
            If curXt > ptabexp(i) And curXt < posSigCar Then
                If ajus_hor = A_DER_TAB Then
                    curXt = posSigCar   'actualiza "curXt"
                End If
                If ajus_hor = A_IZQ_TAB Then
                    curXt = ptabexp(i)    'actualiza "curXt"
                End If
            End If
        Next
    End If
End Sub

Private Sub tCursorA(xt As Integer, yt As Long, Optional ajus_hor As Integer = A_IZQ_TAB)
'Función similar a FijaCursor(), pero realiza el desplazamiento de
'la pantalla cuando sea necesario para que el cursor siempre aparezca visible
'Actualiza la bandera "Redibujar", cuando se requiere dibujar toda la pantalla.
    FijaCursor xt, yt, ajus_hor      'mueve el cursor
    'verifica si es necesario refrescar la pantalla
    If CursorFueraPantalla Then
        Call AjustaPantalla
        Redibujar = True    'pide que se redibuje
    End If
End Sub

Private Sub tCursorA2(xt As Integer, yt As Long, Optional ajus_hor As Integer = A_IZQ_TAB)
'Versión de tCursorA() que verifica si el cursor queda en una posición
'horizontal límite y encuadra la pantalla para que se vea mejor.
'Actualiza la bandera "Redibujar", cuando se requiere dibujar toda la pantalla.
    FijaCursor xt, yt, ajus_hor
    'Verifica si el cursor cae en una zona no visible de la pantalla
    If CursorFueraPantalla Then
        AjustaPantalla
        'Verifica si quedó al final de línea
        If curXt = col2 Then
            FijaCol1 col1 + 12
            AjustaPantalla   'porque puede que se haya ocultado el cursor
                             'y para actualizar curXt
        End If
        'Verifica si cae al inicio
        If curXt = col1 Then
            FijaCol1 col1 - 12  'Intenta desplazar
            AjustaPantalla   'porque puede que se haya ocultado el cursor
                             'y para actualizar curXt
        End If
        Redibujar = True
    End If
End Sub

Private Sub posCursorA(pos As Tpostex, Optional ajus_hor As Integer = A_IZQ_TAB)
'Versión de tCursorA(), pero acepta una posición de tipo "Tpostex"
    tCursorA pos.xt, pos.yt, ajus_hor
End Sub

Private Sub posCursorA2(pos As Tpostex, Optional ajus_hor As Integer = A_IZQ_TAB)
'Versión de tCursorA2(), pero acepta una posición de tipo "Tpostex"
    tCursorA2 pos.xt, pos.yt, ajus_hor
End Sub

Public Sub CursorText(xt As Integer, yt As Long)
'Función pública para fijar la posición del cursor
    Redibujar = False
    tCursorA2 maxTamLin + 1, nlin, A_IZQ_TAB    'para encuadrar desde arriba
    Call tCursorA2(xt, yt)
    curXd = curXt     'actualiza posición deseada
    Call LimpSelec(True)      'Limpia selección
    'Marca la posición del cursor antes de una selección
    Call FijarSel0      'Fija punto base
    Call EncenCursor    'Enciende para que sea visible en la nueva posición
    If Redibujar Then Call Refrescar
End Sub

Public Sub SelectHasta(xt As Integer, yt As Long)
'Función pública para fijar el fin de una selección. Se debe habe definido el inicio
'con CursorText
    tCursorA2 xt, yt
    Call ExtenderSel 'Extiende selección
End Sub
'*****************************************************************************
'******************** FUNCIONES DE DIBUJO DEL TEXTO  *****************
'*****************************************************************************

Private Sub LeerXPosSel(x1 As Long, x2 As Long)
'Devuelve las coordenadas horizontales x1 y x2 del bloque de selección
'a mostrar en pantalla. Devuelve en de pixeles
    'Calcula posición horizontal de inicio y fin de bloque
    x1 = (sel1.xt - col1) * anccarP
    If x1 < 0 Then x1 = 0
    If x1 > ScaleWidth Then x1 = ScaleWidth
    x2 = (sel2.xt - col1) * anccarP - 1   'incluye sólo el recuadro del caracter no el siguiente
    If x2 < 0 Then x2 = 0
    If x2 > ScaleWidth Then x2 = ScaleWidth
End Sub

Private Sub LeerXCurSel(xc1 As Integer, xc2 As Integer)
'Devuelve las coordenadas horizontales xc1 y xc2 del bloque de selección
'a mostrar en pantalla. Devuelve en coordenadas de Cursor.
'xc1 es la posición del caracter en la línea donde se inicia la selección y
'xc2 es la posición del caracter en la línea donde termina la selección.
    'Calcula posición horizontal de inicio y fin de bloque
    xc1 = sel1.xt - col1 + 1
    If xc1 < 1 Then xc1 = 1
    If xc1 > maxColVis Then
        xc1 = maxColVis
    End If
    xc2 = sel2.xt - col1 + 1  'incluye sólo el recuadro del caracter no el siguiente
    If xc2 < 1 Then xc2 = 1
    If xc2 > maxColVis Then
        xc2 = maxColVis
    End If
End Sub

Private Sub DibBloFonC(yc As Long, xc1 As Integer, xc2 As Integer, col_fon As Long)
'Dibuja un recuadro relleno con color "col_fon". Las coordenadas son de cursor
Dim xp1 As Long, xp2 As Long
Dim yp As Long
    'Valida parámetros
'    If xc2 < 1 Then Exit Sub
'    If xc1 < maxColVis Then Exit Sub
    If xc1 < 1 Then xc1 = 1
    If xc2 > maxColVis + 1 Then xc2 = maxColVis + 1
    'Convierte a pixeles
    xp1 = (xc1 - 1) * anccarP '+ ancNLinP
    xp2 = (xc2 - 1) * anccarP '+ ancNLinP
    yp = (yc - 1) * altcarP 'actualiza coordenada vertical
    FijaRelleno col_fon
    FijaLapiz PS_SOLID, 1, col_fon
    If xp2 = xp1 Then xp2 = xp2 + 2     'para que sea visible la selección de 0 columnas
    Rectangle pic.hdc, xp1, yp, xp2, yp + altcarP
End Sub

Private Sub cierraSegm(lin() As TDesSeg, i As Integer, _
                       pfin As Long, col As Long, tip As Integer)
'Cierra un segmento de texto y actualiza "pfin".
'"n" es el tamaño actual de la matriz. El segmento a cerrar llega hasta "i-1"
'"pfin" es la posición de fin del bloque anterior
Dim n As Integer
    If i > pfin + 1 Then  'hay segmento antes?
        n = UBound(lin) + 1
        ReDim Preserve lin(n)
        lin(n).tam = i - pfin - 1
        pfin = i - 1  'actualiza el limite
        lin(n).col = col    'texto normal
        lin(n).tip = tip    'asigna tipo
    End If
End Sub

Private Sub cierraSegm1(lin() As TDesSeg, i As Integer, _
                        pfin As Long, col As Long, tip As Integer)
'Similar a cierraSegm() pero considera que el segmento llega hasta "i"
Dim n As Integer
    If i > pfin + 1 Then   'hay segmento antes?
        n = UBound(lin) + 1
        ReDim Preserve lin(n)
        lin(n).tam = i - pfin
        pfin = i   'actualiza el limite
        lin(n).col = col    'color de texto
        lin(n).tip = tip    'asigna tipo
    End If
End Sub

Private Function cierraSegmId(lin() As TDesSeg, i As Integer, _
                       pfin As Long, col As Long, tip As Integer, _
                       tmp As String, iden As String, largo As Integer) As Boolean
'Cierra un segmento y crea uno nuevo, basado siempre que se encuentre
'un identificador en la cadena "tmp" a partir de la posición "i".
    If Mid$(tmp, i, largo) = iden And Not (Mid$(tmp, i + largo, 1) Like CAR_IDEN_VALM) Then
        cierraSegm lin(), i, pfin, col, tip   'cierra anterior
        i = i + largo - 1   'adelanta el índice, se adelantará 1 en el "for"
        'crea segmento de palabra reservada
        cierraSegm1 lin(), i, pfin, mColPalRes, TSEG_PRS
        
        cierraSegmId = True 'indica que se encontró segmento
    Else
        cierraSegmId = False
    End If
End Function

Private Function cierraSegmId2(lin() As TDesSeg, i As Integer, _
                       pfin As Long, col As Long, tip As Integer, _
                       tmp As String, iden As String, largo As Integer) As Boolean
'Igual a "cierraSegmId", pero para el segundo grupo de palabras reservadas
    If Mid$(tmp, i, largo) = iden And Not (Mid$(tmp, i + largo, 1) Like CAR_IDEN_VALM) Then
        cierraSegm lin(), i, pfin, col, tip   'cierra anterior
        i = i + largo - 1   'adelanta el índice, se adelantará 1 en el "for"
        'crea segmento de palabra reservada
        cierraSegm1 lin(), i, pfin, mColPalRes2, TSEG_PRS2
        
        cierraSegmId2 = True 'indica que se encontró segmento
    Else
        cierraSegmId2 = False
    End If
End Function

Private Sub AnalizarIdentificador(c As String, tmp As String, _
                        lin() As TDesSeg, i As Integer, _
                        pfin As Long, colseg As Long, tipseg As Integer)
'Analiza los identificadores en busca de palabras reservadas
    Select Case c
    Case "A"
        If cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "ALUMINIO", 8) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "ANTIMONIO", 9) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "AMERICIO", 8) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "AZUFRE", 6) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        End If
    Case "B"
        If cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "BARIO", 5) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "BERKELIO", 8) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "BERILIO", 7) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "BORO", 4) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        End If
    Case "C"
        If cierraSegmId2(lin(), i, pfin, colseg, tipseg, tmp, "CADMIO", 6) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "CALCIO", 6) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "CARBONO", 7) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "CLORO", 5) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        End If
    Case "E"
        If cierraSegmId2(lin(), i, pfin, colseg, tipseg, tmp, "ERBIO", 5) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId2(lin(), i, pfin, colseg, tipseg, tmp, "ESTAÑO", 6) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        End If
    Case "F"
        If cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "FLUOR", 5) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "FÓSFORO", 7) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        End If
    Case "H"
        If cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "HELIO", 5) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "HIDRÓGENO", 9) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "HIERRO", 6) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        End If
    Case "I"
        If cierraSegmId2(lin(), i, pfin, colseg, tipseg, tmp, "INDIO", 5) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId2(lin(), i, pfin, colseg, tipseg, tmp, "ITERBIO", 7) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId2(lin(), i, pfin, colseg, tipseg, tmp, "ITRIO", 5) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        End If
    Case "L"
        If cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "LITIO", 5) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "LUTECIO", 7) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        End If
    Case "M"
        If cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "MAGNESIO", 8) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "MERCURIO", 8) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId2(lin(), i, pfin, colseg, tipseg, tmp, "MOLIBDENO", 9) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        End If
    Case "N"
        If cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "NEÓN", 4) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "NITRÓGENO", 9) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        End If
    Case "O"
        If cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "ORO", 3) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "OXÍGENO", 7) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        End If
    Case "P"
        If cierraSegmId2(lin(), i, pfin, colseg, tipseg, tmp, "PALADIO", 7) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId2(lin(), i, pfin, colseg, tipseg, tmp, "PLATA", 5) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "PLATINO", 7) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "PLOMO", 5) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "PLUTONIO", 8) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        End If
    Case "R"
        If cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "RADIO", 5) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "RADÓN", 5) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        End If
    Case "S"
        If cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "SELENIO", 7) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "SILICIO", 7) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "SODIO", 5) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        End If
    Case "T"
        If cierraSegmId2(lin(), i, pfin, colseg, tipseg, tmp, "TANTALIO", 8) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId2(lin(), i, pfin, colseg, tipseg, tmp, "TECNECIO", 8) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "TITANIO", 7) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "TUNGSTENO", 9) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        End If
    Case "U"
        If cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "URANIO", 6) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        End If
    Case "V"
        If cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "VANADIO", 7) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        End If
    Case "X"
        If cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "XENÓN", 5) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        End If
    Case "Z"
        If cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "ZINC", 4) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        ElseIf cierraSegmId(lin(), i, pfin, colseg, tipseg, tmp, "ZIRCONIO", 8) Then
            tipseg = TSEG_DES    'prepara tipo siguiente
        End If
    End Select

End Sub

Private Sub ActualDescColFil(txt As String, lin() As TDesSeg)
'Actualiza el vector descriptor de segmentos de una fila.
'Sirve para facilitar el análisis semántico y el coloreado de sintaxis.
'Sólo soporta filas menores a 32767 caracteres
Dim pfin As Long  'posición de fin de bloque
Dim i As Integer, n As Integer
Dim colseg As Long      'color de segmento
Dim tipseg As Integer   'tipo de segmento
Dim buscafincad1 As Boolean
Dim buscafincad2 As Boolean
Dim c As String, tmp As String
Dim IniIden As Boolean  'bandera de inicio de identificador
    'n = 0   'número de elementos de la matriz
    ReDim lin(0)
    If txt = "" Then Exit Sub
    colseg = mColTxtNor    'inicia color de segmento
    tipseg = TSEG_NOR
    pfin = 0    'Posición del caracter final del bloque
    tmp = UCase(txt)    'La comparación se hace ignorando la caja
    
    IniIden = True      'El inicio de la línea es inicio de identificador
    For i = 1 To Len(txt)
        c = Mid$(tmp, i, 1)    'caracter i
        If buscafincad1 Then   'Estamos dentro de una constante de cadena
            If c = """" Then
                'se toma hasta i para incluir las comillas
                cierraSegm1 lin(), i, pfin, mColTxtCad, TSEG_CAD   'cierra
                buscafincad1 = False     'Marca fin de constante cadena
            End If
        ElseIf buscafincad2 Then   'Estamos dentro de una constante de cadena
            If c = "'" Then
                'se toma hasta i para incluir las comillas
                cierraSegm1 lin(), i, pfin, mColTxtCad, TSEG_CAD   'cierra
                buscafincad2 = False     'Marca fin de constante cadena
            End If
        ElseIf Mid$(tmp, i, 2) = "--" Then  'Busca comentario
            cierraSegm lin(), i, pfin, colseg, tipseg     'cierra anterior
            'crea segmento final
            n = UBound(lin) + 1
            ReDim Preserve lin(n)
            lin(n).tam = Len(txt) - pfin
            lin(n).col = mColTxtCom    'comentario
            lin(n).tip = TSEG_COM
            Exit Sub    'Ya no hay más segmentos después de este
        ElseIf c = """" Then      'Busca comilla
            cierraSegm lin(), i, pfin, colseg, tipseg   'cierra anterior
            buscafincad1 = True  'marca bandera
        ElseIf c = "'" Then      'Busca comilla
            cierraSegm lin(), i, pfin, colseg, tipseg  'cierra anterior
            buscafincad2 = True   'marca bandera
        'busca palabras reservadas
        ElseIf IniIden Then     'Es el inicio de identificador (palabras reservadas)
            AnalizarIdentificador c, tmp, lin(), i, pfin, colseg, tipseg
        End If
        'Valida estado de "IniIden" para el siguiente caracter
        If c Like CAR_IDEN_VALM Then IniIden = False Else IniIden = True
    Next
    'Terminó la exploración, cierra segmento
    If buscafincad1 Then
        'Cadena sin delimitador final
        cierraSegm1 lin(), i, pfin, mColTxtCad, TSEG_NOR     'cierra
    ElseIf buscafincad2 Then
        'Cadena sin delimitador final
        cierraSegm1 lin(), i, pfin, mColTxtCad, TSEG_NOR     'cierra
    Else
        cierraSegm lin(), i, pfin, colseg, tipseg
    End If
End Sub

Private Sub DibLinSel(y As Long, yc As Long, yt As Long, linvis As String)
'Dibuja una línea de texto en modo BN.
'La línea a dibujar debe estár en el bloque se selección
Dim xc1 As Integer, xc2 As Integer
Dim tmp As String
    'Dibuja el fondo con la selección
    LeerXCurSel xc1, xc2    'obtiene coordenadas de selección
    If tipSelec = 1 Then    '--------Selección por columnas
        'Dibuja el texto antes del bloque
        pic.ForeColor = mColTxtNor
        tmp = Left$(linvis, xc1 - 1)    'protege de líneas pequeñas
        TextOut pic.hdc, 0, y, tmp, Len(tmp)
        'Dibuja el texto después del bloque
        tmp = Mid$(linvis, xc2)         'protege de líneas pequeñas
        TextOut pic.hdc, (xc2 - 1) * anccarP, y, tmp, Len(tmp)
        'Dibuja fondo de selección
        Call DibBloFonC(yc, xc1, xc2, mColFonSel)
        'Dibuja texto seleccionado
        pic.ForeColor = mColTxtSel
        If xc2 > xc1 Then   'Verificación de la menor corrdenada
            tmp = Mid$(linvis, xc1, xc2 - xc1)       'protege de líneas pequeñas
            TextOut pic.hdc, (xc1 - 1) * anccarP, y, tmp, Len(tmp)
        Else
            tmp = Mid$(linvis, xc2, xc1 - xc2)       'protege de líneas pequeñas
            TextOut pic.hdc, (xc2 - 1) * anccarP, y, tmp, Len(tmp)
        End If
    Else                    '--------Selección normal
        'Dibuja el fondo,
        If yt = sel1.yt And yt = sel2.yt Then 'única línea con la selección
            'Dibuja el texto antes del bloque
            pic.ForeColor = mColTxtNor
            tmp = Left$(linvis, xc1 - 1)    'protege de líneas pequeñas
            TextOut pic.hdc, 0, y, tmp, Len(tmp)
            'Dibuja el texto después del bloque
            tmp = Mid$(linvis, xc2)         'protege de líneas pequeñas
            TextOut pic.hdc, (xc2 - 1) * anccarP, y, tmp, Len(tmp)
            'Dibuja fondo de selección
            Call DibBloFonC(yc, xc1, xc2, mColFonSel)
            'Dibuja texto seleccionado
            pic.ForeColor = mColTxtSel
            tmp = Mid$(linvis, xc1, xc2 - xc1)       'protege de líneas pequeñas
            TextOut pic.hdc, (xc1 - 1) * anccarP, y, tmp, Len(tmp)
        ElseIf yt = sel1.yt Then    'primera línea de la seleción
            'Dibuja el texto antes del bloque
            pic.ForeColor = mColTxtNor
            tmp = Left$(linvis, xc1 - 1)    'protege de líneas pequeñas
            TextOut pic.hdc, 0, y, tmp, Len(tmp)
            'Dibuja fondo de selección
            xc2 = Len(linvis) + 1   'selecciona hasta el final
            xc2 = xc2 + 1           'un caracter más para indicar que la selección va hasta el final
            Call DibBloFonC(yc, xc1, xc2, mColFonSel)
            'Dibuja texto seleccionado
            pic.ForeColor = mColTxtSel
            tmp = Mid$(linvis, xc1)       'protege de líneas pequeñas
            TextOut pic.hdc, (xc1 - 1) * anccarP, y, tmp, Len(tmp)
        ElseIf yt = sel2.yt Then    'última línea de la seleción
            'Dibuja el texto después del bloque
            pic.ForeColor = mColTxtNor
            tmp = Mid$(linvis, xc2)         'protege de líneas pequeñas
            TextOut pic.hdc, (xc2 - 1) * anccarP, y, tmp, Len(tmp)
            'Dibuja fondo de selección
            xc1 = 1                 'selecciona desde el inicio
            Call DibBloFonC(yc, xc1, xc2, mColFonSel)  'fondo de selección
            'Dibuja texto seleccionado
            pic.ForeColor = mColTxtSel
            tmp = Mid$(linvis, 1, xc2 - 1)          'protege de líneas pequeñas
            TextOut pic.hdc, 0, y, tmp, Len(tmp)
        ElseIf yt > sel1.yt And yt < sel2.yt Then   'línea completamente seleccionada
            'Dibuja fondo de selección
            xc1 = 1                 'selecciona desde el inicio
            xc2 = Len(linvis) + 1   'selecciona hasta el final
            xc2 = xc2 + 1   'un caracter más para indicar que la selección va hasta el final
            Call DibBloFonC(yc, xc1, xc2, mColFonSel)
            'Dibuja texto seleccionado
            pic.ForeColor = mColTxtSel
            TextOut pic.hdc, 0, y, linvis, Len(linvis)
        Else
            'Aquí significa que la línea no está en la zona de selección
            'Nunca debería darse esta condición, si se verifica antes de
            'llamar a esta rutina
        End If
    End If
End Sub

Private Sub DibujaLinBN(yt As Long, Optional BorraFondo As Boolean = True)
'Refresca una línea completa que está en la posición indicada en coordenadas
'de texto (1..nLin). Por defecto borra el fondo antes de dibujar el texto.
'Dibuja el texto considerando la zona seleccionada.
'El texto se dibuja siempre empezando en la coordenada x=0.
Dim y As Long
Dim yc As Long
Dim linvis As String
Dim lin As String
    'Verifica si escapa del texto
    If yt < 1 Or yt > nlin Then Exit Sub
    yc = yt - fil1 + 1   'Coord. de cursor
    If yc < 1 Or yc > maxLinVis Then Exit Sub   'fuera de pantalla
    'toma línea afectada
    linvis = Mid(linexp(yt), col1, maxColVis)
    y = (yt - fil1) * altcarP    'posición vertical en pixels
    'Dibuja fondo de línea
    If BorraFondo Then Call DibBloFonC(yc, 1, maxColVis + 1, mColFonEdi)
    If haysel Then
        'Dibuja texto con selección
        If yt >= sel1.yt And yt <= sel2.yt Then
            Call DibLinSel(y, yc, yt, linvis)
        Else    'Línea sin selección
            'Dibuja el texto sin selección
            pic.ForeColor = mColTxtNor
            TextOut pic.hdc, 0, y, linvis, Len(linvis)
        End If
    Else
        'Dibuja el texto sin selección
        pic.ForeColor = mColTxtNor
        TextOut pic.hdc, 0, y, linvis, Len(linvis)
    End If
End Sub

Private Sub DibujaLinActual()
'Dibuja la línea actual (curYt).
    DibujaLin curYt, True
End Sub

Private Sub DibujaLin(yt As Long, Optional BorraFondo As Boolean = True)
'Refresca una línea completa que está en la posición indicada en coordenadas
'de texto (1..nLin). Por defecto borra el fondo antes de dibujar el texto.
'Dibuja la línea de texto a partir de su Vector de descripción, logrando
'texto multicolor.
Dim xc As Integer
Dim yc As Long, yp As Long
Dim i As Integer, j As Integer
Dim seg As String   'texto del segmento
Dim txtVis As String
Dim pos As Integer  'variable para buscar segmento
Dim tamaju As Integer   'tamaño ajustado del primer segmento
Dim txt As String
Dim lin() As TDesSeg
    'Verifica si escapa del texto
    If yt < 1 Or yt > nlin Then Exit Sub
    yc = yt - fil1 + 1   'Coord. de cursor
    If yc < 1 Or yc > maxLinVis Then Exit Sub   'fuera de pantalla
    
    'Verifica si hay selección para dbujar en B/N
    If LineaEnSel(yt) Then
        'No nos complicamos
        Call DibujaLinBN(yt, BorraFondo)
        Exit Sub
    End If
    
    'analiza toda la línea porque los colores
    'pueden depender de elementos no visibles
    txt = linexp(yt)                'lee toda la línea
    If lincol(yt).tip = TLIN_DES Then
        'No hay descripción, hay que actualizar
        ActualDescColFil txt, lin()
        lincol(yt).tip = TLIN_MIX
        lincol(yt).seg = lin    'copia descripción de línea
    Else    'ya tiene descripción de segmento
        lin() = lincol(yt).seg()
    End If
    
    yp = (yt - fil1) * altcarP  'actualiza coordenada vertical en pixeles
    'Dibuja fondo de línea
    If BorraFondo Then Call DibBloFonC(yc, 1, maxColVis + 1, mColFonEdi)
    '-------Busca segmento de inicio "i"-------
    i = 1   'Se asume que empieza en el segmento 1
    pos = 1
    For i = 1 To UBound(lin)
        pos = pos + lin(i).tam
        If pos > col1 Then
            'Este es el segmento que contiene el punto inicial.
            'Calcula el tamaño ajustado del primer segmento a
            'dibujar que puede estar fragmentado.
            tamaju = pos - col1
            Exit For
        End If
    Next
    txtVis = Mid$(txt, col1, maxColVis)     'texto visible para graficar
    'Si no cae en ningún segmento, i termina con Ubound(lin)+1
    '-------Imprime desde el segmento "i"------
    xc = 1  'caracter inicial
    For j = i To UBound(lin)
        pic.ForeColor = lin(j).col
        If j = i Then
            'Sólo para el primer segmento, el tamaño puede variar
            seg = Mid$(txtVis, xc, tamaju)       'texto del segmento
            TextOut pic.hdc, (xc - 1) * anccarP, yp, seg, Len(seg)
            xc = xc + tamaju
        Else
            seg = Mid$(txtVis, xc, lin(j).tam)   'texto del segmento
            TextOut pic.hdc, (xc - 1) * anccarP, yp, seg, Len(seg)
            xc = xc + lin(j).tam
        End If
        If xc > maxColVis Then
            Exit For 'ya no es visible
        End If
    Next
End Sub

Private Sub Dibujar(Optional RefresPIC As Boolean = True)
'Dibuja el texto completo que es visible en el control "pic". Dibuja en un solo color.
'Por defecto limpia el fondo a dibujar
Dim yp As Long
Dim i As Long
Dim textVis As String    'texto visible en la ventana
Dim numlin As String
Dim x1 As Long, x2 As Long
Dim rc As RECT
    'rc.Left = 0: rc.Top = 0    'ya está inicializado
    If RefresPIC Then   'limpia ventana
'        pic.Refresh: Exit Sub   'El "Refresh", llamará a "Paint" y a Dibujar()
        rc.Right = pic.ScaleWidth: rc.Bottom = pic.ScaleHeight
        hBrush = CreateSolidBrush(mColFonEdi)
        FillRect pic.hdc, rc, hBrush 'Es más rápido que "Rectangle"
    End If
Call RecalculaMaxColVis
    If nlin = 0 Or fil1 = 0 Then Exit Sub
    If maxLinVis = 0 Or maxColVis = 0 Then Exit Sub
    yp = 0
    LeerXPosSel x1, x2
    FijaRelleno mColFonSel   'Relleno para la selección
    'Dibuja por líneas
    For i = fil1 To nfilFin
        If i = curYt Then
            DibujaLinActual
        Else
            DibujaLin i, False  'dibuja sin pintar el fondo
        End If
    Next
    If verNumLin Then   'Dibuja número de línea
        'Borra columna
        rc.Right = ancNLinP: rc.Bottom = pic.ScaleHeight
        hBrush = CreateSolidBrush(mColFonNli)
        FillRect UserControl.hdc, rc, hBrush
        UserControl.ForeColor = vbBlack
        'Dibuja números
        yp = 0
        For i = fil1 To nfilFin
            numlin = i
            TextOut UserControl.hdc, 0, yp + 2, numlin, Len(numlin)
            yp = yp + altcarP    'actualiza coordenada vertical
        Next
    End If
    'Actualiza barra de desplazamiento
    pintando = True     'activa bandera para evitar lanzar otra vez el evento "Paint"
    If CInt((fil1 - 1) * facdesV + 1) > VScroll1.Max Then
        'Esta situación no debería producirse pero, parece que con muchas
        'lineas el redondeo puede ocasionar exceso
        MsgBox "Error de ajuste vertical: " & (fil1 - 1) * facdesV + 1 - VScroll1.Max
        VScroll1.Value = VScroll1.Max
    Else
        VScroll1.Value = (fil1 - 1) * facdesV + 1
    End If
    If col1 > HScroll1.Max Then
        MsgBox "Error de ajuste horizontal."
        HScroll1.Max = col1
        HScroll1.Value = col1
    Else
        HScroll1.Value = col1
    End If
    'MsgBox HScroll1.Max
    pintando = False
End Sub

Private Function CurMenorPos(pos As Tpostex) As Boolean
'Compara el cursor actual con una posición
Dim poscur As Tpostex
    poscur.xt = curXt
    poscur.yt = curYt
    CurMenorPos = MenorPos(poscur, pos)
End Function

Private Function CurMayorPos(pos As Tpostex) As Boolean
'Compara el cursor actual con una posición
Dim poscur As Tpostex
    poscur.xt = curXt
    poscur.yt = curYt
    CurMayorPos = MayorPos(poscur, pos)
End Function

Private Function PosSigPos(pos As Tpostex, n As Long) As Tpostex
'Devuelve la siguiente posición a partir de una posición base,
'proyectado "n" caracteres adelante. Trabaja en coordenadas de texto
'Sólo funciona en el modo de seleción normal.
Dim xt As Integer
Dim yt As Long
Dim lin1 As String  'primera línea
    PosSigPos = pos
    If n = 0 Then Exit Function
    lin1 = linrea(pos.yt)   'lee primera línea
    xt = PosXTreal(pos.xt, pos.yt)
    If xt + n <= Len(lin1) + 1 Then
        'Caso simple. No se escapa de la línea
        PosSigPos.yt = pos.yt
        PosSigPos.xt = PosXTexp(xt + n, pos.yt) 'posición expandida
    Else
        'Caso más complicado, porque pasa a otras líneas
        n = n - (Len(lin1) + 1 - xt)    'quita tamaño del restante de línea 1
        yt = pos.yt
        While n > 0
            'Aquí deben estar los dos caracteres CR y LF
            n = n - 2   'Los quitamos
            yt = yt + 1     'nos movemos a la siguiente línea
            If n < 0 Then
                'Rayos!, algo no anda bien
                MsgBox "Error de ajuste de salto de línea"
                Exit Function
            End If
            'Aquí puede estar la siguiente línea
            If n = 0 Then
                'Ya no hay más líneas
                PosSigPos.yt = yt   'en la siguiente
                PosSigPos.xt = 1    'al inicio
                Exit Function
            ElseIf n > 0 Then
                'Aún no se cubren todos los caracteres
                If n > Len(linrea(yt)) Then
                    'Toma siguiente línea completa
                    n = n - Len(linrea(yt))
                Else
                    'Es la última línea
                    PosSigPos.yt = yt   'en la siguiente
                    PosSigPos.xt = PosXTexp(n + 1, yt)
                    Exit Function
                End If
            End If
        Wend
    End If
End Function

Public Sub SelecIdentificador(xt As Integer, yt As Long, largo As Long)
'Selecciona una cadena en el editor a partir de la posición indicada. Recibe coordenadas
'de texto. No redibuja.
Dim p As Tpostex
    p.xt = PosXTexp(xt, yt)     'Necesita la posición expandida
    p.yt = yt
    'Selecciona cadena
    posCursorA p
    Call FijarSel0      'Fija punto base
    posCursorA2 PosSigPos(p, largo)
    Call ExtenderSel 'Extiende selección
End Sub

Private Sub DibTextPos(p1 As Tpostex, p2 As Tpostex)
'Dibuja las líneas entre las posiciones p1 y p2.
'Si p1 es menor que p2, se corrige.
Dim f As Long
    If p1.yt <= p2.yt Then
        For f = p1.yt To p2.yt
            Call DibujaLin(f)
        Next
    Else
        For f = p2.yt To p1.yt
            Call DibujaLin(f)
        Next
    End If
End Sub

Private Sub ActLimitesBarDesp()
'Actualiza los límites de las barras de desplazamiento
Dim nLinVis As Integer
    'Calcula el número de líneas visibles en el editor
    nLinVis = nfilFin - fil1 + 1

    If nlin > 32700 Then
        'Verifica si se puede trabajar con desplazamientos simples
        'de la barra vertical
        facdesV = 1 / (Int(nlin / 32700) + 1)
    Else    'Si se puede
        facdesV = 1     'factor de desplazamiento vertical
    End If
    If nlin + 1 > maxLinVis Then
        VScroll1.Enabled = True
        VScroll1.Max = (nlin - nLinVis) * facdesV + 1 'Valor máximo
        If nLinVis > 0 Then VScroll1.LargeChange = nLinVis
    Else
        VScroll1.Enabled = False
    End If
    
    If maxTamLin + 1 > maxColVis Then
        HScroll1.Enabled = True
        HScroll1.Max = MaxCol1
        If maxColVis > 0 Then HScroll1.LargeChange = maxColVis
    Else    'deshabilita
        HScroll1.Enabled = False
    End If
End Sub

'**********************************************************************************
'******************************FUNCIONES DE NIVEL MAYOR****************************
'**********************************************************************************
Public Sub CargarArch(arch As String)
'Método inicial para cargar un archivo de texto en el editor
Dim nar As Integer
Dim linea As String
Dim a() As String
Dim i As Long
    If Dir(arch) = "" Then
        MsgBox "No se encuentra archivo: " & arch
        Exit Sub
    End If
    Call LimpiarLineas
    'abre archivo de datos
    archivo = arch
    nar = FreeFile
    Open archivo For Input As #nar
    'OJOOOO que si hay una línea en blanco al final, no se lee
    'porque "Line Input" lee incluyendo el salto de línea final si lo encuentra
    Do While Not EOF(nar)
        Line Input #nar, linea
        If InStr(linea, Chr(10)) <> 0 Then
            'Hay saltos de línea, debe ser un formato unix
            tipArch = 1
            a = Split(linea, Chr(10))
            For i = 0 To UBound(a)
                AgregaLinea a(i)    'Agregamos línea
            Next
        Else
            AgregaLinea linea   'Agrega la línea tal como se lee
        End If
        If nlin > MAX_LIN_EDI Then
            MsgBox "Demasiadas líneas para leer en el editor"
            Exit Do
        End If
    Loop
    Close #nar
    tCursorA 1, 1
    Call ActualizaNLinFin   'actualiza las líneas visible
'    Call UserControl_Resize
    Call Dibujar
    Call InicDeshacer       'Inicia deshacer
    Call FijarTextNoModif   'Fija punto de "No Modificado"
End Sub

Public Sub GuardarArch()
'Método inicial para guardar el archivo de texto del editor
    If archivo = "" Then
        MsgBox "No se ha especificado un nombre para el archivo", vbExclamation
        Exit Sub
    End If
    GuardarArchComo archivo
End Sub

Public Sub GuardarArchComo(nomb As String)
'Método inicial para guardar el archivo de texto del editor
Dim nar As Integer
Dim f As Long
    archivo = nomb  'actualiza nombre actual
    nar = FreeFile
    Open archivo For Output As #nar
    For f = 1 To nlin
        Print #nar, linrea(f)
    Next
    Close #nar
'    Call InicDeshacer       'Inicia deshacer
    Call FijarTextNoModif   'Fija punto de "No Modificado"
End Sub

Private Sub LimpSelec(Optional refres As Boolean = False)
'Quita el área seleccionada, del control. Por defecto no actualiza la pantalla,
'pero si "refres" es TRUE, se redibuja la pantalla si es necesario
    If haysel Then      'Había una selección
        haysel = False  'desactiva
        If refres Then  'Hay que refrescar
            If sel1.yt = sel2.yt Then   'Sólo basta con dibujar la línea
                Call DibujaLin(sel1.yt)
            Else    'Hay que dibujar varias líneas
                Call Dibujar
            End If
        End If
        sel1 = sel0         'Inicia selección nula
        sel2 = sel0         'Inicia selección nula
    End If
End Sub

Public Sub SeleccionaTodo()
'Selecciona todo el texto contenido en el editor
    If nlin = 0 Then Exit Sub
    Redibujar = False
    'se mueve al inicio
    tCursorA 1, 1
    Call FijarSel0      'Fija punto base
    'nos movemos al final
    tCursorA2 maxTamLin + 1, nlin
    curXd = curXt      'actualiza posición deseada
    Call ExtenderSel
End Sub

Public Sub SelecLinea(yt As Long)
'Función pública para seleccionar línea del texto
    Redibujar = True       'para no complicarnos, dibuja todo
    If haysel Then Call LimpSelec
    'Selecciona línea
    tCursorA 1, yt
    Call FijarSel0      'Fija punto base
    tCursorA2 maxTamLin + 1, yt
    Call ExtenderSel 'Extiende selección
End Sub

Public Sub PegaSeleccion()
'Pega el texto seleccionado al control
'No hay problema en pegar selecciones grandes de texto
    On Error GoTo errPegSel
    CurInsertar Clipboard.GetText
    On Error GoTo 0
    Exit Sub
errPegSel:
    MsgBox "Error Leyendo el portapapeles.", vbExclamation
    On Error GoTo 0
End Sub

Public Sub CortaSeleccion()
'Elimina el texto seleccionado y copia la selección al portatpapeles
    Call CopiaSeleccion
    If haysel Then Call CurEliminar
End Sub

Public Sub CopiaSeleccion()
'Copia el texto seleccionado al portapapeles
Dim hData As Long   'Manejador de memoria
Dim nbytes As Long
    If nlin = 0 Then Exit Sub   'por seguridad
    '---------Obtiene selección-------------
    'Copia selección a memoria
    hData = BloAMem(nbytes, sel1, sel2)
    If hData = 0 Then
        MsgBox "No se puede copiar texto. Error asignando memoria", vbCritical
        Exit Sub
    End If
    '---------Copia al portapapeles-------------
    If OpenClipboard(0) Then
        Call EmptyClipboard
        'SetClipboardData CF_METAFILEPICT, hGlobal
        Call SetClipboardData(CF_TEXT, hData)
        'No es necesario ya liberar "hData", porque el portapapapeles lo hará
        'cuando ya no lo necesite
        Call CloseClipboard
    Else    'Fallo abrir el Portapapeles
        'Liberamos la memoria porque no la vamos a usar
        GlobalFree hData
    End If
End Sub

Private Function ElimBloq(s1 As Tpostex, s2 As Tpostex) As Long
'Elimina el bloque de texto definido por s1 y s2. Actualiza s1 y s2
'No realiza refresco de pantalla. Devuelve las filas eliminadas.
Dim f1 As Long, f2 As Long
Dim f As Long
Dim xt1 As Integer, xt2 As Integer
Dim xtmp As Integer
Dim tmp As String
    f1 = s1.yt    'posición vertical
    f2 = s2.yt    'posición vertical
    If tipSelec = 1 Then    '---------Selección por columnas----------
        'calcula posiciones horizontales en el texto
        xt1 = s1.xt: xt2 = s2.xt
        If xt1 > xt2 Then   'Verifica si hay que invertir
            xtmp = xt1: xt1 = xt2: xt2 = xtmp
        End If
        'Quita selección de las líneas afectadas
        For f = f1 To f2
            'Para no complicarnos con las tabulaciones, expandimos toda la línea
            'antes de eliminar. Lo optimo sería analizar que tabulaciones son
            'afectadas y expandirlas sólo a ellas.
            linrea(f) = LineaFin(linrea(f))
            EliminarCad linrea(f), xt1, xt2 - xt1
        Next
        ElimBloq = f2 - f1 + 1 'número de filas eliminadas
    Else              '---------------Selección normal----------------
        'calcula posiciones horizontales reales
        'se asume que siempre s1 está antes que s2
        xt1 = PosXTreal(s1.xt, s1.yt)
        xt2 = PosXTreal(s2.xt, s2.yt)
        'Elimina de acuerdo a la posición
        If f1 = f2 Then     'Están en la misma fila
            EliminarCad linrea(f1), xt1, xt2 - xt1
            ElimBloq = 1   'zólo se elimina de una fila
        ElseIf f1 < f2 Then
            'Procesa fila contiene inicio de selección
            EliminarCad linrea(f1), xt1
            'Procesa fila contiene fin de selección
            EliminarCad linrea(f2), 1, xt2 - 1
            'Elimina las filas intermedias
            If linrea(f2) = "" Then
                '-----Se eliminan la fila final completa
                EliminarLineas f1 + 1, f2 - f1
            Else
                '-----Se eliminan filas a medias. Hay que juntar
                'rescata lo que va a quedar de la fila final
                tmp = linrea(f2)
                'elimina hasta la fila final
                EliminarLineas f1 + 1, f2 - f1
                'agrega lo que rescató
                linrea(f1) = linrea(f1) & tmp
            End If
            ElimBloq = f2 - f1 + 1 'número de filas eliminadas
        End If
    End If
    'Actualiza nuevas posiciones de selección
    s2 = s1
    'Actualiza por si han desaparecido líneas de la pantalla
    Call ActualizaNLinFin
End Function

Private Function ElimSelecDib(Optional refres As Boolean = True) As Long
'Elimina la zona seleccionada. Actualiza las variables sel1 y sel2
'Refresca la pantalla por defecto. Devuelve las filas eliminadas.
'Se asume que la selección debe contener al cursor actual.
Dim f As Long
Dim yt2 As Long
Dim nfilsel As Long
    If bloqText Then Exit Function   'Hay protección
    yt2 = sel2.yt   'Guarda posición final antes de eliminar
                    'bloque, por si se necesita
    nfilsel = ElimBloq(sel1, sel2)
    If nfilsel = 1 Then    'sólo modificó una línea
        lincol(curYt).tip = TLIN_DES    'Fuerza a actualizar colores
        Redibujar = False   'prepara bandera
        'ubica de nuevo el cursor para validar posición
        posCursorA2 sel1    'Aquí se puede actualizar "Redibujar"
    Else    'Se han modificado varias líneas
        'Actualiza estado de sintaxis
        If tipSelec = 1 Then    'Varias filas alteradas
            For f = sel1.yt To yt2
                lincol(f).tip = TLIN_DES     'Actualizar en la 1ra línea
            Next
        Else
            lincol(sel1.yt).tip = TLIN_DES    'Actualizar en la 1ra línea
        End If
        'ubica cursor
        posCursorA2 sel1
        Redibujar = True    'Fuerza a redibujar
    End If
    haysel = False  'antes de dibujar, para refrescar correctamente
    If refres Then  'Verifica si se debe refrescar
        If Redibujar Then
            'Ha habido movimiento de pantalla o se han eliminado varias líneas
            Call Dibujar
'            Call ActualizaNLinFin    'Para dibujar correctamente
            Call ActLimitesBarDesp 'actualiza límites de Scroll Bar's
        Else    'Sólo es necesario refrescar la línea afectada
            Call DibujaLin(curYt)
        End If
        Call EncenCursor    'el cursor estaba desactivado
    End If
    Call FijarSel0      'Fija punto base de selección
    ElimSelecDib = nfilsel  'Devuelve filas eliminadas
End Function

Public Sub CurInsertar(cad As String)
'Inserta una cadena en la posición actual del cursor
'Guarda información para deshacer
Dim xt As Integer, yt As Long
Dim xtr As Integer
Dim tmp As String
Dim f As Long
Dim lcad() As String
Dim ncad As Long
Dim tam As Integer  'Tamaño de cadena
Dim pcur As Tpostex 'Guarda posición de cursor
Dim pmax As Tpostex 'Guarda posición de cursor
Dim nadi As Long    'Líneas adicionales
Dim s1 As Tpostex, s2 As Tpostex    'para selección
'Dim yt2 As Long     'yt2 para el modo multilínea
Dim nfilsel As Long
    If bloqText Then Exit Sub   'Hay protección
    
'    If cad = "" Then Exit Sub
    If nlin = 0 Then    'caso especial. No hay nada
        'Es el primer caracter que se crea
        AgregaLinea ""
        maxTamLin = 1
        FijaCursor 1, 1
        curXd = 1     'Inicia posición X deseada de cursor
        Call FijarSel0  'inicia parámetros de selección
    End If
    'Verifica si hay texto seleccionado
    If haysel Then
        'Se debe primero eliminar el texto selecionado.
        GuarAcc TU_ELI, sel1, TextSel()  'para deshacer
        s1 = sel1: s2 = sel2    'Guardar por si acaso
        nfilsel = ElimSelecDib(False) 'Elimina sin refrescar
        'Verifica si hay Inserción en modo multilínea
        If tipSelec = 1 And nfilsel > 1 And InStr(cad, vbCrLf) = 0 Then
            'Crea rápidamente una cadena de varias líneas
            ReDim lcad(1 To nfilsel)
            For f = 1 To nfilsel: lcad(f) = cad: Next
            tmp = Join(lcad, vbCrLf)
            'Inserta en la posición actual, se hace recursivamente para
            'evitar lidiar con el "deshacer" y los detalles del refresco.
            CurInsertar tmp     'Llamada recursiva
            'Restaura la selección para poder seguir escribiendo
            s1.xt = s1.xt + Len(cad)    'corrige desplazamiento
            's2.xt = s2.xt + Len(cad)    'corrige desplazamiento
            s2.xt = s1.xt   'debe haber siempre cero columnas
            sel1 = s1: sel2 = s2
            haysel = True
            Call Refrescar
            Exit Sub        'No hay más que hacer
        End If
    End If
    'Calcula posición horizontal real en la cadena considerando
    'tabulaciones. Se hace después de eliminar una posible selección
    'para tener la nueva posición
    'Toma posiciones en el texto
    xt = curXt
    yt = curYt
    If xt < 1 Or yt < 1 Then Exit Sub
    xtr = PosXTreal(xt, yt)
    'Inserta cadena en el texto
    If InStr(cad, vbCrLf) = 0 Then
        If Len(linrea(yt)) + Len(cad) > MAX_ANC_LIN Then   'Protección
            MsgBox "ERROR: Línea muy larga"
            Exit Sub
        End If
        GuarAcc TU_INS, LeePosCur, cad   'Guarda acción para deshacer
        '------- Insertar Cadena de una sola línea -----------
        InsertarCad linrea(yt), xtr, cad
        lincol(yt).tip = TLIN_DES   'Fuerza a actualizar colores
        linact = linexp(yt)    'actualiza línea actual porque se ha cambiado
                               'Se debe hacer antes de llamar a DesplazaCursor()
        'Actualiza tamaño de línea más larga
        'Se deba hacer antes de DesplazaCursor() para tener actualizada
        'la barra de desplazamiento antes de dibujar
        If Len(linact) > maxTamLin Then
            maxTamLin = Len(linact)
            Call ActLimitesBarDesp  'actualiza barra
        End If
        If Redibujar Then   'Verifica si es necesario dibujar todo
            Call Dibujar
        Else
            Call DibujaLinActual        'Refresca la línea actual
        End If
        DesplazaCursor DIR_DER, Len(cad)    'Mueve cursor
    Else
        '------ Insertar cadena de varias líneas ---------
        If tipSelec = 1 Then '---------------Modo por columnas---------------
            pcur = LeePosCur    'Guarda posición actual de cursor
            pmax = MaxPos       'Guarda posición del final
            lcad = Split(cad, vbCrLf)  'separa en líneas
            ncad = UBound(lcad) 'último índice
            'Agrega fila por fila
            nadi = 0 'inicia contador
            For f = 0 To ncad
                If yt + f > nlin Then   'No hay línea. hay que agregarla
                    Call AgregaLinea(Space(xt - 1)) 'se le da el tamaño necesario
                    nadi = nadi + 1
                Else    'Ya hay línea.
                    'No nos complicamos y expandemos las tabulaciones
                    'Aquí podemos tener problemas con deshacer
                    linrea(yt + f) = linexp(yt + f)
                    tam = Len(linrea(yt + f))   'guarda tamaño
                    If tam < xt - 1 Then
                        'Cadena muy pequeña, completar con espacios
                        linrea(yt + f) = linrea(yt + f) & Space(xt - tam - 1)
                    End If
                End If
                InsertarCad linrea(yt + f), xt, lcad(f)
                lincol(yt + f).tip = TLIN_DES 'Fuerza a actualizar colores
            Next
            If nadi > 0 Then    '¿Hubo líneas agregadas?
                'Guarda la acción de deshacer de golpe
                For f = 1 To nadi
                    'Puede ser lento para miles de líneas
                    tmp = tmp & vbCrLf & Space(xt - 1)
                Next
                GuarAcc TU_INSn, pmax, tmp  'Insertar normal
            End If
            GuarAcc TU_INS, pcur, cad    'Guarda acción principal para deshacer
            'mueve cursor al final de la cadena pegada
            tCursorA2 xt + Len(lcad(0)), curYt + ncad
            curXd = curXt      'actualiza posición deseada
            Call ActualizaNLinFin   'Para dibujar correctamente
            Call ActLimitesBarDesp  'actualiza límites de Scroll Bar's
            Call FijarSel0      'Fija punto base, como lo haría "DesplazaCursor()"
            'Ha cambiado bastante, mejor actualizamos todo
            Call Dibujar
            Call EncenCursor    'Enciende para que sea visible en la nueva posición
        Else                 '---------------Modo normal---------------------
            GuarAcc TU_INS, LeePosCur, cad   'Guarda acción para deshacer
            lcad = Split(cad, vbCrLf)  'separa en líneas
            ncad = UBound(lcad)     'último índice
            tmp = linrea(yt)        'lee línea afectada
            'Inserta las líneas necesarias en la posición del cursor
            InsertarLineas yt, ncad
            'corta línea y agrega línea inicial del portapapeles
            linrea(yt) = Mid$(tmp, 1, xtr - 1) & lcad(0)
            lincol(yt).tip = TLIN_DES 'Fuerza a actualizar colores
            'actualiza líneas internmedias
            For f = 1 To ncad - 1
                linrea(yt + f) = lcad(f)
                lincol(yt + f).tip = TLIN_DES 'Para forzar a analizar colores
            Next
            'actualiza línea final
            linrea(yt + ncad) = lcad(ncad) & Mid$(tmp, xtr)
            lincol(yt + ncad).tip = TLIN_DES 'Fuerza a actualizar colores
            'mueve cursor al final de la cadena pegada
            tCursorA Len(lcad(ncad)) + 1, curYt + ncad
            curXd = curXt      'actualiza posición deseada
            Call ActualizaNLinFin   'Para dibujar correctamente
            Call ActLimitesBarDesp  'actualiza límites de Scroll Bar's
            Call FijarSel0      'Fija punto base, como lo haría "DesplazaCursor()"
            'Ha cambiado bastante, mejor actualizamos todo
            Call Dibujar
            Call EncenCursor    'Enciende para que sea visible en la nueva posición
        End If
    End If
End Sub

Public Sub CurEliminar(Optional forzardib As Boolean = False)
'Elimina un caracter en la posición actual del cursor.
'"forzardib" indica que se fuerza a redibujar toda la pantalla
'Guarda información para deshacer
Dim xt As Integer
Dim yt As Long
Dim xtr As Integer
Dim tmp As String
    If bloqText Then Exit Sub   'Hay protección
    '---------Si hay selección, sólo elimina y sale-------------
    If haysel Then
        GuarAcc TU_ELI, sel1, TextSel() 'para deshacer
        Call ElimSelecDib
        Exit Sub
    End If
    '-------No hay selección, se debe eliminar lo solicitado------
    If nlin = 0 Then Exit Sub   'no hay nada que eliminar
    'Toma posiciones en el texto
    xt = curXt
    yt = curYt
    If xt < 1 Or yt < 1 Then Exit Sub
    'Calcula posición horizontal real en la cadena considerando tabulaciones
    xtr = PosXTreal(xt, yt)
    If xtr = Len(linrea(yt)) + 1 And yt < nlin Then
        GuarAcc TU_ELI, LeePosCur(), vbCrLf    'para deshacer
        'Esta al final de la línea y hay más líneas.
        'copia contenido de línea siguiente
        tmp = linrea(yt + 1)
        'agrega a línea actual
        linrea(yt) = linrea(yt) & tmp
        lincol(yt).tip = TLIN_DES  'Fuerza a actualizar colores
        'reposiciona cursor para validar y actualizar "linact"
        tCursorA xt, yt
        curXd = curXt      'actualiza posición deseada
        'Eliminamos la siguiente línea
        EliminarLineas yt + 1, 1
        Call ActualizaNLinFin
        Call ActLimitesBarDesp  'Actualiza límites de Scroll Bar's
        Call FijarSel0      'Fija punto base, como lo haría "DesplazaCursor()"
        Call Dibujar    'Refresca. En realidad sólo podría ser necesario dibujar
                        'sólo las líneas de abajo.
    Else    'Caso de eliminación normal
        tmp = EliminarCad(linrea(yt), xtr, 1)
        lincol(yt).tip = TLIN_DES  'Fuerza a actualizar colores
        GuarAcc TU_ELI, LeePosCur(), tmp   'para deshacer
        linact = linexp(yt)    'actualiza línea actual porque se ha cambiado
        'Se debería actualizar "maxTamLin", si es que disminuye
        'pero es un proceso muy pesado para hacerse aquí
        If forzardib Then   '¿Dibujar todo?
            Call Dibujar    'Dibuja todo
        Else    'Sólo dibuja la línea afectada
            Call DibujaLinActual    'Refresca línea
        End If
        'El cursor se queda en su lugar
    End If
    Call EncenCursor    'Lo hace visible para visualizar mejor
End Sub

Public Sub CurEliminarCad(cad As String)
'Elimina "n" caracteres desde la posición actual del cursor en el control .
'El valor "n" se determina por el tamaño de "cad" y por el tipo de selección
'No se debería usar para editar el texto del editor, sino sólo para
'las opciones de "deshacer"
Dim p1 As Tpostex, p2 As Tpostex
Dim n As Long
Dim a() As String
Dim f As Long
Dim yt2 As Long
    If bloqText Then Exit Sub       'Hay protección
    If Len(cad) = 0 Then Exit Sub   'Verificación
    'Define el bloque a eliminar
    p1 = LeePosCur                  'Toma inicio de bloque
    If tipSelec = 1 Then '---------------Modo por columnas---------------
       a = Split(cad, vbCrLf)       'separa cadena
       p2.yt = p1.yt + UBound(a)    'calcula límite
       p2.xt = p1.xt + Len(a(0))    'calcula límite
       
    Else                 '---------------Modo normal---------------
        n = Len(cad)
        p2 = PosSigPos(p1, n)   'calcula límite de bloque
    End If
    yt2 = p2.yt     'Guarda fila final
    '-----------------Actualiza Pantalla--------------------
    If ElimBloq(p1, p2) = 1 Then    'sólo modificó una línea
        lincol(p1.yt).tip = TLIN_DES  'Para actualizar colores
        Redibujar = False   'prepara bandera
        'ubica de nuevo el cursor para validar posición
        posCursorA2 p1
        If Redibujar Then
            'Ha habido movimiento de pantalla
            Call Dibujar
        Else    'Sólo es necesario refrescar la línea afectada
            Call DibujaLin(curYt)
        End If
    Else    'Se han modificado varias líneas
        lincol(p1.yt).tip = TLIN_DES  'Para actualizar colores
        If tipSelec = 1 Then
            'Eliminación en varias líneas, marca las siguientes
            For f = p1.yt + 1 To yt2
                lincol(f).tip = TLIN_DES
            Next
        End If
        'ubica cursor
        posCursorA2 sel1
        Call Dibujar           'Refresca todo
        Call ActualizaNLinFin  'Para dibujar correctamente
        Call ActLimitesBarDesp 'actualiza límites de Scroll Bar's
    End If
    Call FijarSel0      'Fija punto base de selección
    Call EncenCursor    'el cursor estaba desactivado
End Sub

Public Sub CurEliminarB()
'Elimina un caracter en la posición actual del cursor, hacia atrás
Dim alinicio As Boolean
    If bloqText Then Exit Sub   'Hay protección
    '---------Si hay selección, sólo elimina y sale-------------
    If haysel Then ElimSelecDib: Exit Sub
    '-------No hay selección, se debe eliminar lo solicitado------
    If nlin = 0 Then Exit Sub   'no hay nada que eliminar
    'Retrocede el cursor
    Redibujar = False   'inicia bandera
    If curXt = 1 Then   'Está al inicio de línea
        If curYt > 1 Then       'hay línea anterior
            tCursorA2 Len(linexp(curYt - 1)) + 1, curYt - 1
            curXd = curXt      'actualiza posición deseada
        Else
            alinicio = True     'está al inicio
        End If
    Else
        tCursorA2 curXt - 1, curYt
        curXd = curXt      'actualiza posición deseada
    End If
    Call FijarSel0      'Fija punto base, como lo haría "DesplazaCursor()"
    'Elimina en la nueva posición. Indicando si es necesario redibujar
    If Not alinicio Then    'Verifica si se puede
        Call CurEliminar(Redibujar)    'Elimina en la nueva posición
    End If
End Sub

'*************************************************************************************
'************************* FUNCIONES PARA OPCIONES DE BÚSQUEDA ***********************
'*************************************************************************************

Private Function BuscarCadPos(bus As String, p1 As Tpostex, p2 As Tpostex, _
                    Optional ignCaja As Boolean = True) As Tpostex
'Busca una cadena de texto cargado en el editor, en el bloque definido
'por pos1 y pos2, a partir de la posición pos1 hacia adelante.
'Devuelve la posición donde empieza la cadena encontrada. Si no encuentra,
'devuelve (0,0)
Dim cad As String
Dim pos As Integer
Dim f As Long
Dim cmp As VbCompareMethod
    If MenorPos(p2, p1) Then Exit Function
    If ignCaja Then
        cmp = vbTextCompare
    Else
        cmp = vbBinaryCompare
    End If
    If p1.yt = p2.yt Then    'Texto de una sola línea-------
        cad = TextPosLin(p1, p2)
        pos = InStr(1, cad, bus, cmp) 'busca
        If pos <> 0 Then
            BuscarCadPos.xt = p1.xt + PosXTexp(pos, p1.yt)
            BuscarCadPos.yt = p1.yt
            Exit Function
        End If
    Else                     'Hay varias líneas---------------
        'busca en línea inicial
        cad = TextPosIni(p1)
        pos = InStr(1, cad, bus, cmp)   'busca
        If pos <> 0 Then
            BuscarCadPos.xt = p1.xt + PosXTexp(pos, p1.yt) - 1
            BuscarCadPos.yt = p1.yt
            Exit Function
        End If
        'busca en líneas intermedias
        For f = p1.yt + 1 To p2.yt - 1
            cad = linrea(f)
            pos = InStr(1, cad, bus, cmp) 'busca
            If pos <> 0 Then
                BuscarCadPos.xt = PosXTexp(pos, f)
                BuscarCadPos.yt = f
                Exit Function
            End If
        Next
        'busca en línea final
        cad = TextPosFin(p2)
        pos = InStr(1, cad, bus, cmp) 'busca
        If pos <> 0 Then
            BuscarCadPos.xt = PosXTexp(pos, p2.yt)
            BuscarCadPos.yt = p2.yt
            Exit Function
        End If
    End If
End Function

Public Sub InicBuscar(bus As String, _
                      Optional ambito As Integer = AMB_TODO, _
                      Optional ignCaja As Boolean = True, _
                      Optional palComp As Boolean = False)
'Inicia una búsqueda definiendo sus parámetros.
'La cadena "bus" debe ser de una sola línea.
    If ambito = AMB_TODO Then
        'Se buscará en todo el texto
        PosBus1 = MinPos
        PosBus2 = MaxPos
    ElseIf ambito = AMB_SELE Then
        'Se buscará en la selección
        PosBus1 = sel1
        PosBus2 = sel2
    End If
    PosEnc = PosBus1    'Fija posición inicial para buscar
    CadBus = bus        'Guarda cadena de búsqueda
    CajBus = ignCaja    'Guarda parámetro de caja
    PalCBus = palComp
End Sub

Public Function BuscarSig() As String
'Realiza una búsqueda iniciada con "InicBuscar"
'La búsqueda se hace a partir de la posición donde se dejó en la última búsqueda.
'Devuelve la cadena de búsqueda.
Dim p As Tpostex
Dim p2 As Tpostex
    PosEnc = LeePosCur()    'Busca desde el cursor
    'Protecciones
    If PosNulo(PosEnc) Then Exit Function
    If PosNulo(PosBus2) Then Exit Function
    If MayorPos(PosBus2, MaxPos) Then PosBus2 = MaxPos
    If MayorPos(PosEnc, PosBus2) Then Exit Function
    'búsqueda
    If PalCBus Then 'Debe ser palabra completa
        p = BuscarCadPos(CadBus, PosEnc, PosBus2, CajBus)
        Do While p.xt <> 0
            p2 = PosSigPos(p, Len(CadBus))
            If EsPalabraCompleta(p, p2) Then Exit Do
            PosEnc = p2
            p = BuscarCadPos(CadBus, PosEnc, PosBus2, CajBus)
        Loop
    Else    'Búsqueda normal
        p = BuscarCadPos(CadBus, PosEnc, PosBus2, CajBus)
        If p.xt <> 0 Then p2 = PosSigPos(p, Len(CadBus))
    End If
    'verifica si encontró
    If p.xt <> 0 Then
        Redibujar = True       'para no complicarnos, dibuja todo
        If haysel Then Call LimpSelec
        'Selecciona cadena
        posCursorA p
        Call FijarSel0      'Fija punto base
        posCursorA2 p2
        Call ExtenderSel 'Extiende selección
        BuscarSig = CadBus  'devuelve cadena
    Else
        BuscarSig = CadBus  'devuelve cadena
        MsgBox "No se encuentra el texto: '" & CadBus & "'", vbExclamation
    End If
End Function

Private Function EsPalabraCompleta(p1 As Tpostex, p2 As Tpostex) As Boolean
'Indica si la palabra en la posición [p,p+lar] es una palabra completa, es decir,
'que no está en medio de un identificador
Dim Cant As String
Dim Csig As String
    If p1.xt = 0 Then Exit Function
    Cant = UCase(CarPosAnt(p1)) 'caracter anterior
    Csig = UCase(CarPos(p2))    'caracter siguiente
    If Not (Cant Like CAR_IDEN_VALM) And Not (Csig Like CAR_IDEN_VALM) Then
        EsPalabraCompleta = True
    End If
End Function

'*************************************************************************************
'************************* FUNCIONES PARA EL MANEJO DE DESHACER **********************
'*************************************************************************************
Private Sub elimUndos(n As Integer)
'Elimina "n" acciones al inicio de la matriz Undos().
'La eliminación se hace siempre desde la acción 1, al inicio de la matriz.
Dim i As Integer
    If n > nUndo Then n = nUndo 'protección
    If n <= 0 Then Exit Sub     'por protección y evitar pérdida de tiempo cuando n=0
    'Desplaza elementos
    For i = 1 To nUndo - n
        Undos(i) = Undos(i + n)
    Next
    'Elimina elementos finales
    nUndo = nUndo - n
    ReDim Preserve Undos(nUndo)
    'Mantiene la distancia con nTxtModif
    'obsrevar que nTxtModif, puede ser negativo, lo que significa que
    'ya no se puede recuperar el estado de "texto no modificado" porque
    'ya no se tienen la cantidad de acciones disponibles que se necesitan
    'para llegar a este estado.
    nTxtModif = nTxtModif - n
End Sub

Private Sub InicDeshacer()
'Inicia la herramienta DEHACER. Debe llamarse al inicio del editor
'y cuando ya no se pueden (o deben) recuperar los cambios anteriores
    elimUndos nUndo     'Elimina todas las acciones que puedan existir
    Deshaciendo = False
End Sub

Private Sub FijarTextNoModif()
'Fija el punto en que el el texto no está modificado. Es decir
'el texto que no requiere grabación.
'Debería llamarse en el momento de abrir un archivo o grabarse.
    nTxtModif = nUndo   'Apunta al índice nUndo actual.
End Sub

Public Function TextModificado() As Boolean
'Indica si el texto ha sufrido modificación
    If nUndo = nTxtModif Then
        TextModificado = False   'El texto está como al inicio
    Else
        TextModificado = True    'Se ha modificado el texto
    End If
End Function

Private Sub GuarAcc(acc As Integer, pos As Tpostex, cad As String)
'Guarda una acción de modificación del texto para poder deshacerla luego
'Se debe llamar cada vez que se realiza un cambio en el texto
    If Deshaciendo Then Exit Sub    'No guarda acciones de tipo deshacer
    If TamMemDeshacer > TAM_MAX_UNDO Then
        InicDeshacer    'Protección a grabar muchas acciones
    End If
    If nUndo > NAC_MAX_UNDO Then
        'Se alcanzó el tamaño máximo de acciones a deshacer
        elimUndos 1     'elimina la más antigua
    End If
    nUndo = nUndo + 1   'incrementa contador de acciones
    ReDim Preserve Undos(nUndo) 'agrega elemento
    Undos(nUndo).acc = acc
    Undos(nUndo).pos = pos
    Undos(nUndo).cad = cad  'Copia cadena de la acción.
End Sub

Private Sub EjecAcc(acc As Integer, pos As Tpostex, cad As String)
'Ejecuta una acción de modificación del texto de tipo Tundo
Dim tipsel As Integer     'Modo de Selección
    Deshaciendo = True  'Para evitar guardar las acciones del "deshacer"
    Select Case acc
    Case TU_INS     'Inserta texto
        Redibujar = False
        posCursorA2 pos 'ubica cursor
        CurInsertar cad 'Inserta la cadena indicada
    Case TU_INSn    'Inserta texto en modo normal
        tipsel = tipSelec      'Guarda tipo de selección
        tipSelec = 0
        Redibujar = False
        posCursorA2 pos 'ubica cursor
        CurInsertar cad 'Inserta la cadena indicada
        tipSelec = tipsel   'Restaura tipo de selección
    Case TU_ELI     'Elimina texto
        Redibujar = False
        posCursorA2 pos 'ubica cursor
        CurEliminarCad cad   'Elimina la cadena indicada
    Case TU_ELIn    'Elimina en modo normal
        tipsel = tipSelec      'Guarda tipo de selección
        tipSelec = 0
        Redibujar = False
        posCursorA2 pos 'ubica cursor
        CurEliminarCad cad   'Elimina la cadena indicada
        tipSelec = tipsel   'Restaura tipo de selección
    Case TU_SNOR    'Pasa a modo normal
        tipSelec = 0
    Case TU_SCOL    'Pasa a modo por columnas
        tipSelec = 1
    End Select
    Deshaciendo = False
End Sub

Private Sub EjecutarAcc(acc As Tundo)
'Ejecuta una acción de tipo Tundo
    EjecAcc acc.acc, acc.pos, acc.cad
End Sub

Private Sub DeshacerAcc(acc As Tundo)
'Deshace una acción de tipo Tundo
    EjecAcc -acc.acc, acc.pos, acc.cad
End Sub

Public Sub Deshacer()
'Deshace una acción previamente realizada
    If nUndo > 0 Then
        DeshacerAcc Undos(nUndo)
        nUndo = nUndo - 1   'decrementa
        ReDim Preserve Undos(nUndo)  'elimina siguiente acción
    End If
End Sub

Private Function TamMemDeshacer() As Long
'Devuelve la cantidad de caracteres (no de bytes) usados en la
'grabación de las acciones de "deshacer".
Dim tmp As Long
Dim i As Integer
    For i = 1 To nUndo
        tmp = tmp + Len(Undos(i).cad)
    Next
    TamMemDeshacer = tmp
End Function

'*************************************************************************************
'***************************** CÓDIGO DE RESPUESTA A EVENTOS *************************
'*************************************************************************************
Private Sub pic_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
'Se ha soltado un archivo
Dim arc As String
    If Data.GetFormat(vbCFFiles) Then
        'Hay nombre de archivos
        arc = Data.Files(1)     'Devuelve el primer archivo
        'Dispara evento
        RaiseEvent ArchivoSoltado(arc)
    End If
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Se ha hecho click en el control
Dim xt As Integer
Dim yt As Long
'    If x < ancNLinP Then
'        'MsgBox "En columna de líneas"
'        Exit Sub
'    End If
    
    'Movemos el cursor en la posición indicada
    xt = Round(x / anccarP) + col1
    yt = Int(y / altcarP) + fil1
    Redibujar = False
    tCursorA2 xt, yt
    curXd = curXt      'actualiza posición deseada
    If Button = 1 Then
        '---------botón izquierdo----------
        If Shift = 1 Then   'Con shift encendido
            Call ExtenderSel
        Else
            Call LimpSelec(True)        'Limpia selección
            'Marca la posición del cursor antes de una selección
            Call FijarSel0
            If Redibujar Then Call Refrescar
            Call EncenCursor    'Enciende para que sea visible en la nueva posición
        End If
        'Inicia bandera para arrastre
        pulsadoI = True
        xt0 = xt: yt0 = yt  'Inicia coordenadas para movimiento
    Else
        '---------botón derecho----------
    End If
    ultBotPul = Button  'Actualiza último botón pulsado
    'Verifica desactivación de Ayuda Contextual
    If HayAyudC Then
        If curYt <> ytIniIden Then  'Se ha movido la línea
            FinAyudContextual   'Termina la ayuda contextual
        End If
    End If
    
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim xt As Integer
Dim yt As Long
'    If X  < ancNLinP Then
'        Exit Sub
'    End If
    If pulsadoI Then
        'Para indicar arrastre
'        UserControl.MousePointer = vbNoDrop
        '-----Continúa selección----
        'Movemos el cursor en la posición indicada
        xt = Int(x / anccarP) + col1
        yt = Int(y / altcarP) + fil1
        If xt0 = xt And yt0 = yt Then Exit Sub
        Redibujar = False
        tCursorA2 xt, yt      'Para realizar desplazamiento
        Call ExtenderSel
    End If
    xt0 = xt: yt0 = yt
'    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        pulsadoI = False
'        UserControl.MousePointer = vbDefault    'Termina arrastre
    ElseIf Button = 2 Then
        'Activa menú contextual si es que hay
        If Not (menuContext Is Nothing) Then
            PopupMenu menuContext
        End If
    End If
End Sub

Private Sub pic_DblClick()
    If ultBotPul = 1 Then
        '------Doble click izquierdo ---------
        Redibujar = False
        tCursorA2 curIniPal2(), curYt
        Call FijarSel0
        tCursorA2 curFinPal(), curYt
        Call ApagCursor 'Para que no interfiera
        Call ExtenderSel
    End If
End Sub

Private Sub UserControl_Initialize()
Dim tm As TEXTMETRIC
    'Oculta lista contextual
    lstCont.Visible = False
    
    
    pic.ScaleMode = vbPixels
    UserControl.ScaleMode = vbPixels
    facdesV = 1
    Call LimpiarLineas
    
    verDesHor = False
    
    FijaFil1 1      'fil1= 1
    FijaCol1 1      'col1= 1
    nEspTab = 6     '6 espacios para tabulación
    
    'Inicia colores del editor
    mColFonEdi = vbWhite  'vbblack     'fondo de editor
    mColTxtNor = vbBlack  'vbWhite     'texto normal
    mColFonSel = RGB(130, 130, 255)
    mColTxtSel = vbWhite
    mColFonNli = RGB(200, 200, 200)
    mColTxtCom = RGB(128, 128, 128)
    mColTxtCad = RGB(90, 90, 255)
    mColPalRes = RGB(0, 190, 0)     'verde oscuro
    mColPalRes2 = vbRed     'verde oscuro
    mColTxtFun = RGB(0, 190, 0)     'verde oscuro
    
    pic.BackColor = mColFonEdi
    
    'Define el tipo de texto
    
    
    
    Call FijaTexto(mColTxtNor, 9, 0, "Courier New")
    UserControl.Font = "Courier New"
    UserControl.FontSize = 9
    
    Call GetTextMetrics(pic.hdc, tm)
    anccarP = tm.tmAveCharWidth  'Siempre en pixels
    altcarP = tm.tmHeight        'Siempre en pixels
    
    maxTamLin = 0   'no hay líneas cargadas aún
    VScroll1.Min = 1
    HScroll1.Min = 1
    pintando = False
    
    'inicia cursor
    curXt_ant = 1
    curYt_ant = 1
    FijaCursor 1, 1
    curXd = 1     'Inicia posición X deseada de cursor
    Call FijarSel0  'inicia parámetros de selección
    ActivarCursor
    Call InicDeshacer
    Call InicBuscar("")
    Call InicAyudaContext("")   'Inicia ayuda contextual
    tipArch = 0     'Tipo DOS
    'Posiciones fijas de la barra de estado
    VScroll1.Top = 0
    HScroll1.Left = 0
    
    pic.Top = 0     'Ubica el picture en el control
End Sub

Private Sub UserControl_Terminate()
'Elimina objetos creados
    DeleteObject hPen
    DeleteObject hFont
    DeleteObject hBrush
End Sub

Public Sub Refrescar()
'Actualiza la apriencia y el contenido del control
    Call UserControl_Resize
End Sub

Private Sub pic_Paint()
    'Dibuja sin refrescar PIC para evitar llamadas recursivas
    'Además se supone que si se refresca es porque se ha borrado
    'la información.
    Call Dibujar(False)
End Sub

Private Sub RecalculaMaxColVis()
'Realiza el Redimensionamiento horizontal, calculando "maxColVis" de acuerdo
'a si se debe mostrar el número de línea y aal número de líneas totales .
Dim ancVScroll As Long
    If verDesVer Then
        ancVScroll = VScroll1.Width
    Else
        ancVScroll = 0
    End If
    If verNumLin Then
        If nlin > 100000 Then
            ancNLinP = 47
        ElseIf nlin > 10000 Then
            ancNLinP = 40
        ElseIf nlin > 1000 Then
            ancNLinP = 33
        ElseIf nlin > 100 Then
            ancNLinP = 26
        Else
            ancNLinP = 19
        End If
    Else
        ancNLinP = 2    'deja un pequeño espacio lateral
    End If
    'posiciona el "pic"
    pic.Left = ancNLinP
    pic.Width = Posit(ScaleWidth - ancVScroll - ancNLinP)
    maxColVis = pic.Width \ anccarP
    col2 = col1 + maxColVis - 1     'actualiza col2
End Sub

Private Sub UserControl_Resize()
Dim altHScroll As Single   'alto de HScroll1
Dim altEstado As Single   'alto de HScroll1
Dim ancVScroll As Single  'ancho de VScroll1
    'Fija ancho de Barra de desplazamiento horizontal
    If verDesVer Then
        HScroll1.Width = Posit(ScaleWidth - VScroll1.Width)
    Else
        HScroll1.Width = ScaleWidth
    End If
    'si se quiere dejar espacio para algo
    'HScroll1.Width = Posit(ScaleWidth - VScroll1.Width - 500)
    'Posición de barra de desplazamiento vertical
    If verDesVer Then
        VScroll1.Height = ScaleHeight
        ancVScroll = VScroll1.Width
        VScroll1.Left = ScaleWidth - ancVScroll
        VScroll1.Visible = True
    Else
        ancVScroll = 0
        VScroll1.Visible = False
    End If
    Call RecalculaMaxColVis
    'Redimensionamiento vertical
    If verDesHor Then
        altHScroll = HScroll1.Height
        'aún no se sabe la posición vertical de HScroll1
        HScroll1.Visible = True
    Else
        altHScroll = 0
        HScroll1.Visible = False
    End If
    altEstado = 0
    pic.Height = Posit(ScaleHeight - altHScroll - altEstado)
    'pic.Left = ancNLinP
    'pic.Width = Posit(ScaleWidth - ancVScroll - ancNLinP)
    HScroll1.Top = pic.Height
    'máximo que se puede mostrar
    maxLinVis = pic.ScaleHeight \ altcarP
    If maxLinVis = 0 Then
        Exit Sub  'Probablemente minimizado o demasiado pequeño
    End If
    fil2 = fil1 + maxLinVis - 1
    Call ActualizaNLinFin   'actualiza las líneas visible
    Call Dibujar            'refresca por si acaso. Aquí se puede cambiar el
                            'dimensionamiento horizontal
    Call ActLimitesBarDesp  'actualiza límites de Scroll Bar's
End Sub

Private Sub pic_KeyDown(KeyCode As Integer, Shift As Integer)
Dim conShift As Boolean
Dim conCtrl As Boolean
    RaiseEvent KeyDown(KeyCode, Shift)  'dispara evento
    If KeyCode = 67 And Shift = 4 Then
        RaiseEvent CambiaModo   'Petición de cambio de modo
    End If
    If KeyCode = 16 Then Exit Sub   'Ignora <Shift> solo
    If KeyCode = 17 Then Exit Sub   'Ignora <Ctrl> solo
    If KeyCode = 18 Then Exit Sub   'Ignora <Alt> solo
    If (Shift And 1) = 1 Then conShift = True
    If (Shift And 2) = 2 Then conCtrl = True
    If AyudContextKeyDown(KeyCode, Shift) Then
        'No procesamos la tecla porque ya lo procesó el menú de ayuda
        Exit Sub
    End If
    Select Case KeyCode
    'Teclas de desplazamiento
    Case 39 'flecha derecha
        If conCtrl Then    'Ctrl
            DesplazaCursor DIR_DERPAL, , conShift
        Else
            DesplazaCursor DIR_DER, , conShift
        End If
    Case 37 'flecha izquierda
        If conCtrl Then    'Ctrl
            DesplazaCursor DIR_IZQPAL, , conShift
        Else
            DesplazaCursor DIR_IZQ, , conShift
        End If
    Case 40 'flecha abajo
        If conCtrl Then
            DesplazaCursor DIR_ABAPAR, , conShift
        Else
            DesplazaCursor DIR_ABA, , conShift
        End If
    Case 38 'flecha arriba
        If conCtrl Then
            DesplazaCursor DIR_ARRPAR, , conShift
        Else
            DesplazaCursor DIR_ARR, , conShift
        End If
    Case 34     'pag.abajo
        DesplazaCursor DIR_PABA, maxLinVis - 1, conShift
    Case 33     'pag.arriba
        DesplazaCursor DIR_PARR, maxLinVis - 1, conShift
    Case 36     'inicio
        If conCtrl Then    'Ctrl
            DesplazaCursor DIR_HOM, , conShift
        Else
            DesplazaCursor DIR_INI, , conShift
        End If
    Case 35     'fin
        If conCtrl Then    'Ctrl
            DesplazaCursor DIR_END, , conShift
        Else
            DesplazaCursor DIR_FIN, , conShift
        End If
    Case 116    'F5 refresca
        Call Dibujar    'Dibuja el control
    'de modificación
    Case 46     'DEL
        If conShift Then
            Call CortaSeleccion
        Else
            Call CurEliminar
        End If
    Case 45     'INSERT
        If conCtrl Then
            Call CopiaSeleccion
        ElseIf conShift Then
            Call PegaSeleccion
        End If
    Case 13
        
        CurInsertar vbCrLf
    Case 8
        Call CurEliminarB
    End Select
    Call AyudContextKeyDown2(KeyCode, Shift)
End Sub

Private Sub pic_KeyPress(KeyAscii As Integer)
'Procesa una tecla pulsada
    
    If KeyAscii >= 32 Or KeyAscii = 9 Then
'        If Not HayAyudC Then
            CurInsertar Chr(KeyAscii)
'        End If
    End If
'    If KeyAscii = 27 And Not HayAyudC Then
'        RaiseEvent TeclaEscape  'Evento de escape
'    End If
    Call AyudContextKeyPress(KeyAscii)
End Sub

'*************************************************************************************
'************************ FUNCIONES DEL MENÚ DE AYUDA CONTEXTUAL *********************
'*************************************************************************************
Public Sub InicAyudaContext(cad_con As String)
'Inicia el motor de ayuda contextual
Dim nar As Integer
    nFilAyudC = 5   'Número de filas en el menú contextual
    ListandoTab = False
    ArcListaTab = App.Path & "\lista.lst"
    ReDim ListaTablas(0)
    ReDim IdentAyudC(0)     'Inicia identificadores
End Sub

Private Sub ExcribeElem(matcad() As String, nele As Integer, elem As String)
'Escribe elemento en matriz de cadena, verificando límite.
    If nele > UBound(matcad) Then
        'Faltó espacio
        ReDim Preserve matcad(nele + 100)
    End If
    'Escribe elemento
    matcad(nele) = elem
    nele = nele + 1 'actualiza índice
End Sub

Private Sub LlenaIdenAyudContextual(Optional agregar As Boolean = False)
'Llena las palabras reservadas. Si "agregar" es TRUE, no se limpia la lista.
Dim n As Integer
    If agregar Then n = UBound(IdentAyudC) + 1 Else n = 1
    ExcribeElem IdentAyudC, n, "ALUMINIO"
    ExcribeElem IdentAyudC, n, "ANTIMONIO"
    ExcribeElem IdentAyudC, n, "AMERICIO"
    ExcribeElem IdentAyudC, n, "AZUFRE"
    ExcribeElem IdentAyudC, n, "BARIO"
    ExcribeElem IdentAyudC, n, "BERKELIO"
    ExcribeElem IdentAyudC, n, "BERILIO"
    ExcribeElem IdentAyudC, n, "BORO"
    ExcribeElem IdentAyudC, n, "CADMIO"
    ExcribeElem IdentAyudC, n, "CALCIO"
    ExcribeElem IdentAyudC, n, "CARBONO"
    ExcribeElem IdentAyudC, n, "CLORO"
    ExcribeElem IdentAyudC, n, "ERBIO"
    ExcribeElem IdentAyudC, n, "ESTAÑO"
    ExcribeElem IdentAyudC, n, "FLUOR"
    ExcribeElem IdentAyudC, n, "FÓSFORO"
    ExcribeElem IdentAyudC, n, "HELIO"
    ExcribeElem IdentAyudC, n, "HIDRÓGENO"
    ExcribeElem IdentAyudC, n, "HIERRO"
    ExcribeElem IdentAyudC, n, "INDIO"
    ExcribeElem IdentAyudC, n, "ITERBIO"
    ExcribeElem IdentAyudC, n, "ITRIO"
    ExcribeElem IdentAyudC, n, "LITIO"
    ExcribeElem IdentAyudC, n, "LUTECIO"
    ExcribeElem IdentAyudC, n, "MAGNESIO"
    ExcribeElem IdentAyudC, n, "MERCURIO"
    ExcribeElem IdentAyudC, n, "MOLIBDENO"
    ExcribeElem IdentAyudC, n, "NEÓN"
    ExcribeElem IdentAyudC, n, "NITRÓGENO"
    ExcribeElem IdentAyudC, n, "ORO"
    ExcribeElem IdentAyudC, n, "OXÍGENO"
    ExcribeElem IdentAyudC, n, "PALADIO"
    ExcribeElem IdentAyudC, n, "PLATA"
    ExcribeElem IdentAyudC, n, "PLATINO"
    ExcribeElem IdentAyudC, n, "PLOMO"
    ExcribeElem IdentAyudC, n, "PLUTONIO"
    ExcribeElem IdentAyudC, n, "RADIO"
    ExcribeElem IdentAyudC, n, "RADÓN"
    ExcribeElem IdentAyudC, n, "SELENIO"
    ExcribeElem IdentAyudC, n, "SILICIO"
    ExcribeElem IdentAyudC, n, "SODIO"
    ExcribeElem IdentAyudC, n, "TANTALIO"
    ExcribeElem IdentAyudC, n, "TECNECIO"
    ExcribeElem IdentAyudC, n, "TITANIO"
    ExcribeElem IdentAyudC, n, "TUNGSTENO"
    ExcribeElem IdentAyudC, n, "URANIO"
    ExcribeElem IdentAyudC, n, "VANADIO"
    ExcribeElem IdentAyudC, n, "XENÓN"
    ExcribeElem IdentAyudC, n, "ZINC"
    ExcribeElem IdentAyudC, n, "ZIRCONIO"
    ExcribeElem IdentAyudC, n, "Lista de Elementos" & vbCrLf & _
                               "<elemento1>," & vbCrLf & _
                               "<elemento2> " & vbCrLf & _
                               "<elemento3> " & vbCrLf & _
                               "Fin Lista."
    ReDim Preserve IdentAyudC(n)    'Elimina valores no usados
End Sub



Private Sub lstCont_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Call FinAyudContextual
    End If
End Sub

Private Function anchoMaxLista() As Long
'Devuelve el ancho máximo del texto en la lista lstCont
Dim i As Integer
Dim anc As Single
Dim cad As String
Dim tx As SIZE
Dim lHDC As Long
    lHDC = GetDC(lstCont.hWnd)  'toma DC
    For i = 0 To lstCont.ListCount - 1
        cad = lstCont.List(i)
        Call GetTextExtentPoint32(lHDC, cad, Len(cad), tx)
        If tx.cx > anc Then anc = tx.cx
    Next
    anchoMaxLista = anc
End Function

Private Sub AbreAyudContextual(xtIni As Integer, ytIni As Long)
Const ALT_MIN_CON = 70     'Alto mínimo de menú contextual (aprox. 5 líneas)
Const ANC_MIN_CON = 100    'Ancho mínimo de menú contextual
Const ANC_MAX_CON = 200    'Ancho mínimo de menú contextual
    'Verifica si se puede iniciar ayuda contextual
    If xtIni < 1 Or ytIni < 1 Then Exit Sub
    If UserControl.Height < ALT_MIN_CON + altcarP Then
        Exit Sub
    End If
    If UserControl.Width < ANC_MIN_CON Then
        Exit Sub
    End If
    altMenCon = ALT_MIN_CON
    'calcula el ancho a mostrar de acuerdo a los ítems
    ancMenCon = anchoMaxLista()
    If ancMenCon < ANC_MIN_CON Then ancMenCon = ANC_MIN_CON
    If ancMenCon > ANC_MAX_CON Then ancMenCon = ANC_MAX_CON
    'Este truco se hace para ver cual es la mayor altura
    'del listbox que contiene líneas completas
    lstCont.Height = altMenCon    '
    altMenCon = lstCont.Height
    'guarda posición de inicio
    xtIniIden = xtIni
    ytIniIden = ytIni   'guarda posición de inicio
    'Calcula posición horizontal
    xCont0 = anccarP * (curXt - col1 + 1)
    If xCont0 + ancMenCon > UserControl.Width Then
        xCont0 = UserControl.Width - ancMenCon
    End If
    'Calcula posición vertical
    yCont0 = altcarP * (curYt - fil1 + 1)
    If yCont0 + altMenCon > UserControl.Height Then
        yCont0 = yCont0 - altMenCon
        If yCont0 < 0 Then Exit Sub     'Escapó de la pantalla
    End If
    '-------------- Finalmente activa la ayuda contextual ----------
    'Asigna propiedades
    lstCont.Left = xCont0
    lstCont.Top = yCont0
    'lstCont.Height = altMenCon 'ya se asignó
    lstCont.Width = ancMenCon
    lstCont.Visible = True
    HayAyudC = True     'Marca bandera
End Sub

Private Sub FinAyudContextual()
    'termina ayuda contextual
    lstCont.Visible = False
    HayAyudC = False    'Marca bandera
End Sub

Private Function AyudContextKeyDown(KeyCode As Integer, Shift As Integer) As Boolean
'Procesa el evento KeyDown para el menú de ayuda contextual
'Si la tecla ha sido procesada, devuelve verdadero
Dim i As Integer
    'Si no está activa la ayuda contextual, se sale
    If Not HayAyudC Then Exit Function
    i = lstCont.ListIndex
    'Identifica la tecla
    Select Case KeyCode
    'Teclas de desplazamiento
    Case 39 'flecha derecha
    Case 37 'flecha izquierda
    Case 40 'flecha abajo
        If i < lstCont.ListCount - 1 Then lstCont.ListIndex = i + 1
        AyudContextKeyDown = True
    Case 38 'flecha arriba
        If i = 0 Then   'Pasa a abajo
            If lstCont.ListCount > 0 Then lstCont.ListIndex = lstCont.ListCount - 1
        ElseIf i > 0 Then
            lstCont.ListIndex = i - 1
        End If
        AyudContextKeyDown = True
    Case 34  'pag.abajo
        If i < lstCont.ListCount - nFilAyudC Then
            lstCont.ListIndex = i + nFilAyudC
        Else
            If lstCont.ListCount > 0 Then lstCont.ListIndex = lstCont.ListCount - 1
        End If
        AyudContextKeyDown = True
    Case 33  'pag.arriba
        If i > nFilAyudC Then
            lstCont.ListIndex = i - nFilAyudC
        Else
            If lstCont.ListCount > 0 Then lstCont.ListIndex = 0
        End If
        AyudContextKeyDown = True
    Case 36  'inicio
    Case 35  'fin
    Case 46  'DEL
    Case 45  'INSERT
    Case 13
        AyudContextKeyDown = True   'Captura el enter
    End Select
End Function

Private Sub AyudContextKeyDown2(KeyCode As Integer, Shift As Integer)
'Procesa el evento KeyDown para el menú de ayuda contextual
'Debe llamarse después de que el editor ha procesado el evento, para tener
'el estado final del editor
Dim i As Integer
    'Si no está activa la ayuda contextual, se sale
    If Not HayAyudC Then Exit Sub
    'Verifica condiciones de término
    If curXt < xtIniIden Then
        Call FinAyudContextual
        Exit Sub
    End If
    If curYt <> ytIniIden Then
        Call FinAyudContextual
        Exit Sub
    End If
    i = lstCont.ListIndex
    'Identifica la tecla
    Select Case KeyCode
    'Teclas de desplazamiento
    Case 39 'flecha derecha
    Case 37 'flecha izquierda
    Case 36  'inicio
        AyudContextKeyPress 0   'Sólo Para validación
    Case 35  'fin
        AyudContextKeyPress 0   'Sólo Para validación
    Case 46  'DEL
        AyudContextKeyPress 0   'Sólo Para validación
    Case 45  'INSERT
        AyudContextKeyPress 0   'Sólo Para validación
    End Select
End Sub

Private Function CarAnterior(cad As String, pos As Integer) As String
'Devuelve el caracter anterior a una posición en una cadena
'Todos los caracteres se devuelven en mayúscula
    If pos > 1 Then
        CarAnterior = UCase(Mid$(cad, pos - 1, 1))
    Else
        CarAnterior = ""
    End If
End Function

Private Function CarActual(cad As String, pos As Integer) As String
'Devuelve el caracter actual en la posición en una cadena
'Todos los caracteres se devuelven en mayúscula
    If pos <= Len(cad) Then
        CarActual = UCase(Mid$(cad, pos, 1))
    Else
        CarActual = ""
    End If
End Function

Private Function LeePrimerosCar(xt As Integer, yt As Long) As String
'Obtiene la parte de una línea anterior a una posición.
'Se eliminan tabulaciones, espacios múltiples y se convierte a mayúscula
Dim p As Tpostex
Dim lin As String
    If xt <= 1 Then Exit Function
    lin = Left$(linrea(yt), xt - 1) 'toma parte anterios
    lin = Replace(lin, vbTab, " ")  'convierte tabulaciones
    lin = UCase(Trim(lin))          'convierte a mayúscula
    lin = Replace(lin, "  ", " ")   'elimina espacios múltiples, OJO!!!!, no es 100% seguro
    LeePrimerosCar = lin
End Function

Private Sub AyudContextKeyPress(KeyAscii As Integer)
'Procesa el evento KeyPress para el menú de ayuda contextual
Dim c As String
Dim xIniId As Integer   'Posición Inicial de Identificador
Dim lin As String       'Línea de trabajo
Dim iden As String      'Identificador
Dim i As Integer
Dim ncar As Long
    'Verifica condición de fin
    If KeyAscii = 27 Or KeyAscii = 32 Then
        If HayAyudC Then FinAyudContextual   '... Había
        Exit Sub
    End If
    If (KeyAscii = 9 Or KeyAscii = 13) And HayAyudC Then
        'Se acepta la opción del menú contextual
        i = lstCont.ListIndex   'lee opción
        If i <> -1 Then
            'Hay item seleccionado
            'Se selecciona el identificador a reemplazar
            ncar = PosXTreal(curXt, curYt) - xtIniIden
            SelecIdentificador xtIniIden, curYt, ncar
            CurInsertar lstCont.List(i)
        End If
        FinAyudContextual   'Termina la ayuda contextual
        Exit Sub
    End If
    'Procesa el comportamiento
    If HayAyudC Then     'Ya hay menú contextual...
        'toma identificador
        i = PosXTreal(curXt, curYt) 'Guarda posición actual
        If i <= xtIniIden Then  'Puede que se haya retrocedido
            FinAyudContextual
            Exit Sub
        End If
        iden = Mid$(linrea(ytIniIden), xtIniIden, i - xtIniIden)
        If Len(iden) < 1 Then
            FinAyudContextual   'Termina la ayuda contextual
            Exit Sub
        End If
    Else            'No había menú contextual
        'Verifica condiciones para activarlo
        c = UCase(Chr(KeyAscii))
        If c Like CAR_IDEN_VALM Then
            'Busca inicio de identificador
            lin = linrea(curYt) 'lee línea
            If CarActual(lin, curXt) Like CAR_IDEN_VALM Then
                Exit Sub    'Estamos en medio de un identificador
            End If
            i = PosXTreal(curXt, curYt) 'Guarda posición actual
            xIniId = i
            Do While xIniId > 1 And CarAnterior(lin, xIniId) Like CAR_IDEN_VALM
                xIniId = xIniId - 1
            Loop
            'verifica identificador previo
            iden = Mid$(linrea(curYt), xIniId, i - xIniId)
            'Verifica si se cumplen condiciones
            If Len(iden) < 1 Then Exit Sub
            'Llena la lista para verificar
            lin = LeePrimerosCar(xIniId, curYt)   'lee línea a comparar, por si la necesita
            
            'Llena lista de identificadores. Aquí pdoría elegirse la lista a usar, dependiendo
            'del contexto.
            Call LlenaIdenAyudContextual
            
            Call ListaIdenAyudCont(iden)
            If lstCont.ListCount = 0 Then Exit Sub  'No hay coincidencias
            'Inicia finalmente
            Call AbreAyudContextual(xIniId, curYt)
            Exit Sub 'y sale
        End If
    End If
    If Not HayAyudC Then Exit Sub    'No hay menú, salir
    Call ListaIdenAyudCont(iden)
    If lstCont.ListCount = 0 Then FinAyudContextual  'No hay coincidencias
End Sub

Private Sub ListaIdenAyudCont(iden As String)
'Genera una lista de identificadores similares a un identificador
Dim i As Integer
    lstCont.Clear
'    lstCont.AddItem iden
    iden = UCase(iden)
    For i = 1 To UBound(IdentAyudC)
        If UCase(IdentAyudC(i)) Like iden & "*" Then
            lstCont.AddItem IdentAyudC(i)
        End If
    Next
    If lstCont.ListCount > 0 Then
        lstCont.ListIndex = 0   'selecciona el primero
    End If
End Sub

'-----Eventos del desplazamiento
Private Sub VScroll1_Change()
Dim dy As Long
    If pintando Then Exit Sub   'Se está en medio de un redibujo
    dy = VScroll1.Value / facdesV - fil1
    FijaFil1 fil1 + dy
    Call Dibujar    'Dibuja el control
End Sub

Private Sub HScroll1_Change()
Dim dx As Integer
    If pintando Then Exit Sub   'Se está en medio de un redibujo
    dx = HScroll1.Value - col1
    FijaCol1 col1 + dx  'desplaza
    Call Dibujar    'Dibuja el control
End Sub

'*************************************************************************************
'****************** FUNCIONES GRÁFICAS Y DE DESPLAZAMIENTO DEL CURSOR ****************
'*************************************************************************************
Private Sub FijaLapiz(estilo As Long, ancho As Long, color As Long)
'Establece el lápiz actual de dibujo
Dim hdc As Long
    hdc = pic.hdc
    If hPen <> 0 Then DeleteObject hPen     'si ya hay un lápiz, lo elimina
    hPen = CreatePen(estilo, ancho, color)
    SelectObject hdc, hPen                  'queda pendiente eliminarlo
End Sub

Private Sub FijaRelleno(colorr As Long)
'Establece el relleno actual
Dim hdc As Long
    hdc = pic.hdc
    If hBrush <> 0 Then DeleteObject hBrush 'si hay relleno, lo elimina
    hBrush = CreateSolidBrush(colorr)
    SelectObject hdc, hBrush                'queda pendiente eliminarlo
End Sub

Private Sub FijaTexto(color As Long, tam As Long, nDegrees As Single, _
              Optional Letra As String = "Times New Roman", _
              Optional negrita As Boolean = False, _
              Optional cursiva As Boolean = False, _
              Optional subrayado As Boolean = False)
    pic.ForeColor = color
    pic.Font.SIZE = tam
    pic.Font.Name = Letra
End Sub

Private Sub DesplazaCursor(direccion As Integer, Optional paso As Long = 1, _
                           Optional actselec As Boolean = False)
'Mueve el cursor en la pantalla, considerando que deba caer en una posición
'válida. Desplaza y refresca la pantalla si es necesario
'Implementado a responder los desplazamientos por teclado. Se ha diseñado
'para que los rerfrescos de pantalla sean sólo cuando es necesario.
Dim posi As Integer 'posición inicial
Dim lin As String   'linea de trabajo
    Redibujar = False       'inicia para verificar si se debe redibujar
    'Verifica si escapa de la pantalla o de zona válida
    Select Case direccion
    Case DIR_IZQ
        If curXt = 1 And curYt > 1 Then   'Esta al inicio de línea y hay anterior
            tCursorA2 Len(linexp(curYt - 1)) + 1, curYt - 1, A_IZQ_TAB
        Else
            tCursorA2 curXt - paso, curYt, A_IZQ_TAB
        End If
        curXd = curXt     'actualiza posición deseada
    Case DIR_DER
        If curXt >= Len(linact) + 1 And curYt < nlin Then  'Final de línea y hay siguiente
            tCursorA2 1, curYt + 1, A_DER_TAB
        Else
            tCursorA2 curXt + paso, curYt, A_DER_TAB
        End If
        curXd = curXt     'actualiza posición deseada
    Case DIR_ARR
        If curXd <> curXt Then curXt = curXd  'intenta recuperar posición
        tCursorA2 curXt, curYt - paso, A_IZQ_TAB
    Case DIR_ABA
        If curXd <> curXt Then curXt = curXd  'intenta recuperar posición
        tCursorA2 curXt, curYt + paso, A_IZQ_TAB
    Case DIR_PARR   'Página arriba
        If curXd <> curXt Then curXt = curXd  'intenta recuperar posición
        VerticalScroll CInt(-paso)
        tCursorA2 curXt, curYt - paso, A_IZQ_TAB
    Case DIR_PABA   'Página abajo
        If curXd <> curXt Then curXt = curXd  'intenta recuperar posición
        VerticalScroll CInt(paso)
        tCursorA2 curXt, curYt + paso, A_IZQ_TAB
    Case DIR_ARRPAR
        If curXd <> curXt Then curXt = curXd  'intenta recuperar posición
        tCursorA2 curXt, curIniPar(), A_IZQ_TAB
    Case DIR_ABAPAR
        If curXd <> curXt Then curXt = curXd  'intenta recuperar posición
        tCursorA2 curXt, curFinPar(), A_IZQ_TAB
    Case DIR_INI    'lleva hasta el inicio de la línea
        tCursorA 1, curYt, A_IZQ_TAB
        curXd = curXt     'actualiza posición deseada
    Case DIR_FIN    'lleva hasta el final de la línea
        tCursorA maxTamLin + 1, curYt, A_DER_TAB
        curXd = curXt     'actualiza posición deseada
    Case DIR_HOM    'Al inicio del texto
        tCursorA 1, 1, A_IZQ_TAB
        curXd = curXt     'actualiza posición deseada
    Case DIR_END    'Al final del texto
        tCursorA2 maxTamLin + 1, nlin, A_IZQ_TAB
        curXd = curXt     'actualiza posición deseada
    Case DIR_IZQPAL
        If curXt = 1 And curYt > 1 Then   'Esta al inicio de línea y hay anterior
            tCursorA2 Len(linexp(curYt - 1)) + 1, curYt - 1, A_IZQ_TAB
        Else
            tCursorA2 curIniPal(), curYt, A_IZQ_TAB
        End If
        curXd = curXt     'actualiza posición deseada
    Case DIR_DERPAL 'A la derecha por palabra
        If curXt >= Len(linact) + 1 And curYt < nlin Then  'Final de línea y hay siguiente
            tCursorA2 1, curYt + 1, A_DER_TAB
        Else
            tCursorA2 curSigPal(), curYt, A_IZQ_TAB
        End If
        curXd = curXt      'actualiza posición deseada
    End Select
    'Procesa el dibujo de la selección
    If actselec Then    'Activar selección
        Call ExtenderSel
    Else
        If Redibujar Then   'Se movió la pantalla, redibujamos todo
            Call LimpSelec  'Aprovechamos para limpiar si hay selección
            Call Dibujar
        Else
            Call LimpSelec(True)      'Limpia selección
        End If
        'Marca la posición del cursor antes de una selección
        Call FijarSel0      'Fija punto base
        Call EncenCursor    'Enciende para que sea visible en la nueva posición
    End If
End Sub

Private Sub ExtenderSel()
'Refresca el texto del control considerando que la posición actual del cursor define
'un límite (inicial o final) en el bloque de selección. El otro límite es "sel0" que
'debe haber sido previamente fijado.
'Lo que se trata de hacer con este procedimiento es evitar llamar innecesariamente
'a Dibujar(). Pero si la variable "Redibujar" está en true, se redibuja todo
'incondicionalmente.
Dim Ptmp As Tpostex      'define el intervalo de la zona a dibujar
    If tipSelec = 1 Then Redibujar = True   'Fuerza actualización Total en este modo
    Call DesactivarCursor   'Se desactiva para no interferir
    haysel = True   'se marca la bandera
    If CurMenorPos(sel0) Then   '--------Cursor antes de sel0---------
        If MayorPos(sel2, sel0) Then    'Pasó de después a antes
            Ptmp = sel2     'guarda límite de dibujo
            sel1 = LeePosCur()    'Inicio de la selección está ahora en el cursor
            sel2 = sel0
            If Redibujar Then Call Dibujar Else Call DibTextPos(sel1, Ptmp)
        Else 'Selección sigue de Cursor a sel0 (sólo aumenta o disminuye)
            sel1 = LeePosCur()     'Inicio de la selección está ahora en el cursor
            sel2 = sel0
            'Verifica si se dibuja todo o sólo las líneas afectadas
            If Redibujar Then Call Dibujar Else Call DibTextPos(sel1, sel1ant)
        End If
    ElseIf CurMayorPos(sel0) Then   '--------Cursor después de sel0---------
        If MenorPos(sel1, sel0) Then 'Pasó de antes a después.
            Ptmp = sel1     'guarda límite de dibujo
            sel1 = sel0
            sel2 = LeePosCur()     'Fin de la selección está ahora en el cursor
            If Redibujar Then Call Dibujar Else Call DibTextPos(Ptmp, sel2)
        Else 'Selección sigue de sel0 a Cursor (sólo aumenta o disminuye)
            sel1 = sel0
            sel2 = LeePosCur()    'Fin de la selección está ahora en el cursor
            'Verifica si se dibuja todo o sólo las líneas afectadas
            If Redibujar Then Call Dibujar Else Call DibTextPos(sel2, sel2ant)
        End If
    Else            '--------Cursor en la misma posición---------
        'No hay nada que procesar, sólo ver si hay que borrar
        sel1 = sel0
        sel2 = sel0
        If Redibujar Then Call Dibujar Else Call DibTextPos(sel1ant, sel2ant)
    End If
    'actualiza anteriores
    sel1ant = sel1
    sel2ant = sel2
End Sub

Public Sub VerticalScroll(dy As Integer)
'Función pública para realizar un desplazamiento vertical de la pantalla.
'Se ha creado pensando en el desplazamiento con el evento "MouseWheel"
Dim vFinal As Integer
    vFinal = VScroll1.Value + dy
    If vFinal > VScroll1.Max Then vFinal = VScroll1.Max
    If vFinal < VScroll1.Min Then vFinal = VScroll1.Min
    VScroll1.Value = vFinal
End Sub
