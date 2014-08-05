VERSION 5.00
Begin VB.Form frmPrincipal 
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   5865
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   8280
   ClipControls    =   0   'False
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VBEditor.ctlEdit editor 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      _ExtentX        =   12938
      _ExtentY        =   8493
   End
   Begin VB.Menu mn_archivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
      End
      Begin VB.Menu mnAbrir 
         Caption         =   "&Abrir ..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnGuardar 
         Caption         =   "&Guardar"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnGuardarComo 
         Caption         =   "G&uardar Como ..."
      End
      Begin VB.Menu mn_00 
         Caption         =   "-"
      End
      Begin VB.Menu mn_salir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnEdicion 
      Caption         =   "&Edición"
      Begin VB.Menu mnDeshacer 
         Caption         =   "Deshacer"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mntab00 
         Caption         =   "-"
      End
      Begin VB.Menu mnCortar 
         Caption         =   "Co&rtar"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnCopiar 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnPegar 
         Caption         =   "&Pegar"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnSelTodo 
         Caption         =   "&Seleccionar Todo"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnSep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnModColumna 
         Caption         =   "&Modo Columna"
      End
      Begin VB.Menu mnTabAEspacio 
         Caption         =   "&Tabulacion a espacios"
      End
   End
   Begin VB.Menu mnVer 
      Caption         =   "&Ver"
      Begin VB.Menu mnVerBarHor 
         Caption         =   "Barra de Despl. &Horiz."
      End
      Begin VB.Menu mnVerBarVer 
         Caption         =   "Barra de Despl. &Vert."
      End
      Begin VB.Menu mnVerNumLin 
         Caption         =   "&Número de Línea"
      End
      Begin VB.Menu mnSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnActualizar 
         Caption         =   "Actualizar &Pantalla"
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************
'                               FORMULARIO PRINCIPAL
'*************************************************************************************

Option Explicit

Private Sub editor_ArchivoSoltado(arc As String)
    'Se ha recibido un archivo soltado
    editor.CargarArch arc      'abre archivo
    Caption = NOMB_PROG & " - " & editor.archivo
End Sub

Private Sub editor_CambiaModo()
'El editor pide un cambio de modo con Alt-C
    Call mnModColumna_Click
End Sub

Private Sub Form_Load()
    Caption = NOMB_PROG
   
    'Inicia variables por defecto del programa
    ed_ver_hor = False
    ed_ver_ver = False
    ed_ver_num = False
    'Pone texto de ejemplo
    editor.Text = "Algunos ""elementos químicos"" son: " & vbCrLf & _
    "Hidrógeno, Nitrógeno, Oxígeno, y carbono." + vbCrLf + _
    "Algunos menos conocidos son:" + vbCrLf + _
    "cadmio y erbio." + vbCrLf + _
    "¿Cuál conoces?: "
    
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
    If ConsultaGuardar = vbCancel Then
        'Sólo si se canceló, se sale
        cancel = True
        Exit Sub
    End If
End Sub

Private Function ConsultaGuardar() As VbMsgBoxResult
Dim result As VbMsgBoxResult
Dim preg As String
    If editor.archivo = "" Then
        preg = "¿Guardar Cambios?"
    Else
        preg = "¿Guardar Cambios en " & editor.archivo & "?"
    End If
    If editor.TextModificado Then
        result = MsgBox(preg, vbYesNoCancel Or vbExclamation)
        ConsultaGuardar = result    'Devuelve respuesta
        If result = vbCancel Then Exit Function
        
        If result = vbYes Then
            'Guarda archivo
            If editor.archivo = "" Then
                Call mnGuardarComo_Click
                If editor.archivo = "" Then
                    'Aquí también se ha cancelado
                    ConsultaGuardar = vbCancel
                    Exit Function
                End If
            Else
                editor.GuardarArch
            End If
        End If
    End If
End Function

Private Sub Form_Resize()
    'Usa move() para cambiar de golpe la ubicación y no
    'llamar repetidamente a Resize() del editor
    editor.Move 0, 0, ScaleWidth, Posit(ScaleHeight)
End Sub

'*******************************************************************************
'********************************** MENÚ ARCHIVO *********************************
Private Sub mnNuevo_Click()
    If ConsultaGuardar = vbCancel Then Exit Sub
    
    editor.archivo = ""
    editor.Text = ""
    Caption = NOMB_PROG
End Sub

Private Sub mnAbrir_Click()
'Obtiene el nombre del archivo de trabajo
Dim nar As Integer
Dim linea As String
Dim arc As String
    If ConsultaGuardar = vbCancel Then Exit Sub
    
    arc = DialogoAbrir(frmPrincipal.hWnd, "Archivo", _
                  "*.txt", App.Path, "Abrir Archivo")
    If arc = "" Then Exit Sub
    If Dir(arc) = "" Then
        MsgBox "No se encuentra: " & arc
        Exit Sub
    End If
    editor.CargarArch arc     'abre archivo
    Caption = NOMB_PROG & " - " & editor.archivo
End Sub

Private Sub mnGuardar_Click()
    editor.GuardarArch
End Sub

Private Sub mnGuardarComo_Click()
Dim arc As String
    arc = DialogoGuardar(frmPrincipal.hWnd, "Archivo", _
                  "*.txt", App.Path, "Guardar Archivo")
    If arc = "" Then Exit Sub
    'abre la vista preliminar
    editor.GuardarArchComo arc
    Caption = NOMB_PROG & " - " & editor.archivo
End Sub

Private Sub mn_salir_Click()
    Unload Me
End Sub

'*****************************************************************************
'******************************** MENÚ EDICIÓN *******************************
Private Sub mnEdicion_click()
'Abre menú de edición
    'Actualiza el estado del deshacer
    If editor.nUndo = 0 Then
        mnDeshacer.Enabled = False
    Else
        mnDeshacer.Enabled = True
    End If
    'Actualiza modo del editor
    If editor.TipoSelec = 0 Then
        mnModColumna.Checked = False
    Else
        mnModColumna.Checked = True
    End If
End Sub

Private Sub mnDeshacer_Click()
    editor.Deshacer
End Sub

Private Sub mnCortar_Click()
    editor.CortaSeleccion
End Sub

Private Sub mnCopiar_Click()
    editor.CopiaSeleccion
End Sub

Private Sub mnPegar_Click()
    editor.PegaSeleccion
End Sub

Private Sub mnSelTodo_Click()
    editor.SeleccionaTodo
End Sub

Private Sub mnModColumna_Click()
'Activa o desactiva el modo columna
    If mnModColumna.Checked = True Then
        mnModColumna.Checked = False
        editor.TipoSelec = 0
    Else
        mnModColumna.Checked = True
        editor.TipoSelec = 1
    End If
    Call editor.Refrescar
End Sub

'****************************************************************************
'********************************** MENÚ VER*********************************
Private Sub mnVerBarHor_Click()
    If mnVerBarHor.Checked = True Then
        mnVerBarHor.Checked = False
        editor.verDesHor = False
    Else
        mnVerBarHor.Checked = True
        editor.verDesHor = True
    End If
    Call editor.Refrescar
End Sub

Private Sub mnVerBarVer_Click()
    If mnVerBarVer.Checked = True Then
        mnVerBarVer.Checked = False
        editor.verDesVer = False
    Else
        mnVerBarVer.Checked = True
        editor.verDesVer = True
    End If
    Call editor.Refrescar
End Sub

Private Sub mnVerNumLin_Click()
    If mnVerNumLin.Checked = True Then
        mnVerNumLin.Checked = False
        editor.verNumLin = False
    Else
        mnVerNumLin.Checked = True
        editor.verNumLin = True
    End If
    Call editor.Refrescar
End Sub

'****************************************************************************
'***************************** MENÚ HERRAMIENTAS ****************************
Private Sub mnActualizar_Click()
    Call editor.Refrescar
End Sub

Private Sub mnTabAEspacio_Click()
    editor.TabToSpaces
End Sub

