<<<<<<< HEAD
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_InputBox 
   Caption         =   "Form InputBox"
   ClientHeight    =   6684
   ClientLeft      =   4116
   ClientTop       =   4452
   ClientWidth     =   10800
   OleObjectBlob   =   "Form_InputBox.frx":0000
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "Form_InputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ------------------------------------------------------------ '
' ---                Funcion creada por                    --- '
' ---         MILAGROS HUERTA GÓMEZ DE MERODIO             --- '
' ------------------------------------------------------------ '
' ---                   Form InputBox                      --- '
' ------------------------------------------------------------ '
' ---    Puedes usarla libremente en tus aplicaciones,     --- '
' ---    pero no asignarte la autoría.                     --- '
' ---    Sirve para enviar mensajes con otro formato       --- '
' ---    y poder posicionarlo donde quieras                --- '
' ------------------------------------------------------------ '
Option Explicit
Private Sub UserForm_Initialize()
Dim a, b, i As Integer
    Form_InputBox.Caption = Nombre_Mensaje_I
    If Posicion_Izda_I = 0 And Posicion_Top_I = 0 Then
        Form_InputBox.StartUpPosition = 2     ' Centrar en pantalla
    Else
        Form_InputBox.Left = Posicion_Izda_I
        Form_InputBox.Top = Posicion_Top_I
    End If
' ----------------------------------------------------------- '
' --- Tamaño mensaje en funcion de la longitud del texto ---  '
' ----------------------------------------------------------- '
    If Len(Mensaje_Mostrar_I) < 200 Then
        i = 1
    ElseIf Len(Mensaje_Mostrar_I) < 300 Then
        i = 2
    ElseIf Len(Mensaje_Mostrar_I) < 400 Then
        i = 3
    ElseIf Len(Mensaje_Mostrar_I) < 500 Then
        i = 4
    ElseIf Len(Mensaje_Mostrar_I) < 600 Then
        i = 5
    Else
        i = 6
    End If
    a = 30 * (2 + i)
    b = 208 + 2 * 30 * i
    
    Form_InputBox.Height = b + 26
    frmMensaje_Mostrar.Height = a - 10
    frmEntrada_Dato.Top = a
    frmBoton_Aceptar.Top = a + 22
    frmBoton_Cancelar.Top = a + 22
    TextoBoton_Aceptar.Top = a + 23.5
    TextoBoton_Cancelar.Top = a + 23.5
    frmMensaje_Mostrar = Mensaje_Mostrar_I
' ----------------------------------------------------------------------- '
' --- Nombre de Botones y visibles en funcion de la variable btnTexto --- '
' ----------------------------------------------------------------------- '
'   TextoBoton_Aceptar.Text = Boton_A      ' Si quires cambiar el texto del botón
    frmBoton_Aceptar.Visible = True
    TextoBoton_Aceptar.Text = Boton_A

    If numBotones_I = 1 Then
        frmBoton_Cancelar.Visible = False
        TextoBoton_Cancelar.Visible = False
    ElseIf numBotones_I = 2 Then
'  TextoBoton_Cancelar.Text = Boton_C ' Si quires cambiar el texto del botón
        frmBoton_Cancelar.Visible = True
        TextoBoton_Cancelar.Visible = True
        TextoBoton_Cancelar.Text = Boton_C
    End If
    
    frmEntrada_Dato.Value = Dato_Input
End Sub
Private Sub frmEntrada_Dato_Change()
    Valor_Dato = Me.frmEntrada_Dato
End Sub
Private Sub frmBoton_Aceptar_Click()
    continuar_I = vbOK
    Unload Me
End Sub
Private Sub frmBoton_Cancelar_Click()
    continuar_I = vbCancel
    Unload Me
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' Macro para evitar que se pueda cerrar el formulario en la X de arriba a la derecha
    If CloseMode = 0 Then
        Cancel = True
    End If
End Sub
=======
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_InputBox 
   Caption         =   "Form InputBox"
   ClientHeight    =   6675
   ClientLeft      =   4110
   ClientTop       =   4455
   ClientWidth     =   10800
   OleObjectBlob   =   "Form_InputBox.frx":0000
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "Form_InputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ------------------------------------------------------------ '
' ---                Funcion creada por                    --- '
' ---         MILAGROS HUERTA GÓMEZ DE MERODIO             --- '
' ------------------------------------------------------------ '
' ---                   Form InputBox                      --- '
' ------------------------------------------------------------ '
' ---    Puedes usarla libremente en tus aplicaciones,     --- '
' ---    pero no asignarte la autoría.                     --- '
' ---    Sirve para enviar mensajes con otro formato       --- '
' ---    y poder posicionarlo donde quieras                --- '
' ------------------------------------------------------------ '
Option Explicit
Private Sub UserForm_Initialize()
Dim a, b, i As Integer
    Form_InputBox.Caption = Nombre_Mensaje
    If Posicion_Izda = 0 And Posicion_Top = 0 Then
        Form_InputBox.StartUpPosition = 2     ' Centrar en pantalla
    Else
        Form_InputBox.Left = Posicion_Izda
        Form_InputBox.Top = Posicion_Top
    End If
' ----------------------------------------------------------- '
' --- Tamaño mensaje en funcion de la longitud del texto ---  '
' ----------------------------------------------------------- '
    If Len(Mensaje_Mostrar) < 200 Then
        i = 1
    ElseIf Len(Mensaje_Mostrar) < 300 Then
        i = 2
    ElseIf Len(Mensaje_Mostrar) < 400 Then
        i = 3
    ElseIf Len(Mensaje_Mostrar) < 500 Then
        i = 4
    ElseIf Len(Mensaje_Mostrar) < 600 Then
        i = 5
    Else
        i = 6
    End If
    a = 30 * (2 + i)
    b = 208 + 2 * 30 * i
    
    Form_InputBox.Height = b + 26
    frmMensaje_Mostrar.Height = a - 10
    frmEntrada_Dato.Top = a
    frmBoton_Aceptar.Top = a + 22
    frmBoton_Cancelar.Top = a + 22
    TextoBoton_Aceptar.Top = a + 23.5
    TextoBoton_Cancelar.Top = a + 23.5
    frmMensaje_Mostrar = Mensaje_Mostrar
' ----------------------------------------------------------------------- '
' --- Nombre de Botones y visibles en funcion de la variable btnTexto --- '
' ----------------------------------------------------------------------- '
'   TextoBoton_Aceptar.Text = Boton_1      ' Si quires cambiar el texto del botón
    frmBoton_Aceptar.Visible = True

    If numBotones = 1 Then
        frmBoton_Cancelar.Visible = False
        TextoBoton_Cancelar.Visible = False
    ElseIf numBotones = 2 Then
'  TextoBoton_Cancelar.Text = Boton_2 ' Si quires cambiar el texto del botón
        frmBoton_Cancelar.Visible = True
        TextoBoton_Cancelar.Visible = True
    End If
    
    frmEntrada_Dato.Value = Dato_Input
End Sub
Private Sub frmEntrada_Dato_Change()
    Valor_Dato = Me.frmEntrada_Dato
End Sub
Private Sub frmBoton_Aceptar_Click()
    continuar = vbOK
    Unload Me
End Sub
Private Sub frmBoton_Cancelar_Click()
    continuar = vbCancel
    Unload Me
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' Macro para evitar que se pueda cerrar el formulario en la X de arriba a la derecha
    If CloseMode = 0 Then
        Cancel = True
    End If
End Sub
>>>>>>> 8aa7ce1eaa1a33a6d6aae9155f949b631c50ede7
