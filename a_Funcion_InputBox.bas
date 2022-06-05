<<<<<<< HEAD
Attribute VB_Name = "a_Funcion_InputBox"
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
Public Nombre_Mensaje_I, Mensaje_Mostrar_I, Dato_Input As String
Public Titulo_Mensaje_I As Integer
Public Boton_A, Boton_C As String
Public continuar_I As Integer
Public numBotones_I As Byte
Public Posicion_Izda_I, Posicion_Top_I  As Integer
Public Valor_Dato As Variant
Function frmInputBox(texoMensaje As Variant, btnTexto As String, _
                   Optional datoEntrada As Variant, Optional tituloForm As String, _
                   Optional posLeftForm As Integer, Optional posTopForm As Integer, _
                   Optional textRigth As Boolean, Optional readRigth As Boolean)
' ------------------------------------------------------------------------------------------------- '
' --- Boton de Texto solo puede ser "bO" o "bOC", si se pone otro valor, toma "bOC" por defecto --- '
' ------------------------------------------------------------------------------------------------- '
Dim i, j, N As Integer
Dim longMensajeCorto, longMensajeMedio As Integer
    numBotones_I = Len(btnTexto) - 1
    longMensajeCorto = 160
    longMensajeMedio = 360
    If btnTexto <> "bO" And btnTexto <> "bOC" Then
        btnTexto = "bOC"
    End If

    Select Case btnTexto
    Case "bO"
        Boton_A = "ACEPTAR"
    Case "bOC"
        Boton_A = "Aceptar"
        Boton_C = "Cancelar"
    Case Else
        End
    End Select
    
    Nombre_Mensaje_I = tituloForm
    Dato_Input = datoEntrada
    Mensaje_Mostrar_I = texoMensaje
    Posicion_Izda_I = posLeftForm
    Posicion_Top_I = posTopForm
        
    Form_InputBox.Show

End Function
=======
Attribute VB_Name = "a_Funcion_InputBox"
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
Public Nombre_Mensaje, Mensaje_Mostrar, Nombre_Icono, Dato_Input As String
Public Boton_1, Boton_2, Boton_3 As String
Public continuar As Integer
Public numBotones As Byte
Public Posicion_Izda, Posicion_Top  As Integer
Public Valor_Dato As Variant
Public Titulo_Mensaje As Integer
Function frmInputBox(texoMensaje As Variant, btnTexto As String, _
                   Optional datoEntrada As Variant, Optional tituloForm As String, _
                   Optional posLeftForm As Integer, Optional posTopForm As Integer, _
                   Optional textRigth As Boolean, Optional readRigth As Boolean)
' ------------------------------------------------------------------------------------------------- '
' --- Boton de Texto solo puede ser "bO" o "bOC", si se pone otro valor, toma "bOC" por defecto --- '
' ------------------------------------------------------------------------------------------------- '
Dim i, j, N As Integer
Dim longMensajeCorto, longMensajeMedio As Integer
    numBotones = Len(btnTexto) - 1
    longMensajeCorto = 160
    longMensajeMedio = 360
    If btnTexto <> "bO" And btnTexto <> "bOC" Then
        btnTexto = "bOC"
    End If

    Select Case btnTexto
    Case "bO"
        Boton_1 = "Aceptar"
    Case "bOC"
        Boton_1 = "Aceptar"
        Boton_2 = "Cancelar"
    Case Else
        End
    End Select
    
    Nombre_Mensaje = tituloForm
    Dato_Input = datoEntrada
    Mensaje_Mostrar = texoMensaje
    Posicion_Izda = posLeftForm
    Posicion_Top = posTopForm
        
    Form_InputBox.Show

End Function
>>>>>>> 8aa7ce1eaa1a33a6d6aae9155f949b631c50ede7
