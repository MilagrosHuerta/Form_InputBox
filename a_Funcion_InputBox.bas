Attribute VB_Name = "a_Funcion_InputBox"
' ------------------------------------------------------------ '
' ---                Funcion creada por                    --- '
' ---         MILAGROS HUERTA G?MEZ DE MERODIO             --- '
' ------------------------------------------------------------ '
' ---                   Form InputBox                      --- '
' ------------------------------------------------------------ '
' ---    Puedes usarla libremente en tus aplicaciones,     --- '
' ---    pero no asignarte la autor?a.                     --- '
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
