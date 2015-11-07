Public Class frmRegistro
    Private Function ValidaString(ByVal cadena As String, ByVal logitudMaxima As Integer, ByVal Nulo As Boolean) As Integer
        Dim CodigoError As Integer = 0

        If cadena.Length > logitudMaxima Then
            CodigoError = 1
        End If

        If Not Nulo And cadena.Length < 1 Then
            CodigoError = 2
        End If

        Return CodigoError

    End Function
    ''' <summary>
    ''' Validar un formulario
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnRegistrar_Click(sender As Object, e As EventArgs) Handles btnRegistrar.Click
        Dim CadenaErrores As String = ""
        Dim Resultado As Integer = 0
        Dim TTip As ToolTip = New ToolTip()

        Resultado = ValidaString(tbNombre.Text, 100, False)
        If Resultado > 0 Then
            Select Case Resultado
                Case 1
                    CadenaErrores &= "El Campo Nombre no permite más de 100 carácteres." & vbNewLine
                Case 2
                    CadenaErrores &= "El Campo Nombre no puede estar vacío" & vbNewLine
            End Select
            tbNombre.BackColor = Color.LightCoral
            TTip.SetToolTip(tbNombre, CadenaErrores)
            ErrorProvider1.SetError(tbNombre, CadenaErrores)

        End If

        CadenaErrores = ""
        Resultado = ValidaString(tbApellidos.Text, 255, True)
        If Resultado > 0 Then
            Select Case Resultado
                Case 1
                    CadenaErrores &= "El Campo Nombre no permite más de 100 carácteres." & vbNewLine
                Case 2
                    CadenaErrores &= "El Campo Nombre no puede estar vacío" & vbNewLine
            End Select
            tbApellidos.BackColor = Color.LightCoral
            TTip.SetToolTip(tbApellidos, CadenaErrores)
            ErrorProvider1.SetError(tbApellidos, CadenaErrores)
        End If

        If CadenaErrores.Length > 0 Then
            MsgBox(CadenaErrores)
        Else

            MsgBox("Jugador insertado correctamente")
        End If
    End Sub

End Class
