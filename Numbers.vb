Public class Numbers
    Private Function ValidarHora(hora As String) As String
        Dim horas As String
        Dim minutos As String
        fValidHora = ""

        If hora.Length > 5 Or Mid(hora, 3, 1) <> ":" Then Exit Function
        horas = Mid(hora, 1, InStr(hora, ":") - 1)
        minutos = Mid(hora, InStr(hora, ":") + 1, hora.Length - 1)
        If minutos = "" Then minutos = "00"

        If Not IsNumeric(horas) Or Not IsNumeric(minutos) Then Exit Function
        If horas = 0 Or Val(horas) >= 24 Or Val(minutos) >= 60 Then Exit Function
        ValidarHora = horas.PadLeft(2, "0") & ":" & minutos.PadLeft(2, "0")
    End Function
End Class
