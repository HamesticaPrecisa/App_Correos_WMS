Imports System.Data
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports System
Imports System.IO
Imports System.Collections

Module ModPublico


    Function ArchivoAString(ByVal ruta As String) As String
        Dim objReader As New StreamReader(ruta)
        Dim sLine As String = ""
        Dim Texto As String = ""
        Do
            sLine = objReader.ReadLine()
            If Not sLine Is Nothing Then
                Texto = Texto + sLine + vbNewLine
            End If
        Loop Until sLine Is Nothing
        objReader.Close()
        Return Texto
    End Function
 
    Public Function devuelve_fecha2(ByVal fecha As DateTime) As String
        Dim a, m, d As String

        d = fecha.Day
        If Val(d) < 10 Then d = "0" & d
        m = fecha.Month
        If Val(m) < 10 Then m = "0" & m
        a = fecha.Year
        devuelve_fecha2 = d & "/" & m & "/" & a
        Return devuelve_fecha2
    End Function

    Public Function devuelve_fecha(ByVal fecha As DateTime) As String
        Dim a, m, d As String

        d = fecha.Day
        If Val(d) < 10 Then d = "0" & d
        m = fecha.Month
        If Val(m) < 10 Then m = "0" & m
        a = fecha.Year
        devuelve_fecha = a & "/" & m & "/" & d
        Return devuelve_fecha
    End Function

    Public Function RutDigito(ByVal Rut As String) As Boolean
        Try
            Dim ru, RU2 As Integer
            Dim Digito As Integer
            Dim Contador As Integer
            Dim Multiplo As Integer
            Dim Acumulador As Integer
            Dim r As String
            Rut = Rut.ToUpper()
            If Len(Rut) >= 9 Then
                ru = Mid(Rut, 1, Len(Rut) - 2)
                RU2 = ru
                Contador = 2
                Acumulador = 0
                While ru <> 0
                    Multiplo = (ru Mod 10) * Contador
                    Acumulador = Acumulador + Multiplo
                    ru = ru \ 10
                    Contador = Contador + 1
                    If Contador = 8 Then
                        Contador = 2
                    End If
                End While
                Digito = 11 - (Acumulador Mod 11)

                r = RU2 & "-" & CStr(Digito)
                If Digito = 10 Then r = RU2 & "-" & "K"
                If Digito = 11 Then r = RU2 & "-" & "0"

                If r = Rut Or "0" + r = Rut Then
                    RutDigito = True
                ElseIf CerosAnteriorString(r, 10) = Rut Then
                    RutDigito = True
                Else
                    RutDigito = False
                End If
            Else

                RutDigito = False
            End If
        Catch ex As Exception
            RutDigito = False
        End Try

    End Function

    Public Function FncConvierteNumero(ByVal mone As String) As Double
        Dim largo As Integer
        Dim cons, str As String

        cons = ""
        str = ""
        largo = Len(mone)
        For i = 1 To largo
            cons = Mid(mone, i, 1)
            If IsNumeric(cons) Then
                str = str & cons
            End If
        Next
        Return Val(str)
    End Function

    Public Sub SoloNumeros(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Char.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Sub LimpiarCajas(ByVal Objeto As Object)
        For Each c As Control In Objeto.Controls
            If TypeOf c Is TextBox Then
                c.Text = ""
            End If
        Next
    End Sub

    Sub LimpiarCajasMaskedTextBox(ByVal Objeto As Object)
        For Each c As Control In Objeto.Controls
            If TypeOf c Is MaskedTextBox Then
                c.Text = ""
            End If
        Next
    End Sub

    Public Function QuitarCaracteres(ByVal cadena As String, Optional ByVal chars As String = ".:<>{}[]^+,;_-/*?¿!$%&/¨Ññ()='áéíóúÁÉÍÓÚ¡" + Chr(34)) As String
        Dim i As Integer
        Dim nCadena As String
        'Asignamos valor a la cadena de trabajo para
        'no modificar la que envía el cliente.
        nCadena = cadena
        For i = 1 To Len(chars)
            nCadena = Replace(nCadena, Mid(chars, i, 1), "")
        Next i
        'Devolvemos la cadena tratada
        QuitarCaracteres = nCadena
        Return nCadena
    End Function

    Function CerosAnteriorString(ByVal numero As String, ByVal largo As Integer)

        Dim valorCeros As String = ""

        For i As Integer = 0 To ((largo - 1) - numero.Length)
            valorCeros = valorCeros + "0"
        Next

        Return valorCeros + numero
    End Function

    Function CerosPosteriorString(ByVal numero As String, ByVal largo As Integer)

        Dim valorCeros As String = ""

        For i As Integer = 0 To ((largo - 1) - numero.Length)
            valorCeros = valorCeros + "0"
        Next

        Return numero + valorCeros
    End Function

    Public Function validar_Mail(ByVal sMail As String) As Boolean
        ' retorna true o false   
        Return Regex.IsMatch(sMail, _
                "^([\w-]+\.)*?[\w-]+@[\w-]+\.([\w-]+\.)*?[\w]+$")
    End Function

    Public Sub ValidaTemperatura(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)

        If Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ChrW(189) Or e.KeyChar <> ChrW(109) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> ChrW(45) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Public Function ValidarHora(ByVal texto As String) As Boolean
        Dim ValidaHora As Boolean = IsDate(texto)
        Return ValidaHora
    End Function

    Public Function DevuelveHora()

        Dim hora As Date = buscaHoraServidor()
        Dim horaReturn As String = hora.ToString("HH:mm")
        Return horaReturn
    End Function

    Public Function GeneraNumeroDePallet(ByVal numero_Pallet As String) As String

        Dim sumapar As Integer = 0
        Dim sumaimp As Integer = 0

        Dim numeroVerificador As Integer = 0
        Dim impares As Integer = 0

        For i As Integer = 1 To Len(numero_Pallet)
            If i Mod 2 = 0 Then
                sumapar = sumapar + Val(Mid(numero_Pallet, i, 1))
            Else
                sumaimp = sumaimp + Val(Mid(numero_Pallet, i, 1))
            End If
        Next

        impares = (sumaimp * 3) + sumapar

        For i As Integer = 1 To 100
            If impares <= i * 10 Then
                Exit For
            End If
        Next i

        numeroVerificador = impares * 10 - impares

        Dim numero As String = numero_Pallet + numeroVerificador
        Return numero
    End Function

    Public Function EstadoCheckBox(ByVal check As Integer) As String
        Dim estado As String = "" + check.ToString()
        Return estado
    End Function

'12 Digitos + Verificador
    Function DigitoVerificadorEAN13(ByVal numero As String) As Integer


        Dim separado(12) As String
        Dim valores(12) As Integer

        Dim suma As Integer = 0

        For i As Integer = 0 To numero.Length - 1
            separado(i) = numero.Chars(i)
        Next


        For i As Integer = 0 To separado.Length - 1
            If i Mod 2 = 0 Then
                suma = Convert.ToInt32(suma + (Val(separado(i)) * 1))

            Else
                suma = Convert.ToInt32(suma + (Val(separado(i)) * 3))

            End If
        Next

        Dim multiplo As Integer = 0


        For i As Integer = 0 To suma Step 10

            multiplo = multiplo + 10
        Next


        Dim verificador As Integer = multiplo - suma
        If verificador = 10 Then
            verificador = 0
        End If

        Return verificador
    End Function

'17 Digitos + Verificador
    Function DigitoVerificadorEAN128UCC(ByVal numero As String) As String

        Dim separado(20) As String
        Dim valores(20) As Integer

        Dim suma As Integer = 0

        For i As Integer = 0 To numero.Length - 1
            separado(i) = numero.Chars(i)
        Next


        For i As Integer = 0 To separado.Length - 1
            If i Mod 2 = 1 Then
                suma = Convert.ToInt32(suma + (Val(separado(i)) * 1))
            Else
                suma = Convert.ToInt32(suma + (Val(separado(i)) * 3))
            End If
        Next

        Dim multiplo As Integer = 0


        For i As Integer = 0 To suma Step 10
            multiplo = multiplo + 10
        Next


        Dim verificador As Integer = multiplo - suma

        If verificador = 10 Then
            verificador = 0
        End If


        Return numero & "" & verificador.ToString()
    End Function

End Module
