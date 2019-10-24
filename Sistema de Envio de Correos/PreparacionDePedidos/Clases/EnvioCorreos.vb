Imports System.Net.NetworkInformation
Imports System.Text
Imports System.Net.Mail
Imports System.Net.Mime
Imports System.Net
Imports System.IO
Imports System.Reflection
Imports System.Collections.Generic
Imports System.Environment
Imports CrystalDecisions.Windows.Forms
Imports System.Configuration
Imports System.Security.Cryptography




Public Class EnvioCorreos
    'Dim correoprobar As String = "operaciones@precisafrozen.cl"
    Dim correoprobar As String = "hamestica@precisatech.cl"

    '--------------------------------------------------------------
    Dim correomostrar As String = "informaciones@precisafrozen.cl"
    Dim correoenvio As String = "informaciones@precisafrozen.cl"
    Dim claveenvio As String = "PrecisaFrozen1853"
    Dim puerto As String = "26"
    Dim host_mail As String = "mail.precisafrozen.cl"
    Dim setLogon As String = "sa" 'USUARIO DE BASE DE DATOS
    Dim setPass As String = "precisa" 'CLAVE DE BASE DE DATOS
    Dim estadoSSL As Boolean = False
    Dim correopedido As String = "pedido@precisafrozen.cl"
    Dim clavepedido As String = "pedido2013"

    'Dim correomostrar As String = "infop11@polartika.com"
    'Dim correoenvio As String = "infop11@polartika.com"
    'Dim claveenvio As String = "polartika11"
    'Dim puerto As String = "25"
    'Dim host_mail As String = "smtpout.secureserver.net"
    'Dim setLogon As String = "sa" 'USUARIO DE BASE DE DATOS
    'Dim setPass As String = "Clave01*" 'CLAVE DE BASE DE DATOS
    'Dim correopedido As String = correoenvio
    'Dim clavepedido As String = claveenvio
    'Dim estadoSSL As Boolean = False

    Public Function VerificaConexionInternet() As Boolean
        If My.Computer.Network.IsAvailable() Then
            Try

                If My.Computer.Network.Ping("www.google.cl", 1000) Then
                    Return True
                    Exit Function
                Else
                    Return False
                    Exit Function
                End If

            Catch ex As PingException
                Return False
                Exit Function
            End Try
        Else
            Return False
            Exit Function
        End If

        Return False
        Exit Function
    End Function

    Public Function probarSMTP()

        Dim smtp As New System.Net.Mail.SmtpClient
        Dim correo As New System.Net.Mail.MailMessage

        With smtp
            .Port = puerto
            .Host = host_mail
            .Credentials = New System.Net.NetworkCredential(correoenvio, claveenvio)
            .EnableSsl = estadoSSL
        End With

        With correo

            .From = New System.Net.Mail.MailAddress(correomostrar)
            .To.Add(correoprobar)
            .Subject = "Probando correo"
            .IsBodyHtml = True
            .Body = ""
        End With

        Try
            smtp.Send(correo)
            MessageBox.Show("Probado")
        Catch ex As SmtpException
            MessageBox.Show(ex.Message)
        End Try


        Return True
    End Function

    '-------------------------------
    'CORREO DE PEDIDOS -------------
    '-------------------------------

    Public Function EnviarCorreoPedidos(ByVal pedido_largo As String, ByVal pedido As String, ByVal cliente As String, ByVal fecha As String, ByVal hora As String, ByVal inf_cod As String) As Boolean

        Dim smtp As New System.Net.Mail.SmtpClient
        Dim correo As New System.Net.Mail.MailMessage
        Dim adjunto As System.Net.Mail.Attachment

        With smtp
            .Port = puerto
            .Host = host_mail
            .Credentials = New System.Net.NetworkCredential(correoenvio, claveenvio)
            .EnableSsl = estadoSSL
        End With

        With correo

            Dim sql As String = "SELECT cli_nomb, mail2 FROM clientes WHERE cli_rut='" + cliente.ToString.Remove(9, cliente.Length - 9) + "'"

            Dim tablaCliente As DataTable = ListarTablasSQL(sql)
            Dim NombreCliente As String = ""
            If tablaCliente.Rows.Count > 0 Then
                NombreCliente = tablaCliente.Rows(0)(0).ToString()
            End If

            .From = New System.Net.Mail.MailAddress(correomostrar)

            Dim uu As String = tablaCliente.Rows(0)(1).ToString().Trim

            Try
                If QuitarCaracteres(uu.ToString()).Length < uu.ToString().Length Then
                    Dim correo_electronico As String = ""
                    For i As Integer = 0 To uu.Length - 1
                        If uu.Chars(i) <> ";" Then
                            correo_electronico = correo_electronico + uu.Chars(i)
                        Else
                            .To.Add(correo_electronico)
                            correo_electronico = ""
                        End If
                    Next
                    .To.Add(correo_electronico)
                End If
            Catch ex As ArgumentException
                EnviarCorreoPedidos = False
            End Try

            'Llamada a funcion que carga copia a correo

            Dim sqlInterno As String = "SELECT inf_copia FROM informes WHERE inf_cod ='" + inf_cod + "'"

            Dim tablaInformes As DataTable = ListarTablasSQL(sqlInterno)
            Dim NombreInformes As String = ""
            If tablaInformes.Rows.Count > 0 Then

                Dim uuu As String = tablaInformes.Rows(0)(0).ToString().Trim

                If uuu <> "" Then
                    If QuitarCaracteres(uuu.ToString()).Length < uuu.ToString().Length Then
                        Dim correo_electronico As String = ""
                        For i As Integer = 0 To uuu.Length - 1
                            If uuu.Chars(i) <> ";" Then
                                correo_electronico = correo_electronico + uuu.Chars(i)
                            Else
                                .Bcc.Add(correo_electronico)
                                correo_electronico = ""
                            End If

                        Next
                        .Bcc.Add(correo_electronico)
                    End If
                End If
            End If

            .Subject = "Pedido Preparado"
            .IsBodyHtml = True

            Dim fechatermino As DateTime = Nothing
            Dim fechapedido As String = pedido_largo.Chars(6) + pedido_largo.Chars(7) + "/" + _
                pedido_largo.Chars(4) + pedido_largo.Chars(5) + "/" + pedido_largo.Chars(0) + pedido_largo.Chars(1) + _
                pedido_largo.Chars(2) + pedido_largo.Chars(3)

            Dim tabla As DataTable = ListarTablasSQL("SELECT fechaTermino FROM pedidos_ficha WHERE pedido = '" + pedido_largo.ToString() + "'")
            Try
                If tabla.Rows.Count > 0 Then
                    fechatermino = tabla.Rows(0)(0).ToString()
                Else
                    fechatermino = buscaHoraServidor()
                End If
            Catch ex As Exception
                fechatermino = buscaHoraServidor()
            End Try

            Dim f1 As String = ""
            Dim f2 As String = ""

            Dim d As String = DateDiff(DateInterval.Day, Convert.ToDateTime(fechapedido + " " + pedido_largo.Chars(8) + pedido_largo.Chars(9) + ":" + pedido_largo.Chars(10) + pedido_largo.Chars(11)), fechatermino)
            Dim h As String = DateDiff(DateInterval.Hour, (Convert.ToDateTime(fechapedido + " " + pedido_largo.Chars(8) + pedido_largo.Chars(9) + ":" + pedido_largo.Chars(10) + pedido_largo.Chars(11)).AddDays(d)), fechatermino)
            Dim m As String = DateDiff(DateInterval.Minute, (Convert.ToDateTime(fechapedido + " " + pedido_largo.Chars(8) + pedido_largo.Chars(9) + ":" + pedido_largo.Chars(10) + pedido_largo.Chars(11)).AddDays(d).AddHours(h)), fechatermino)

            Dim x As String = ""

            If d <> 0 Then
                If d = 1 Then
                    x = x + "" + d.ToString() + " dia    "
                Else
                    x = x + "" + d.ToString() + " dias   "
                End If
            End If

            If h <> 0 Then
                If h = 1 Then
                    x = x + "" + h.ToString() + " hora "
                Else
                    x = x + "" + h.ToString() + " horas "
                End If
            End If

            If m <> 0 Then
                If m = 1 Then
                    x = x + "" + m.ToString() + " minuto"
                Else
                    x = x + "" + m.ToString() + " minutos"
                End If
            End If

            .AlternateViews.Add(Body(pedido_largo.ToString(), fechapedido + " " + pedido_largo.Chars(8) + pedido_largo.Chars(9) + ":" + pedido_largo.Chars(10) + pedido_largo.Chars(11), fechatermino.ToString(), x, NombreCliente))
            .Priority = System.Net.Mail.MailPriority.Normal
            Dim ruta As String = Retorna_Ruta_ArchivoPedidos(pedido_largo, x)

            If ruta = "" Then
                Return True
                Exit Function
            End If

            adjunto = New System.Net.Mail.Attachment(ruta)
            .Attachments.Add(adjunto)

        End With
       
        Try
            Dim existe As String = "SELECT * FROM pedidos_enviados WHERE pedido='" + pedido_largo.ToString() + "'"

            Dim tabla As DataTable = ListarTablasSQL(existe)

            If tabla.Rows.Count > 0 Then
                Return True
                Exit Function
            End If

            smtp.Send(correo)

            Dim sql As String = "INSERT INTO pedidos_enviados " +
                                "VALUES ('" + pedido + "', " + "'" + pedido_largo + "','" + fecha + "', " + "'" + hora + "')"
            MovimientoSQL(sql)

            Dim actualiza As String = "UPDATE pedidos_ficha SET terminado = '2' " +
                                      " WHERE orden = '" + pedido + "'"
            MovimientoSQL(actualiza)

            EnviarCorreoPedidos = True

        Catch ex As Exception
            If (ex.Message.Trim = "No se puede enviar a un destinatario.") Then
                Dim sql As String = "INSERT INTO pedidos_enviados " +
                                "VALUES ('" + pedido + "', " + "'" + pedido_largo + "','" + fecha + "', " + "'" + hora + "')"
                MovimientoSQL(sql)

                Dim actualiza As String = "UPDATE pedidos_ficha SET terminado = '2' " +
                                          " WHERE orden = '" + pedido + "'"
                MovimientoSQL(actualiza)

                EnviarCorreoPedidos = True
            Else
                EnviarCorreoPedidos = False
            End If
        End Try

        Return EnviarCorreoPedidos

    End Function

    Private Function Body(ByVal pedido_largo As String, ByVal fecha_hora_solicitud As String, _
                          ByVal fecha_hora_termino As String, ByVal Tiempo_gestion As String, _
                          ByVal nombre_cliente As String) As AlternateView


        Dim archivo As String = ArchivoAString(Application.StartupPath + "\IndexPart1.txt")

        Dim editable = "<div id='contenedor'>" +
                        "<img src='cid:titulo'><br><br>" +
                        "<div id='cabecera'><br>" +
                        "<h1>Estimado Cliente " + nombre_cliente.ToUpper() + "," +
                        "<br><br>OrderBy informa que su pedido ya se encuentra preparado y listo para ser retirado, segun el " +
                        "siguiente detalle:</h1><br><br>" +
                        "</div><div id='solicitud'>" +
                        "<table border='0'>" +
                        "<tr>" +
                        "<td>Numero Solicitud Web</td>" +
                        "<td>" + pedido_largo + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Fecha y Hora de Solicitud</td>" +
                        "<td>" + Format(Convert.ToDateTime(fecha_hora_solicitud), "dd/MM/yyyy") + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + Format(Convert.ToDateTime(fecha_hora_solicitud), "HH:mm") + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Fecha y Hora de Termino de Preparacion</td>" +
                        "<td>" + Format(Convert.ToDateTime(fecha_hora_termino), "dd/MM/yyyy") + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + Format(Convert.ToDateTime(fecha_hora_termino), "HH:mm") + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Tiempo de Preparacion</td>" +
                        "<td><h3>" + Tiempo_gestion.ToString + "</h3></td>" +
                        "</tr>" +
                        "</table>"
        Dim archivo2 As String = ArchivoAString(Application.StartupPath + "\IndexPart2.txt")

        archivo = archivo + vbNewLine + editable + vbNewLine + archivo2

        Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim cabecera As LinkedResource = New LinkedResource(Application.StartupPath + "\titulo.jpg", MediaTypeNames.Image.Jpeg)
        cabecera.ContentId = "titulo"

        htmlView.LinkedResources.Add(cabecera)

        Return htmlView
    End Function

    Function Retorna_Ruta_ArchivoPedidos(ByVal pedido_largo As String, ByVal diferencia As String) As String

        Retorna_Ruta_ArchivoPedidos = ""

        Dim report As New Rpt_Pedido

        Try
            report.SetParameterValue("codigo", pedido_largo)
            report.SetParameterValue("ge", diferencia)
            report.SetDatabaseLogon(setLogon, setPass)
        Catch ex As CrystalReportsException
            Return ""
        End Try

        Try
            Dim CrExportOptions As ExportOptions
            Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
            Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()

            CrDiskFileDestinationOptions.DiskFileName = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\" + pedido_largo.ToString() + ".Pdf"
            CrExportOptions = report.ExportOptions
            With CrExportOptions
                .ExportDestinationType = ExportDestinationType.DiskFile
                .ExportFormatType = ExportFormatType.PortableDocFormat
                .DestinationOptions = CrDiskFileDestinationOptions
                .FormatOptions = CrFormatTypeOptions
            End With
            report.Export()
            report.Dispose()

            Return My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\" + pedido_largo.ToString() + ".Pdf"
            Exit Function
        Catch ex As Exception
            Return ""
            Exit Function
        End Try

    End Function

    Function Elimina_Archivo(ByVal pedido_largo As String) As Boolean
        Elimina_Archivo = False
        Try
            Dim ruta As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\" + pedido_largo + ".Pdf"
            If System.IO.File.Exists(ruta) = True Then
                Using FileShare = File.Open(ruta, FileMode.Open)
                    System.IO.File.Delete(ruta)
                    Elimina_Archivo = True
                End Using
            Else
                MsgBox("El archivo no fue encontrado")
            End If
        Catch ex As Exception
            Elimina_Archivo = False
        End Try
        Return Elimina_Archivo
    End Function

    Public Function EnviarCorreoSaldosIncorrectos(ByVal ID As String, ByVal FecPed As String, ByVal Ord As String, ByVal rutCli As String, ByVal nomCli As String, ByVal Pallet As String, ByVal CajsPick As String, ByVal SalInf As String, ByVal SalRea As String, ByVal Trackeo As String, ByVal Usu As String, ByVal FecPrep As String) As Boolean
        Dim smtp As New System.Net.Mail.SmtpClient
        Dim correo As New System.Net.Mail.MailMessage
        Dim adjunto As System.Net.Mail.Attachment

        With smtp
            .Port = puerto
            .Host = host_mail
            .Credentials = New System.Net.NetworkCredential(correoenvio, claveenvio)
            .EnableSsl = estadoSSL
        End With

        With correo
            .From = New System.Net.Mail.MailAddress(correomostrar)

            Dim sql As String = "SELECT inf_copia FROM informes WHERE inf_nom='Informe de Saldos Incorrectos Pedidos Locales'"

            Dim tablaCliente As DataTable = ListarTablasSQL(sql)

            Dim uu As String = tablaCliente.Rows(0)(0).ToString().Trim

            If QuitarCaracteres(uu.ToString()).Length < uu.ToString().Length Then
                Dim correo_electronico As String = ""
                For i As Integer = 0 To uu.Length - 1
                    If uu.Chars(i) <> ";" Then
                        correo_electronico = correo_electronico + uu.Chars(i)
                    Else
                        .To.Add(correo_electronico)
                        correo_electronico = ""
                    End If
                Next
                .To.Add(correo_electronico)
            End If

            .Subject = "Preparacion Pedido Local, Saldo Incorrecto. Orden: " & Ord
            .IsBodyHtml = True

            .AlternateViews.Add(BodySaldoIncorrecto(ID, FecPed, Ord, rutCli, nomCli, Pallet, CajsPick, SalInf, SalRea, Trackeo, Usu, FecPrep))
            .Priority = System.Net.Mail.MailPriority.Normal

        End With
        Threading.Thread.Sleep(10000)

        Try
            Try
                Try
                    smtp.Send(correo)

                    Dim sqlActualiza As String = "UPDATE Pedidos_Locales_Saldos_Incorrectos SET Mail_Enviado='1' WHERE ID='" + ID + "'"
                    MovimientoSQL(sqlActualiza)

                    EnviarCorreoSaldosIncorrectos = True
                Catch ex As Exception
                    EnviarCorreoSaldosIncorrectos = False
                End Try
            Catch ex As SmtpException
                EnviarCorreoSaldosIncorrectos = False
            End Try
        Catch ex As SmtpFailedRecipientException
            EnviarCorreoSaldosIncorrectos = False
        End Try

        Return EnviarCorreoSaldosIncorrectos
    End Function

    '-------------------------------
    'CORREO DE RECEPCION -----------
    '-------------------------------

    Public Function EnviarCorreoRecepcion(ByVal Codigo_Recepcion As String, ByVal rut_cliente As String, ByVal inf_cod As String) As Boolean
        Dim smtp As New System.Net.Mail.SmtpClient
        Dim correo As New System.Net.Mail.MailMessage
        Dim adjunto As System.Net.Mail.Attachment

        With smtp
            .Port = puerto
            .Host = host_mail
            .Credentials = New System.Net.NetworkCredential(correoenvio, claveenvio)
            .EnableSsl = estadoSSL
        End With

        With correo

            Dim sql As String = "SELECT cli_nomb, cli_cryd FROM clientes WHERE cli_rut='" + rut_cliente + "'"

            Dim tablaCliente As DataTable = ListarTablasSQL(sql)
            Dim NombreCliente As String = ""
            If tablaCliente.Rows.Count > 0 Then
                NombreCliente = tablaCliente.Rows(0)(0).ToString()
            End If

            .From = New System.Net.Mail.MailAddress(correomostrar)

            Dim uu As String = tablaCliente.Rows(0)(1).ToString().Trim

            If QuitarCaracteres(uu.ToString()).Length < uu.ToString().Length Then
                Dim correo_electronico As String = ""
                For i As Integer = 0 To uu.Length - 1
                    If uu.Chars(i) <> ";" Then
                        correo_electronico = correo_electronico + uu.Chars(i)
                    Else
                        .To.Add(correo_electronico)
                        correo_electronico = ""
                    End If
                Next
                .To.Add(correo_electronico)
            End If

            'Llamada a funcion que carga copia a correo

            Dim sqlInterno As String = "SELECT inf_copia FROM informes WHERE inf_cod ='" + inf_cod + "'"

            Dim tablaInformes As DataTable = ListarTablasSQL(sqlInterno)
            Dim NombreInformes As String = ""
            If tablaInformes.Rows.Count > 0 Then

                Dim uuu As String = tablaInformes.Rows(0)(0).ToString().Trim

                If uuu <> "" Then
                    If QuitarCaracteres(uuu.ToString()).Length < uuu.ToString().Length Then
                        Dim correo_electronico As String = ""
                        For i As Integer = 0 To uuu.Length - 1
                            If uuu.Chars(i) <> ";" Then
                                correo_electronico = correo_electronico + uuu.Chars(i)
                            Else
                                .Bcc.Add(correo_electronico)
                                correo_electronico = ""
                            End If

                        Next
                        .Bcc.Add(correo_electronico)
                    End If
                End If
            End If

            .Subject = "Recepción (" + Codigo_Recepcion + ") finalizada, transporte liberado"
            .IsBodyHtml = True

            Dim dt As DataTable = ListarTablasSQL("SELECT frec_rutcond, cho_nombre, cho_patente, cho_patente2 " +
                                                  "  FROM fichrece, choferes " +
                                                  " WHERE frec_rutcond = cho_rut AND frec_codi = '" + Codigo_Recepcion.ToString() + "'")
            If dt.Rows.Count = 0 Then
                Return True
                Exit Function
            End If

            .AlternateViews.Add(BodyRecepcion(dt.Rows(0)(0).ToString() + "    " + dt.Rows(0)(1).ToString(), dt.Rows(0)(2).ToString(), dt.Rows(0)(3).ToString(), NombreCliente))
            .Priority = System.Net.Mail.MailPriority.Normal

            Dim ruta As String = Retorna_Ruta_ArchivoRecepcion(Codigo_Recepcion)
            If ruta = "" Then
                Return True
                Exit Function
            End If
            adjunto = New System.Net.Mail.Attachment(ruta)
            .Attachments.Add(adjunto)

        End With
        Threading.Thread.Sleep(10000)

        Try
            Try
                Try
                    Dim Se_envio As String = " SELECT * FROM DocumentosEnviados WHERE Denv_seccion='006' AND Denv_Dcto='" + Codigo_Recepcion + "' "
                    Dim tabla As DataTable = ListarTablasSQL(Se_envio)
                    If tabla.Rows.Count > 0 Then
                        Return True
                        Exit Function
                    End If

                    smtp.Send(correo)

                    Dim sql As String = " INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora, Denv_Dcto)" +
                                        " VALUES ('006','" + devuelve_fecha(buscaHoraServidor()) + "','" + DevuelveHora() + "','" + Codigo_Recepcion + "')"
                    MovimientoSQL(sql)

                    Dim sqlActualiza As String = "UPDATE fichrece SET frec_enviada='1' , frec_modificado ='0' WHERE frec_codi='" + Codigo_Recepcion.ToString() + "'"
                    MovimientoSQL(sqlActualiza)

                    Dim sqlElimina As String = "DELETE FROM DocumentosNoEnviados WHERE ID_REGISTRO='" + Codigo_Recepcion.ToString() + "'"
                    MovimientoSQL(sqlElimina)

                    EnviarCorreoRecepcion = True
                Catch ex As Exception
                    If (ex.Message.Trim = "No se puede enviar a un destinatario.") Then
                        Dim sql As String = " INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora, Denv_Dcto)" +
                                        " VALUES ('006','" + devuelve_fecha(buscaHoraServidor()) + "','" + DevuelveHora() + "','" + Codigo_Recepcion + "')"
                        MovimientoSQL(sql)

                        Dim sqlActualiza As String = "UPDATE fichrece SET frec_enviada='1' , frec_modificado ='0' WHERE frec_codi='" + Codigo_Recepcion.ToString() + "'"
                        MovimientoSQL(sqlActualiza)

                        Dim sqlElimina As String = "DELETE FROM DocumentosNoEnviados WHERE ID_REGISTRO='" + Codigo_Recepcion.ToString() + "'"
                        MovimientoSQL(sqlElimina)

                        EnviarCorreoRecepcion = True
                    Else
                        Dim sqlNo As String = "INSERT INTO DocumentosNoEnviados(CODIGO,FECHA,ERROR) VALUES ('" + Codigo_Recepcion.ToString() + "',GETDATE(),'" + ex.ToString + "')"
                        MovimientoSQL(sqlNo)
                        EnviarCorreoRecepcion = False
                    End If
                End Try
            Catch ex As SmtpException
                If (ex.Message.Trim = "No se puede enviar a un destinatario.") Then
                    Dim sql As String = " INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora, Denv_Dcto)" +
                                        " VALUES ('006','" + devuelve_fecha(buscaHoraServidor()) + "','" + DevuelveHora() + "','" + Codigo_Recepcion + "')"
                    MovimientoSQL(sql)

                    Dim sqlActualiza As String = "UPDATE fichrece SET frec_enviada='1' , frec_modificado ='0' WHERE frec_codi='" + Codigo_Recepcion.ToString() + "'"
                    MovimientoSQL(sqlActualiza)

                    Dim sqlElimina As String = "DELETE FROM DocumentosNoEnviados WHERE ID_REGISTRO='" + Codigo_Recepcion.ToString() + "'"
                    MovimientoSQL(sqlElimina)

                    EnviarCorreoRecepcion = True
                Else
                    Dim sqlNo As String = "INSERT INTO DocumentosNoEnviados(CODIGO,FECHA,ERROR) VALUES ('" + Codigo_Recepcion.ToString() + "',GETDATE(),'" + ex.ToString + "')"
                    MovimientoSQL(sqlNo)
                    EnviarCorreoRecepcion = False
                End If
            End Try
        Catch ex As SmtpFailedRecipientException
            If (ex.Message.Trim = "No se puede enviar a un destinatario.") Then
                Dim sql As String = " INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora, Denv_Dcto)" +
                                    " VALUES ('006','" + devuelve_fecha(buscaHoraServidor()) + "','" + DevuelveHora() + "','" + Codigo_Recepcion + "')"
                MovimientoSQL(sql)

                Dim sqlActualiza As String = "UPDATE fichrece SET frec_enviada='1' , frec_modificado ='0' WHERE frec_codi='" + Codigo_Recepcion.ToString() + "'"
                MovimientoSQL(sqlActualiza)

                Dim sqlElimina As String = "DELETE FROM DocumentosNoEnviados WHERE ID_REGISTRO='" + Codigo_Recepcion.ToString() + "'"
                MovimientoSQL(sqlElimina)

                EnviarCorreoRecepcion = True
            Else
                Dim sqlNo As String = "INSERT INTO DocumentosNoEnviados(CODIGO,FECHA,ERROR) VALUES ('" + Codigo_Recepcion.ToString() + "',GETDATE(),'" + ex.ToString + "')"
                MovimientoSQL(sqlNo)
                EnviarCorreoRecepcion = False
            End If
        End Try


        Return EnviarCorreoRecepcion
    End Function

    Function Retorna_Ruta_ArchivoRecepcion(ByVal codigo As String) As String
        Dim valor As String = ""
        Dim Se_envio As String = " SELECT frec_modificado from fichrece where frec_codi='" + codigo.ToString() + "' "
        Dim tabla As DataTable = ListarTablasSQL(Se_envio)
        If tabla.Rows.Count > 0 Then
            valor = tabla.Rows(0)(0).ToString().Trim()
            If valor = "1" Then
                Dim sqlActualiza As String = "delete from DocumentosEnviados where Denv_Dcto='" + codigo.ToString() + "'"
                MovimientoSQL(sqlActualiza)
            Else
                Dim Se_envio1 As String = " SELECT Denv_Dcto FROM DocumentosEnviados WHERE Denv_seccion='006' AND Denv_Dcto='" + codigo.ToString + "' "
                Dim tabla1 As DataTable = ListarTablasSQL(Se_envio1)
                If tabla1.Rows.Count = 0 Then
                    valor = "1"
                End If
            End If
        End If

        If Not File.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Recepcion_" + codigo.ToString() + ".Pdf") Or valor = "1" Then

            Dim report As New Rpt_GuiaRecepcion

            Dim PictureBox1 As New PictureBox
            Dim PictureBox2 As New PictureBox
            Dim PictureBox3 As New PictureBox
            Dim PictureBox4 As New PictureBox

            Dim SqlImagen As String = "SELECT MAX(l.rimg_rececodi) id_pallets, " +
                                      "       (SELECT CONVERT(VARCHAR(MAX), CONVERT(VARBINARY(MAX),rimg_imagen2)) FROM receimagen WHERE rimg_rececodi = '" + codigo + "' AND rimg_num = 1) pic1, " +
                                      "       (SELECT CONVERT(VARCHAR(MAX), CONVERT(VARBINARY(MAX),rimg_imagen2)) FROM receimagen WHERE rimg_rececodi = '" + codigo + "' AND rimg_num = 2) pic2, " +
                                      "       (SELECT CONVERT(VARCHAR(MAX), CONVERT(VARBINARY(MAX),rimg_imagen2)) FROM receimagen WHERE rimg_rececodi = '" + codigo + "' AND rimg_num = 3) pic3, " +
                                      "       (SELECT CONVERT(VARCHAR(MAX), CONVERT(VARBINARY(MAX),rimg_imagen2)) FROM receimagen WHERE rimg_rececodi = '" + codigo + "' AND rimg_num = 4) pic4 " +
                                      "  FROM receimagen l " +
                                      " WHERE l.rimg_rececodi = '" + codigo + "'"

            Dim tablaimagen As DataTable = ListarTablasSQL(SqlImagen)

            If tablaimagen.Rows.Count > 0 Then

                'imagen1
                If tablaimagen.Rows(0)(1).ToString() <> "" Then

                    PictureBox1.Image = Base64ToImage(tablaimagen.Rows(0)(1).ToString())
                    PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage = True
                    PictureBox1.Size = New Size(640, 480)
                Else
                    PictureBox1.Image = My.Resources.blanco
                End If

                'imagen2
                If tablaimagen.Rows(0)(2).ToString() <> "" Then

                    PictureBox2.Image = Base64ToImage(tablaimagen.Rows(0)(2).ToString())
                    PictureBox2.SizeMode = PictureBoxSizeMode.StretchImage = True
                    PictureBox2.Size = New Size(640, 480)
                Else
                    PictureBox2.Image = My.Resources.blanco
                End If

                'imagen3
                If tablaimagen.Rows(0)(3).ToString() <> "" Then

                    PictureBox3.Image = Base64ToImage(tablaimagen.Rows(0)(3).ToString())
                    PictureBox3.SizeMode = PictureBoxSizeMode.StretchImage = True
                    PictureBox3.Size = New Size(640, 480)
                Else
                    PictureBox3.Image = My.Resources.blanco
                End If

                'imagen4
                If tablaimagen.Rows(0)(4).ToString() <> "" Then
                    PictureBox4.Image = Base64ToImage(tablaimagen.Rows(0)(4).ToString())
                    PictureBox4.SizeMode = PictureBoxSizeMode.StretchImage = True
                    PictureBox4.Size = New Size(640, 480)
                Else
                    PictureBox4.Image = My.Resources.blanco
                End If



            Else
                PictureBox1.Image = My.Resources.blanco
                PictureBox2.Image = My.Resources.blanco
                PictureBox3.Image = My.Resources.blanco
                PictureBox4.Image = My.Resources.blanco

            End If

            Dim Ds As New fotos

            Ds.fotografias.AddfotografiasRow(ImageToByte(PictureBox1.Image), ImageToByte(PictureBox2.Image), ImageToByte(PictureBox3.Image), ImageToByte(PictureBox4.Image))

            Try
                report.Subreports(1).SetDataSource(Ds)
                report.SetDatabaseLogon(setLogon, setPass)
                report.SetParameterValue("codigo", codigo)
            Catch ex As CrystalReportsException
                Dim sqlNo As String = "INSERT INTO DocumentosNoEnviados(CODIGO,FECHA,ERROR) VALUES ('" + codigo.ToString() + "',GETDATE(),'" + ex.ToString + "')"
                MovimientoSQL(sqlNo)
                Return ""
                Exit Function
            End Try
            Try
                Dim CrExportOptions As ExportOptions
                Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
                Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()

                CrDiskFileDestinationOptions.DiskFileName = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Recepcion_" + codigo.ToString() + ".Pdf"
                CrExportOptions = report.ExportOptions
                With CrExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.PortableDocFormat
                    .DestinationOptions = CrDiskFileDestinationOptions
                    .FormatOptions = CrFormatTypeOptions
                End With
                report.Export()
                report.Dispose()
                Return My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Recepcion_" + codigo.ToString() + ".Pdf"
                Exit Function
            Catch ex As Exception
                Try
                    Dim sqlNo As String = "INSERT INTO DocumentosNoEnviados(CODIGO,FECHA,ERROR) VALUES ('" + codigo.ToString() + "',GETDATE(),'" + ex.ToString + "')"
                    MovimientoSQL(sqlNo)
                Catch exx As Exception

                End Try

                Return ""
                Exit Function
            End Try
        End If
        Return ""
    End Function

    Private Function BodyRecepcion(ByVal Chofer As String, ByVal patente As String, ByVal Rampla As String, ByVal nombre_cliente As String) As AlternateView

        Dim archivo As String = ArchivoAString(Application.StartupPath + "\Recepcion1.txt")

        Dim Chofer_formato As String = Chofer.Chars(0) + Chofer.Chars(1) + "." + Chofer.Chars(2) + Chofer.Chars(3) + Chofer.Chars(4) + "." + _
            Chofer.Chars(5) + Chofer.Chars(6) + Chofer.Chars(7) + "-" + Chofer.Chars(8)

        Dim editable = "<h1>Estimado Cliente " + nombre_cliente.ToString() + ", <br><br>OrderBy informa que el camion enviado a <spa>DESCARGAR</spa> ya ha sido recepcionado y <span> Liberado!!!</span> de anden </h1><br><br></div>" +
                        "<div id='solicitud'><div>" +
                        "<table border='0'>" +
                        "<tr>" +
                        "<td>Chofer</td>" +
                        "<td>" + Chofer.Remove(0, 9).ToUpper() + "  " + Chofer_formato + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Patente</td>" +
                        "<td>" + patente.ToUpper() + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Rampla</td>" +
                        "<td>" + Rampla.ToUpper() + "</td>" +
                        "</tr>" +
                        "</table><br>"

        Dim archivo2 As String = ArchivoAString(Application.StartupPath + "\Recepcion2.txt")

        archivo = archivo + editable + archivo2

        Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim htmlView2 As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim cabecera As LinkedResource = New LinkedResource(Application.StartupPath + "\tituloRecepcion.jpg", MediaTypeNames.Image.Jpeg)
        cabecera.ContentId = "titulo"

        htmlView.LinkedResources.Add(cabecera)

        Return htmlView

    End Function

    Private Function BodySaldoIncorrecto(ByVal ID As String, ByVal FecPed As String, ByVal Ord As String, ByVal rutCli As String, ByVal nomCli As String, ByVal Pallet As String, ByVal CajsPick As String, ByVal SalInf As String, ByVal SalRea As String, ByVal Trackeo As String, ByVal Usu As String, ByVal FecPrep As String) As AlternateView

        Dim archivo As String = ArchivoAString(Application.StartupPath + "\Recepcion1.txt")

        Dim editable = "<h1 style='color:red;'>Se ha informado un Saldo Incorrecto al preparar Pedido Local <spa>Orden: " & Ord & "</spa></h1><br><br></div>" +
                        "<div id='solicitud'><div>" +
                        "<table border='0'>" +
                        "<tr>" +
                        "<td>Fecha Pedido</td>" +
                        "<td>" + FecPed + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Rut Cliente</td>" +
                        "<td>" + rutCli + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Cliente</td>" +
                        "<td>" + nomCli + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Soportante</td>" +
                        "<td>" + Pallet + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Cajas Pickeadas</td>" +
                        "<td>" + CajsPick + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Saldo Informado</td>" +
                        "<td>" + SalInf + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Saldo Según WMS</td>" +
                        "<td>" + SalRea + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Trackeo de Salida</td>" +
                        "<td>" + Trackeo + "</td>" +
                        "</tr>" +
                        "<td>Responsable</td>" +
                        "<td>" + Usu + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Fecha Preparación</td>" +
                        "<td>" + FecPrep + "</td>" +
                        "</tr>" +
                        "</table><br>"

        Dim archivo2 As String = ArchivoAString(Application.StartupPath + "\Recepcion2.txt")

        archivo = archivo + editable + archivo2

        Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim htmlView2 As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim cabecera As LinkedResource = New LinkedResource(Application.StartupPath + "\tituloPedidoWeb.jpg", MediaTypeNames.Image.Jpeg)
        cabecera.ContentId = "titulo"

        htmlView.LinkedResources.Add(cabecera)

        Return htmlView

    End Function

    '-------------------------------
    'CORREO DE PEDIDO WEB -----------
    '-------------------------------

    Public Function EnviarCorreoPedidoWeb(ByVal Orden As String, ByVal rut_cliente As String, ByVal inf_cod As String) As Boolean

        Dim smtp As New System.Net.Mail.SmtpClient
        Dim correo As New System.Net.Mail.MailMessage
        Dim adjunto As System.Net.Mail.Attachment

        With smtp
            .Port = puerto
            .Host = host_mail
            .Credentials = New System.Net.NetworkCredential(correopedido, clavepedido)
            .EnableSsl = estadoSSL
        End With

        Dim NombreCliente As String = ""

        With correo
            Dim sql As String = "SELECT cli_nomb, cli_mail FROM clientes WHERE cli_rut='" + rut_cliente + "'"

            Dim tablaCliente As DataTable = ListarTablasSQL(sql)

            If tablaCliente.Rows.Count > 0 Then
                NombreCliente = tablaCliente.Rows(0)(0).ToString()
            End If

            .From = New System.Net.Mail.MailAddress(correopedido)

            Dim uu As String = tablaCliente.Rows(0)(1).ToString().Trim

            If QuitarCaracteres(uu.ToString()).Length < uu.ToString().Length Then
                Dim correo_electronico As String = ""
                For i As Integer = 0 To uu.Length - 1
                    If uu.Chars(i) <> ";" Then
                        correo_electronico = correo_electronico + uu.Chars(i)
                    Else
                        .To.Add(correo_electronico)
                        correo_electronico = ""
                    End If
                Next
                .To.Add(correo_electronico)
            End If

            'Modificación HAmestica 26-07-19. Pedidos Exportación
            Dim sqlPedExp As String = "select ID=isnull(a.ID,0) from Pedidos_Tipo_Exportacion_Creados a with(nolock) where a.Orden_Pedido='" & Orden & "'"
            Dim dtPedExp As New DataTable
            dtPedExp = ListarTablasSQL(sqlPedExp)

            If (dtPedExp.Rows.Count > 0) Then
                Dim Resp As String = dtPedExp.Rows(0).Item(0).ToString.Trim
                If (Resp <> "0") Then
                    .To.Add("exportaciones@precisafrozen.cl")
                End If
            End If
            'Fin modificación HAmestica 26-07-19. Pedidos Exportación

            'Llamada a funcion que carga copia a correo
            Dim sqlInterno As String = "SELECT inf_copia FROM informes WHERE inf_cod ='" + inf_cod + "'"

            Dim tablaInformes As DataTable = ListarTablasSQL(sqlInterno)
            Dim NombreInformes As String = ""
            If tablaInformes.Rows.Count > 0 Then

                Dim uuu As String = tablaInformes.Rows(0)(0).ToString().Trim

                If uuu <> "" Then
                    If QuitarCaracteres(uuu.ToString()).Length < uuu.ToString().Length Then
                        Dim correo_electronico As String = ""
                        For i As Integer = 0 To uuu.Length - 1
                            If uuu.Chars(i) <> ";" Then
                                correo_electronico = correo_electronico + uuu.Chars(i)
                            Else
                                .To.Add(correo_electronico)
                                correo_electronico = ""
                            End If

                        Next
                        .To.Add(correo_electronico)
                    End If
                End If
            End If

            .Subject = "PEDIDO INTERNO N° " + Orden + " CLIENTE: " + NombreCliente + ""
            .IsBodyHtml = True

            .AlternateViews.Add(BodyPedidoWeb(Orden, NombreCliente))
            .Priority = System.Net.Mail.MailPriority.Normal

            Dim ruta As String = Retornar_Ruta_ArchivoPedidoWebExcel(Orden)
            If ruta = "" Then
                Return True
                Exit Function
            End If
            adjunto = New System.Net.Mail.Attachment(ruta)
            .Attachments.Add(adjunto)

        End With
        Threading.Thread.Sleep(10000)

        Try
            Dim Se_envio As String = " SELECT Denv_seccion FROM DocumentosEnviados WHERE Denv_seccion='005' AND Denv_Dcto='" + Orden + "' "
            Dim tabla As DataTable = ListarTablasSQL(Se_envio)
            If tabla.Rows.Count > 0 Then
                Return True
                Exit Function
            End If
            smtp.Send(correo)

            Dim sql As String = " INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora, Denv_Dcto)" +
                                " VALUES ('005','" + devuelve_fecha(buscaHoraServidor()) + "','" + DevuelveHora() + "','" + Orden + "')"
            MovimientoSQL(sql)

            Dim sqlActualiza As String = "UPDATE Pedidos_ficha SET ped_enviado='1' WHERE Orden='" + Orden.ToString() + "'"
            MovimientoSQL(sqlActualiza)

            EnviarCorreoPedidoWeb = True
        Catch ex As FileNotFoundException
            EnviarCorreoPedidoWeb = False
        Catch ex As Exception
            If (ex.Message.Trim = "No se puede enviar a un destinatario.") Then
                Dim sql As String = " INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora, Denv_Dcto)" +
                                " VALUES ('005','" + devuelve_fecha(buscaHoraServidor()) + "','" + DevuelveHora() + "','" + Orden + "')"
                MovimientoSQL(sql)

                Dim sqlActualiza As String = "UPDATE Pedidos_ficha SET ped_enviado='1' WHERE Orden='" + Orden.ToString() + "'"
                MovimientoSQL(sqlActualiza)

                EnviarCorreoPedidoWeb = True
            Else
                EnviarCorreoPedidoWeb = False
            End If
        End Try

        enivarAlertaPedidoInconsistente(Orden.ToString.Trim, NombreCliente)

        Return EnviarCorreoPedidoWeb
    End Function

    Sub enivarAlertaPedidoInconsistente(ByVal codigo As String, ByVal cliente As String)
        Try
            Dim cantidadPallet As Integer = 0

            Dim sqlVerifDet As String = "select CantDet=a.Cantidad_Pallets from Pedidos_Cantidad_Validacion_Correo a with(nolock) where a.Orden='" & codigo.ToString.Trim & "'"
            Dim dtVerifDet As New DataTable
            Dim CantDet As Integer = 0


            dtVerifDet = ListarTablasSQL(sqlVerifDet)
            CantDet = CInt(dtVerifDet.Rows(0).Item(0).ToString.Trim)

            Dim ExiErrDet As Boolean = False

            Dim intentos As Integer = 0

            For i = 1 To 3
                intentos += 1

                Dim sqlConsulta As String = "SELECT pallet,cajas,kilos,racd_ca,  " +
                          " racd_ba, racd_co, racd_pi, racd_ni, " +
                          " FechaVencimiento, FechaElaboracion, " +
                          " frec_guiades, LoteCliente, drec_codsag, " +
                          " observacion, mae_descr, drec_sopocli, Estado, cont_descr " +
                          " FROM  VW_PEDIDOS_DETALLE " +
                          " WHERE CAST(Orden AS varchar)='" + codigo.ToString() + "' " +
                          " ORDER BY racd_ca, racd_co ASC "

                Dim tablaConsulta As New DataTable

                tablaConsulta = ListarTablasSQL(sqlConsulta)

                If tablaConsulta.Rows.Count > 0 Then
                    cantidadPallet = tablaConsulta.Rows.Count()
                Else
                    cantidadPallet = 0
                End If

                If (cantidadPallet <> CantDet) Then
                    ExiErrDet = True
                Else
                    ExiErrDet = False
                End If

                If (ExiErrDet = False) Then
                    Exit For
                End If
            Next

            If (ExiErrDet) Then
                Dim smtp As New System.Net.Mail.SmtpClient
                Dim correo As New System.Net.Mail.MailMessage
                Dim adjunto As System.Net.Mail.Attachment

                With smtp
                    .Port = puerto
                    .Host = host_mail
                    .Credentials = New System.Net.NetworkCredential(correopedido, clavepedido)
                    .EnableSsl = estadoSSL
                End With

                With correo
                    .From = New System.Net.Mail.MailAddress(correopedido)
                    .To.Add("mlopez@precisaccgroup.cl")
                    .To.Add("hamestica@precisaccgroup.cl")
                    .Subject = "ALERTA MAIL INCONSISTENTE PEDIDO INTERNO N° " + codigo.ToString.Trim + ". CLIENTE: " + cliente + ""
                    .IsBodyHtml = True
                    .Priority = System.Net.Mail.MailPriority.Normal
                    .AlternateViews.Add(BodyPedidoWeb(codigo.ToString.Trim, cliente))
                    .Priority = System.Net.Mail.MailPriority.Normal

                    Dim ruta As String = Retornar_Ruta_ArchivoPedidoWebExcel(codigo.ToString.Trim)
                    If ruta = "" Then
                        Exit Sub
                    End If
                    adjunto = New System.Net.Mail.Attachment(ruta)
                    .Attachments.Add(adjunto)
                End With

                Threading.Thread.Sleep(10000)

                smtp.Send(correo)
            End If
        Catch ex As Exception

        End Try
    End Sub

    '-------------------------------
    'CORREO DE PEDIDO WEB < 24 Hrs--
    '-------------------------------

    Public Function EnviarCorreoPedidoWeb24Hrs(ByVal Orden As String, ByVal rut_cliente As String, ByVal inf_cod As String) As Boolean
        Dim smtp As New System.Net.Mail.SmtpClient
        Dim correo As New System.Net.Mail.MailMessage
        Dim adjunto As System.Net.Mail.Attachment

        With smtp
            .Port = puerto
            .Host = host_mail
            .Credentials = New System.Net.NetworkCredential(correopedido, clavepedido)
            .EnableSsl = estadoSSL
        End With

        With correo

            Dim sql As String = "SELECT cli_nomb, cli_mail FROM clientes WHERE cli_rut='" + rut_cliente + "'"

            Dim tablaCliente As DataTable = ListarTablasSQL(sql)
            Dim NombreCliente As String = ""
            If tablaCliente.Rows.Count > 0 Then
                NombreCliente = tablaCliente.Rows(0)(0).ToString()
            End If

            .From = New System.Net.Mail.MailAddress(correopedido)

            'Dim uu As String = tablaCliente.Rows(0)(1).ToString().Trim

            'If QuitarCaracteres(uu.ToString()).Length < uu.ToString().Length Then
            '    Dim correo_electronico As String = ""
            '    For i As Integer = 0 To uu.Length - 1
            '        If uu.Chars(i) <> ";" Then
            '            correo_electronico = correo_electronico + uu.Chars(i)
            '        Else
            '            .To.Add(correo_electronico)
            '            correo_electronico = ""
            '        End If
            '    Next
            '    .To.Add(correo_electronico)
            'End If

            'Llamada a funcion que carga copia a correo

            Dim sqlInterno As String = "SELECT inf_copia FROM informes WHERE inf_cod ='" + inf_cod + "'"

            Dim tablaInformes As DataTable = ListarTablasSQL(sqlInterno)
            Dim NombreInformes As String = ""
            If tablaInformes.Rows.Count > 0 Then

                Dim uuu As String = tablaInformes.Rows(0)(0).ToString().Trim

                If uuu <> "" Then
                    If QuitarCaracteres(uuu.ToString()).Length < uuu.ToString().Length Then
                        Dim correo_electronico As String = ""
                        For i As Integer = 0 To uuu.Length - 1
                            If uuu.Chars(i) <> ";" Then
                                correo_electronico = correo_electronico + uuu.Chars(i)
                            Else
                                .To.Add(correo_electronico)
                                correo_electronico = ""
                            End If

                        Next
                        .To.Add(correo_electronico)
                    End If
                End If
            End If

            .Subject = "PEDIDO INTERNO CON MENOS DE 24 HORAS N° " + Orden + " CLIENTE: " + NombreCliente + ""
            .IsBodyHtml = True

            .AlternateViews.Add(BodyPedidoWeb24Hrs(Orden, NombreCliente))
            .Priority = System.Net.Mail.MailPriority.Normal

            Dim ruta As String = Retornar_Ruta_ArchivoPedidoWebExcel(Orden)
            If ruta = "" Then
                Return True
                Exit Function
            End If
            adjunto = New System.Net.Mail.Attachment(ruta)
            .Attachments.Add(adjunto)

        End With
        Threading.Thread.Sleep(10000)

        Try
            Dim Se_envio As String = "SELECT Mail_Enviado FROM Pedidos_24_Horas WHERE Orden='" + Orden + "' "
            Dim tabla As DataTable = ListarTablasSQL(Se_envio)
            If tabla.Rows.Count > 0 Then
                If (tabla.Rows(0).Item(0).ToString.Trim = "1") Then
                    Return True
                    Exit Function
                End If
            End If
            smtp.Send(correo)

            Dim sqlActualiza As String = "UPDATE Pedidos_24_Horas SET Mail_Enviado='1' WHERE Orden='" + Orden.ToString() + "'"
            MovimientoSQL(sqlActualiza)

            EnviarCorreoPedidoWeb24Hrs = True
        Catch ex As FileNotFoundException
            EnviarCorreoPedidoWeb24Hrs = False
        Catch ex As Exception
            If (ex.Message.Trim = "No se puede enviar a un destinatario.") Then
                Dim sqlActualiza As String = "UPDATE Pedidos_24_Horas SET Mail_Enviado='1' WHERE Orden='" + Orden.ToString() + "'"
                MovimientoSQL(sqlActualiza)

                EnviarCorreoPedidoWeb24Hrs = True
            Else
                EnviarCorreoPedidoWeb24Hrs = False
            End If
        End Try

        Return EnviarCorreoPedidoWeb24Hrs
    End Function

    '-------------------------------
    'CORREO ALERTAS RFID
    '-------------------------------

    Public Sub EnviarCorreoAlertaRFID()
        Dim sqlDetAlert As String = "select a.ID_Alerta,a.mov_fecha_frm,a.mov_hora_frm,a.mov_folio,Ubicacion_Origen=c.NCamara,Ubicacion_Destino=a.NCamara,a.Responsable from V_RFID_Movimientos_Alertas a with(nolock) inner join V_RFID_Antenas_Camaras b with(nolock) on(a.mov_ca=b.Camara) inner join V_RFID_Antenas_Camaras c with(nolock) on(b.ID_Disp=c.ID_Disp and c.Tipo='E') where a.Mail_Enviado='0' order by CONVERT(date,a.mov_fecha),a.mov_hora_frm asc"
        Dim dtDetAlert As New DataTable
        dtDetAlert = ListarTablasSQL(sqlDetAlert)

        If (dtDetAlert.Rows.Count > 0) Then
            For j = 0 To dtDetAlert.Rows.Count - 1
                Dim IDAlert As String = dtDetAlert.Rows(j).Item("ID_Alerta").ToString.Trim
                Dim Fecha As String = dtDetAlert.Rows(j).Item("mov_fecha_frm").ToString.Trim.Substring(0, 10)
                Dim Hora As String = dtDetAlert.Rows(j).Item("mov_hora_frm").ToString.Trim
                Dim Pallet As String = dtDetAlert.Rows(j).Item("mov_folio").ToString.Trim
                Dim UbiOri As String = dtDetAlert.Rows(j).Item("Ubicacion_Origen").ToString.Trim
                Dim UbiDes As String = dtDetAlert.Rows(j).Item("Ubicacion_Destino").ToString.Trim
                Dim Responsable As String = dtDetAlert.Rows(j).Item("Responsable").ToString.Trim

                Dim smtp As New System.Net.Mail.SmtpClient
                Dim correo As New System.Net.Mail.MailMessage

                With smtp
                    .Port = puerto
                    .Host = host_mail
                    .Credentials = New System.Net.NetworkCredential(correopedido, clavepedido)
                    .EnableSsl = estadoSSL
                End With

                With correo
                    .From = New System.Net.Mail.MailAddress(correopedido)

                    Dim sqlInterno As String = "SELECT inf_copia FROM informes WHERE inf_nom='Alerta RFID'"
                    Dim tablaInformes As DataTable = ListarTablasSQL(sqlInterno)
                    Dim NombreInformes As String = ""

                    If tablaInformes.Rows.Count > 0 Then
                        Dim uuu As String = tablaInformes.Rows(0)(0).ToString().Trim

                        If uuu <> "" Then
                            If QuitarCaracteres(uuu.ToString()).Length < uuu.ToString().Length Then
                                Dim correo_electronico As String = ""
                                For i As Integer = 0 To uuu.Length - 1
                                    If uuu.Chars(i) <> ";" Then
                                        correo_electronico = correo_electronico + uuu.Chars(i)
                                    Else
                                        .To.Add(correo_electronico)
                                        correo_electronico = ""
                                    End If
                                Next
                                .To.Add(correo_electronico)
                            End If
                        End If
                    End If

                    .IsBodyHtml = True
                    .Priority = System.Net.Mail.MailPriority.High
                End With

                With correo
                    .Subject = "ALERTA RFID. MOVIMIENTO NO JUSTIFICADO. PALLET: " + Pallet + "."

                    Dim archivo As String = ArchivoAString(Application.StartupPath + "\PedidoWeb1.txt")
                    Dim editable = "<h1 style='color:red;'>Se ha detectado un movimiento RFID no justificado.</h1><br><br></div>" +
                        "<div id='solicitud'><div>" +
                        "<table border='0'>" +
                        "<tr>" +
                        "<td>Fecha</td>" +
                        "<td>" + Fecha + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Hora</td>" +
                        "<td>" + Hora + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Pallet</td>" +
                        "<td>" + Pallet + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Ubicación Origen</td>" +
                        "<td>" + UbiOri + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Ubicación Destino</td>" +
                        "<td>" + UbiDes + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Responsable</td>" +
                        "<td>" + Responsable + "</td>" +
                        "</tr>" +
                        "</table><br>"
                    Dim archivo2 As String = ArchivoAString(Application.StartupPath + "\PedidoWeb2.txt")
                    archivo = archivo + editable + archivo2

                    Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
                    Dim cabecera As LinkedResource = New LinkedResource(Application.StartupPath + "\tituloPedidoWeb.jpg", MediaTypeNames.Image.Jpeg)
                    cabecera.ContentId = "titulo"

                    htmlView.LinkedResources.Add(cabecera)

                    .AlternateViews.Add(htmlView)
                End With

                Threading.Thread.Sleep(10000)

                smtp.Send(correo)

                Dim sqlActualiza As String = "UPDATE RFID_Movimientos_Alertas SET Mail_Enviado='1' WHERE ID='" + IDAlert + "'"
                MovimientoSQL(sqlActualiza)
            Next
        End If
    End Sub

    Function Retorna_Ruta_ArchivoPedidoWeb(ByVal codigo As String) As String

        If Not File.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\PedidoWeb_" + codigo.ToString() + ".Pdf") Then

            Dim report As New Rpt_PedidosWeb
            Try

                report.SetParameterValue("codigo", codigo)
                report.SetDatabaseLogon(setLogon, setPass)
            Catch ex As CrystalReportsException
                Return ""
                Exit Function
            End Try

            Try

                Dim CrExportOptions As ExportOptions
                Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
                Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()

                CrDiskFileDestinationOptions.DiskFileName = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\PedidoWeb_" + codigo.ToString() + ".Pdf"
                CrExportOptions = report.ExportOptions
                With CrExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.PortableDocFormat
                    .DestinationOptions = CrDiskFileDestinationOptions
                    .FormatOptions = CrFormatTypeOptions
                End With
                report.Export()
                report.Dispose()
                Return My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\PedidoWeb_" + codigo.ToString() + ".Pdf"
                Exit Function
            Catch ex As FileNotFoundException
                Return ""
                Exit Function
            Catch ex As Exception
                Return ""
                Exit Function
            End Try
            Return ""
        End If

        Return ""
    End Function

    Public Function Retornar_Ruta_ArchivoPedidoWebExcel(ByVal codigo As String) As String
        Try
            If My.Computer.FileSystem.FileExists(My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\PedidoWeb_" + codigo.ToString() + ".xlsx") Then
                Return My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\PedidoWeb_" + codigo.ToString() + ".xlsx"
            Else
                Dim exApp As New Microsoft.Office.Interop.Excel.Application
                Dim exLibro As Microsoft.Office.Interop.Excel.Workbook
                Dim exHoja As Microsoft.Office.Interop.Excel.Worksheet
                Dim exHoja2 As Microsoft.Office.Interop.Excel.Worksheet
                Dim filacol As Integer = 0
                Dim totalkilos As Double = 0
                Dim akilos As Double = 0
                Dim totalcajas As Integer = 0
                Dim totalpallet As Integer = 0
                Dim ColorIndex As Integer = 40


                Dim sqlConsulta As String = ""
                Dim tablaConsulta As DataTable
                Dim cantidadPallet As Integer

                Dim sqlCabecera As String = ""
                Dim tablaCabecera As DataTable

                Dim sqlCajas As String = ""
                Dim tablaCajas As DataTable

                exLibro = exApp.Workbooks.Add


                '********************************************************************'
                '***************INICIO DE DETALLE DE CAJAS***************************'
                '********************************************************************'

                sqlCajas = " SELECT dpc_numpal,dpc_codcaja FROM detapedcaja INNER JOIN " +
                           " DetaReceCajas ON dpc_codcaja=Caj_cod  WHERE dpc_codped='" + codigo.ToString() + "' AND caj_ped='1' "
                tablaCajas = ListarTablasSQL(sqlCajas)

                exHoja2 = exLibro.ActiveSheet

                Dim ImageFileName2 As String = IO.Path.Combine(Application.StartupPath, "Imagen.JPG")
                My.Resources.precisa.Save(ImageFileName2)
                exHoja2.Shapes.AddPicture(ImageFileName2, Microsoft.Office.Core.MsoTriState.msoCTrue, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 0, 175, 52)
                exHoja2.Cells.Item(5, 1) = "DETALLE DE CAJAS"
                exHoja2.Cells.Item(7, 1) = "PALLET"
                exHoja2.Cells.Item(7, 2) = "CAJA"

                exHoja2.Cells.Item(5, 1).Interior.ColorIndex = ColorIndex
                exHoja2.Cells.Item(7, 1).Interior.ColorIndex = ColorIndex
                exHoja2.Cells.Item(7, 2).Interior.ColorIndex = ColorIndex

                For i = 0 To tablaCajas.Rows.Count - 1
                    filacol = i + 8
                    exHoja2.Cells.Item(filacol, 1) = "'" + tablaCajas.Rows(i)(0).ToString()
                    exHoja2.Cells.Item(filacol, 2) = "'" + tablaCajas.Rows(i)(1).ToString()
                Next

                exHoja2.Rows.Item(5).Font.Bold = 1
                exHoja2.Rows.Item(7).Font.Bold = 1
                exHoja2.Columns.AutoFit()
                '********************************************************************'
                '******************FIN DE DETALLE DE CAJAS***************************'
                '********************************************************************'

                '********************************************************************'
                '***************INICIO DE DETALLE DE SOPORTANTES*********************'
                '********************************************************************'
                sqlCabecera = " SELECT pedido,cliente,cli_nomb,transporte,fecha,hora,destino, " +
                              " observacion,ped_carga, " +
                              " FORMAT(CAST(fechapedido AS DATE), 'dd/MM/yyyy') as FECHA, " +
                              " SUBSTRING(CAST(CAST(fechapedido as time) AS VARCHAR),0,6) as HORA, ped_host " +
                              " FROM Pedidos_ficha,clientes " +
                              " WHERE CAST(Orden AS varchar)='" + codigo.ToString() + "' AND cliente = cli_rut "
                tablaCabecera = ListarTablasSQL(sqlCabecera)

                'Inicio validación Cantidad Pallets Pedido. 05-07-19
                'Dim sqlVerifDet As String = "select CantDet=count(a.pallet) from Pedidos_detalle a with(nolock) inner join Pedidos_ficha b with(nolock) on(a.pedido=b.pedido) where b.Orden='" & codigo.ToString.Trim & "'"
                Dim sqlVerifDet As String = "select CantDet=a.Cantidad_Pallets from Pedidos_Cantidad_Validacion_Correo a with(nolock) where a.Orden='" & codigo.ToString.Trim & "'"
                Dim dtVerifDet As New DataTable
                Dim CantDet As Integer = 0

                dtVerifDet = ListarTablasSQL(sqlVerifDet)
                If (dtVerifDet.Rows.Count > 0) Then
                    CantDet = CInt(dtVerifDet.Rows(0).Item(0).ToString.Trim)
                Else
                    CantDet = -1
                End If

                Dim ExiErrDet As Boolean = False

                Dim intentos As Integer = 0

                If (CantDet <> -1) Then
                    For i = 1 To 3
                        intentos += 1

                        sqlConsulta = "SELECT pallet,cajas,kilos,racd_ca,  " +
                                  " racd_ba, racd_co, racd_pi, racd_ni, " +
                                  " FechaVencimiento, FechaElaboracion, " +
                                  " frec_guiades, LoteCliente, drec_codsag, " +
                                  " observacion, mae_descr, drec_sopocli, Estado, cont_descr " +
                                  " FROM  VW_PEDIDOS_DETALLE " +
                                  " WHERE CAST(Orden AS varchar)='" + codigo.ToString() + "' " +
                                  " ORDER BY racd_ca, racd_co ASC "

                        tablaConsulta = ListarTablasSQL(sqlConsulta)

                        If tablaConsulta.Rows.Count > 0 Then
                            cantidadPallet = tablaConsulta.Rows.Count()
                        Else
                            cantidadPallet = 0
                        End If

                        If (cantidadPallet <> CantDet) Then
                            ExiErrDet = True
                        Else
                            ExiErrDet = False
                        End If

                        If (ExiErrDet = False) Then
                            Exit For
                        End If
                    Next
                Else
                    sqlConsulta = "SELECT pallet,cajas,kilos,racd_ca,  " +
                                  " racd_ba, racd_co, racd_pi, racd_ni, " +
                                  " FechaVencimiento, FechaElaboracion, " +
                                  " frec_guiades, LoteCliente, drec_codsag, " +
                                  " observacion, mae_descr, drec_sopocli, Estado, cont_descr " +
                                  " FROM  VW_PEDIDOS_DETALLE " +
                                  " WHERE CAST(Orden AS varchar)='" + codigo.ToString() + "' " +
                                  " ORDER BY racd_ca, racd_co ASC "

                    tablaConsulta = ListarTablasSQL(sqlConsulta)
                End If
                'Fin validación Cantidad Pallets Pedido. 05-07-19

                exHoja = exLibro.Worksheets.Add()

                Dim ImageFileName As String = IO.Path.Combine(Application.StartupPath, "Imagen.JPG")
                My.Resources.precisa.Save(ImageFileName)
                exHoja.Shapes.AddPicture(ImageFileName, Microsoft.Office.Core.MsoTriState.msoCTrue, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 0, 175, 52)

                exHoja.Range("D13", "H13").Merge(True)
                exHoja.Range("A5", "C5").Merge(True)
                exHoja.Range("A6", "C6").Merge(True)
                exHoja.Range("A7", "C7").Merge(True)
                exHoja.Range("A8", "C8").Merge(True)
                exHoja.Range("A9", "C9").Merge(True)
                exHoja.Range("A10", "C10").Merge(True)
                exHoja.Range("A11", "C11").Merge(True)
                exHoja.Range("A12", "C12").Merge(True)

                exHoja.Cells.Item(1, 1) = tablaCabecera.Rows(0)(11).ToString().ToUpper()
                'Inicio validación Cantidad Pallets Pedido. 05-07-19
                exHoja.Cells.Item(2, 1) = intentos
                'Fin validación Cantidad Pallets Pedido. 05-07-19
                exHoja.Cells.Item(5, 1) = "PLANILLA DE PEDIDO WEB"
                exHoja.Cells.Item(6, 1) = "CLIENTE: " + tablaCabecera.Rows(0)(1).ToString().ToUpper + "-" + tablaCabecera.Rows(0)(2).ToString().ToUpper
                exHoja.Cells.Item(7, 1) = "ORDEN: " + codigo.ToString().ToUpper()
                exHoja.Cells.Item(8, 1) = "N° PEDIDO: " + tablaCabecera.Rows(0)(0).ToString().ToUpper
                exHoja.Cells.Item(9, 1) = "CARGA: " + tablaCabecera.Rows(0)(8).ToString().ToUpper + "   FECHA: " + tablaCabecera.Rows(0)(4).ToString().ToUpper + "   HORA: " + tablaCabecera.Rows(0)(5).ToString().ToUpper
                exHoja.Cells.Item(10, 1) = "DESTINO: " + tablaCabecera.Rows(0)(6).ToString().ToUpper + " TRANSPORTE:  " + tablaCabecera.Rows(0)(3).ToString().ToUpper
                exHoja.Cells.Item(11, 1) = "INGRESO  " + "FECHA: " + tablaCabecera.Rows(0)(9).ToString().ToUpper + "   HORA: " + tablaCabecera.Rows(0)(10).ToString().ToUpper
                exHoja.Cells.Item(12, 1) = "OBSERVACIÓN: " + tablaCabecera.Rows(0)(7).ToString().ToUpper()

                exHoja.Cells.Item(14, 1) = "PALLET ORDERBY"
                exHoja.Cells.Item(14, 2) = "PALLET CLIENTE"
                exHoja.Cells.Item(14, 3) = "DESCRIPCIÓN DE PRODUCTO"
                exHoja.Cells.Item(13, 4) = "UBICACIÓN"
                exHoja.Cells.Item(14, 4) = "CA"
                exHoja.Cells.Item(14, 5) = "BA"
                exHoja.Cells.Item(14, 6) = "CO"
                exHoja.Cells.Item(14, 7) = "PI"
                exHoja.Cells.Item(14, 8) = "NI"
                exHoja.Cells.Item(14, 9) = "CAJAS"
                exHoja.Cells.Item(14, 10) = "KILOS"
                exHoja.Cells.Item(14, 11) = "F. ELABORACIÓN"
                exHoja.Cells.Item(14, 12) = "F. VENCIMIENTO"
                exHoja.Cells.Item(14, 13) = "CÓD. SAG"
                exHoja.Cells.Item(14, 14) = "LOTE"
                exHoja.Cells.Item(14, 15) = "GUÍA CLIENTE"
                exHoja.Cells.Item(14, 16) = "OBSERVACIÓN"
                exHoja.Cells.Item(14, 17) = "CONTRATO"
                exHoja.Cells.Item(14, 18) = "ESTADO"

                exHoja.Cells.Item(13, 1).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(13, 2).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(13, 3).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(13, 4).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(13, 5).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(13, 6).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(13, 7).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(13, 8).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(13, 9).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(13, 10).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(13, 11).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(13, 12).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(13, 13).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(13, 14).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(13, 15).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(13, 16).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(13, 17).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(13, 18).Interior.ColorIndex = ColorIndex

                exHoja.Cells.Item(14, 1).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(14, 2).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(14, 3).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(14, 4).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(14, 5).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(14, 6).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(14, 7).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(14, 8).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(14, 9).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(14, 10).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(14, 11).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(14, 12).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(14, 13).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(14, 14).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(14, 15).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(14, 16).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(14, 17).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(14, 18).Interior.ColorIndex = ColorIndex

                For i = 0 To tablaConsulta.Rows.Count - 1
                    akilos = tablaConsulta.Rows(i)(2).ToString()
                    filacol = i + 15
                    totalcajas = totalcajas + tablaConsulta.Rows(i)(1).ToString()
                    totalkilos = totalkilos + akilos
                    totalpallet = totalpallet + 1

                    exHoja.Cells.Item(filacol, 1) = "'" + tablaConsulta.Rows(i)(0).ToString()
                    exHoja.Cells.Item(filacol, 2) = "'" + tablaConsulta.Rows(i)(15).ToString()
                    exHoja.Cells.Item(filacol, 3) = tablaConsulta.Rows(i)(14).ToString()
                    exHoja.Cells.Item(filacol, 4) = tablaConsulta.Rows(i)(3).ToString()
                    exHoja.Cells.Item(filacol, 5) = tablaConsulta.Rows(i)(4).ToString()
                    exHoja.Cells.Item(filacol, 6) = tablaConsulta.Rows(i)(5).ToString()
                    exHoja.Cells.Item(filacol, 7) = tablaConsulta.Rows(i)(6).ToString()
                    exHoja.Cells.Item(filacol, 8) = tablaConsulta.Rows(i)(7).ToString()
                    exHoja.Cells.Item(filacol, 9) = tablaConsulta.Rows(i)(1).ToString()
                    exHoja.Cells.Item(filacol, 10) = akilos
                    exHoja.Cells.Item(filacol, 11) = "'" + tablaConsulta.Rows(i)(9).ToString()
                    exHoja.Cells.Item(filacol, 12) = "'" + tablaConsulta.Rows(i)(8).ToString()
                    exHoja.Cells.Item(filacol, 13) = tablaConsulta.Rows(i)(12).ToString()
                    exHoja.Cells.Item(filacol, 14) = tablaConsulta.Rows(i)(11).ToString()
                    exHoja.Cells.Item(filacol, 15) = tablaConsulta.Rows(i)(10).ToString()
                    exHoja.Cells.Item(filacol, 16) = tablaConsulta.Rows(i)(13).ToString()
                    exHoja.Cells.Item(filacol, 17) = tablaConsulta.Rows(i)(17).ToString()
                    exHoja.Cells.Item(filacol, 18) = tablaConsulta.Rows(i)(16).ToString()
                    exHoja.Rows.Item(filacol).HorizontalAlignment = 3
                    akilos = 0
                Next

                filacol = filacol + 1
                exHoja.Cells.Item(filacol, 1) = "TOTAL: "
                exHoja.Cells.Item(filacol, 2) = totalpallet
                exHoja.Cells.Item(filacol, 9) = totalcajas
                exHoja.Cells.Item(filacol, 10) = totalkilos
                exHoja.Rows.Item(filacol).HorizontalAlignment = 3
                exHoja.Rows.Item(filacol).Font.Bold = 1


                exHoja.Cells.Item(filacol, 1).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(filacol, 2).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(filacol, 3).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(filacol, 4).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(filacol, 5).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(filacol, 6).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(filacol, 7).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(filacol, 8).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(filacol, 9).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(filacol, 10).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(filacol, 11).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(filacol, 12).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(filacol, 13).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(filacol, 14).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(filacol, 15).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(filacol, 16).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(filacol, 17).Interior.ColorIndex = ColorIndex
                exHoja.Cells.Item(filacol, 18).Interior.ColorIndex = ColorIndex

                exHoja.Rows.Item(5).Font.Bold = 1
                exHoja.Rows.Item(13).Font.Bold = 1
                exHoja.Rows.Item(14).Font.Bold = 1

                exHoja.Rows.Item(13).HorizontalAlignment = 3
                exHoja.Rows.Item(14).HorizontalAlignment = 3

                'Inicio validación Cantidad Pallets Pedido. 05-07-19
                If (ExiErrDet) Then
                    filacol = filacol + 1
                    exHoja.Range("B" & filacol, "C" & filacol).Merge(True)
                    exHoja.Cells.Item(filacol, 2).Font.Color = Color.Red
                    exHoja.Cells.Item(filacol, 2) = "* Validar cantidad pallets con informática"
                End If
                'Fin validación Cantidad Pallets Pedido. 05-07-19

                exHoja.Columns.AutoFit()


                '********************************************************************'
                '*****************FIN DE DETALLE DE SOPORTANTES**********************'
                '********************************************************************'

                'CERRAR PROCESOS DE EXCEL'
                exApp.Application.ActiveWorkbook.SaveCopyAs(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\PedidoWeb_" + codigo.ToString() + ".xlsx")
                exHoja = Nothing
                exHoja2 = Nothing
                exLibro.Close(SaveChanges:=False)
                exLibro = Nothing
                exApp.Quit()
                exApp = Nothing

                'DESTRUIR PROCESOS DE EXCEL'
                Dim p As Process
                For Each p In Process.GetProcesses()
                    If Not p Is Nothing And p.ProcessName = "EXCEL" Then
                        p.Kill() 'Cierra el proceso
                    End If
                Next

                Return My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\PedidoWeb_" + codigo.ToString() + ".xlsx"
            End If
        Catch ex As FileNotFoundException
            Return ""
            Exit Function
        Catch ex As Exception
            Return ""
        End Try
        Return ""
    End Function

    Private Function BodyPedidoWeb(ByVal Orden As String, ByVal nombre_cliente As String) As AlternateView

        Dim archivo As String = ArchivoAString(Application.StartupPath + "\PedidoWeb1.txt")
        Dim hora As String = ""
        Dim fecha As String = ""
        Dim observacion As String = ""
        Dim transporte As String = ""
        Dim destino As String = ""
        Dim datos As String = " SELECT hora,fecha,observacion,transporte,destino FROM Pedidos_ficha WHERE Orden='" + Orden + "'"
        Dim tablaDatos As DataTable = ListarTablasSQL(datos)
        If tablaDatos.Rows.Count > 0 Then
            hora = tablaDatos.Rows(0)(0).ToString()
            fecha = tablaDatos.Rows(0)(1).ToString()
            observacion = tablaDatos.Rows(0)(2).ToString()
            transporte = tablaDatos.Rows(0)(3).ToString()
            destino = tablaDatos.Rows(0)(4).ToString()
        End If

        Dim editable = "<h1>Estimado Cliente " + nombre_cliente.ToString() + ", <br><br>OrderBy informa que el pedido N° <spa>" + Orden + "</spa> ha sido agendado</h1><br><br></div>" +
                        "<div id='solicitud'><div>" +
                        "<table border='0'>" +
                        "<tr>" +
                        "<td>Orden</td>" +
                        "<td>" + Orden + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Hora de carga</td>" +
                        "<td>" + hora + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Fecha de carga</td>" +
                        "<td>" + fecha + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Transporte</td>" +
                        "<td>" + transporte + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Destino</td>" +
                        "<td>" + destino + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Observacion</td>" +
                        "<td>" + observacion + "</td>" +
                        "</tr>" +
                        "</table><br>"

        Dim archivo2 As String = ArchivoAString(Application.StartupPath + "\PedidoWeb2.txt")

        archivo = archivo + editable + archivo2

        Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim htmlView2 As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim cabecera As LinkedResource = New LinkedResource(Application.StartupPath + "\tituloPedidoWeb.jpg", MediaTypeNames.Image.Jpeg)
        cabecera.ContentId = "titulo"

        htmlView.LinkedResources.Add(cabecera)

        Return htmlView

    End Function

    Private Function BodyPedidoWeb24Hrs(ByVal Orden As String, ByVal nombre_cliente As String) As AlternateView

        Dim archivo As String = ArchivoAString(Application.StartupPath + "\PedidoWeb1.txt")
        Dim hora As String = ""
        Dim fecha As String = ""
        Dim observacion As String = ""
        Dim transporte As String = ""
        Dim destino As String = ""
        Dim datos As String = " SELECT hora,fecha,observacion,transporte,destino FROM Pedidos_ficha WHERE Orden='" + Orden + "'"
        Dim tablaDatos As DataTable = ListarTablasSQL(datos)
        If tablaDatos.Rows.Count > 0 Then
            hora = tablaDatos.Rows(0)(0).ToString()
            fecha = tablaDatos.Rows(0)(1).ToString()
            observacion = tablaDatos.Rows(0)(2).ToString()
            transporte = tablaDatos.Rows(0)(3).ToString()
            destino = tablaDatos.Rows(0)(4).ToString()
        End If

        Dim editable = "<h1>Estimado Cliente " + nombre_cliente.ToString() + ", <br><br>OrderBy informa que el pedido N° <spa>" + Orden + "</spa> ha sido agendado</h1><br><br></div>" +
                        "<div id='solicitud'><div>" +
                        "<table border='0'>" +
                        "<tr>" +
                        "<td colspan='2' style='color:red;'>* Pedido realizado con menos de 24 horas</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Orden</td>" +
                        "<td>" + Orden + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Hora de carga</td>" +
                        "<td>" + hora + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Fecha de carga</td>" +
                        "<td>" + fecha + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Transporte</td>" +
                        "<td>" + transporte + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Destino</td>" +
                        "<td>" + destino + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Observacion</td>" +
                        "<td>" + observacion + "</td>" +
                        "</tr>" +
                        "</table><br>"

        Dim archivo2 As String = ArchivoAString(Application.StartupPath + "\PedidoWeb2.txt")

        archivo = archivo + editable + archivo2

        Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim htmlView2 As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim cabecera As LinkedResource = New LinkedResource(Application.StartupPath + "\tituloPedidoWeb.jpg", MediaTypeNames.Image.Jpeg)
        cabecera.ContentId = "titulo"

        htmlView.LinkedResources.Add(cabecera)

        Return htmlView

    End Function

    Private Function BodyStockComercialAgrosuper() As AlternateView
        'Dim FECHA As String = ""
        'Dim CODIGO_CLIENTE As String = ""
        'Dim PRODUCTO As String = ""
        'Dim LOTE As String = ""
        'Dim EMBARQUE As String = ""
        'Dim ENVASES As String = ""
        'Dim NETO As String = ""
        'Dim PLANTA As String = ""
        'Dim FOLIO_INTERNO As String = ""
        'Dim FECHA_RECEPCION As String = ""
        'Dim FECHA_PRODUCCION As String = ""
        'Dim STATUS_FRIGORIFICO As String = ""

        'Dim datos As String = "select FECHA=convert(varchar,getdate(),112),CODIGO_CLIENTE=convert(int,a.racd_codi),PRODUCTO=a.racd_codpro+' '+c.mae_descr,LOTE=replace(substring(a.LoteCliente,1,CHARINDEX('-',a.LoteCliente)),'-',''),EMBARQUE=right(rtrim(a.LoteCliente),1),ENVASES=a.racd_unidades,NETO=a.racd_peso,PLANTA='FX12',FOLIO_INTERNO=b.drec_sopocli,FECHA_RECEPCION=convert(varchar,convert(date,b.drec_fecrec),112),FECHA_PRODUCCION=convert(varchar,convert(date,a.racd_fecpro),112),STATUS_FRIGORIFICO=b.Estpallet from rackdeta a with(nolock) inner join detarece b with(nolock) on(a.racd_codi=b.drec_codi) inner join maeprod c with(nolock) on(a.racd_codpro=c.mae_codi) where b.drec_rutcli='78408440K'"
        'Dim tablaDatos As DataTable = ListarTablasSQL(datos)

        'If tablaDatos.Rows.Count > 0 Then
        '    FECHA = tablaDatos.Rows(0)(0).ToString()
        '    CODIGO_CLIENTE = tablaDatos.Rows(0)(1).ToString()
        '    PRODUCTO = tablaDatos.Rows(0)(2).ToString()
        '    LOTE = tablaDatos.Rows(0)(3).ToString()
        '    EMBARQUE = tablaDatos.Rows(0)(4).ToString()
        '    ENVASES = tablaDatos.Rows(0)(5).ToString()
        '    NETO = tablaDatos.Rows(0)(6).ToString()
        '    PLANTA = tablaDatos.Rows(0)(7).ToString()
        '    FOLIO_INTERNO = tablaDatos.Rows(0)(8).ToString()
        '    FECHA_RECEPCION = tablaDatos.Rows(0)(9).ToString()
        '    FECHA_PRODUCCION = tablaDatos.Rows(0)(10).ToString()
        '    STATUS_FRIGORIFICO = tablaDatos.Rows(0)(11).ToString()
        'End If

        Dim archivo As String = ArchivoAString(Application.StartupPath + "\PedidoWeb1.txt")

        Dim editable = "<h1>Estimado Cliente, <br><br>OrderBy informa el stock de sus productos en nuestras dependencias</h1><br><br></div>"

        Dim archivo2 As String = ArchivoAString(Application.StartupPath + "\PedidoWeb2.txt")

        archivo = archivo + editable + archivo2

        Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim htmlView2 As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim cabecera As LinkedResource = New LinkedResource(Application.StartupPath + "\tituloPedidoWeb.jpg", MediaTypeNames.Image.Jpeg)
        cabecera.ContentId = "titulo"

        htmlView.LinkedResources.Add(cabecera)

        Return htmlView

    End Function

    '-------------------------------
    'CORREO CHECKLIST  -------------
    '-------------------------------
    Public Function EnviarCorreoChecklist(ByVal codigo_checklist As String, ByVal cliente As String, ByVal inf_cod As String) As Boolean

        Dim smtp As New System.Net.Mail.SmtpClient
        Dim correo As New System.Net.Mail.MailMessage
        Dim adjunto As System.Net.Mail.Attachment

        With smtp
            .Port = puerto
            .Host = host_mail
            .Credentials = New System.Net.NetworkCredential(correoenvio, claveenvio)
            .EnableSsl = estadoSSL
        End With

        With correo

            Dim sql As String = "SELECT cli_nomb, cli_cryd FROM clientes WHERE cli_rut='" + cliente + "'"

            Dim tablaCliente As DataTable = ListarTablasSQL(sql)
            Dim NombreCliente As String = ""
            If tablaCliente.Rows.Count > 0 Then
                NombreCliente = tablaCliente.Rows(0)(0).ToString()
            End If

            .From = New System.Net.Mail.MailAddress(correomostrar)

            Dim uu As String = tablaCliente.Rows(0)(1).ToString().Trim

            If QuitarCaracteres(uu.ToString()).Length < uu.ToString().Length Then
                Dim correo_electronico As String = ""
                For i As Integer = 0 To uu.Length - 1
                    If uu.Chars(i) <> ";" Then
                        correo_electronico = correo_electronico + uu.Chars(i)
                    Else
                        .To.Add(correo_electronico)
                        correo_electronico = ""
                    End If

                Next
                .To.Add(correo_electronico)
            End If

            'Llamada a funcion que carga copia a correo

            Dim sqlInterno As String = "SELECT inf_copia FROM informes WHERE inf_cod ='" + inf_cod + "'"

            Dim tablaInformes As DataTable = ListarTablasSQL(sqlInterno)
            Dim NombreInformes As String = ""
            If tablaInformes.Rows.Count > 0 Then

                Dim uuu As String = tablaInformes.Rows(0)(0).ToString().Trim

                If uuu <> "" Then
                    If QuitarCaracteres(uuu.ToString()).Length < uuu.ToString().Length Then
                        Dim correo_electronico As String = ""
                        For i As Integer = 0 To uuu.Length - 1
                            If uuu.Chars(i) <> ";" Then
                                correo_electronico = correo_electronico + uuu.Chars(i)
                            Else
                                .Bcc.Add(correo_electronico)
                                correo_electronico = ""
                            End If

                        Next
                        .Bcc.Add(correo_electronico)
                    End If
                End If
            End If

            .Subject = "Informe de salida de camion de la planta (" + codigo_checklist + ")"
            .IsBodyHtml = True

            Dim dt As DataTable = ListarTablasSQL(" SELECT Cl_chorut, cho_nombre, cho_patente, cho_patente2 " +
                                                  "   FROM zCheckList, choferes  " +
                                                  "  WHERE Cl_chorut = cho_rut AND cl_fol ='" + codigo_checklist + "'")


            .AlternateViews.Add(BodyChecklist(dt.Rows(0)(0).ToString() + "    " + dt.Rows(0)(1).ToString(), dt.Rows(0)(2).ToString(), dt.Rows(0)(3).ToString(), NombreCliente))
            .Priority = System.Net.Mail.MailPriority.Normal
            Dim ruta As String = Retorna_Ruta_Checklist(codigo_checklist)
            If ruta = "" Then
                Return True
                Exit Function
            End If
            adjunto = New System.Net.Mail.Attachment(ruta)
            .Attachments.Add(adjunto)
        End With

        Try
            Dim Se_envio As String = " SELECT * FROM DocumentosEnviados WHERE Denv_seccion='100' AND Denv_Dcto='" + codigo_checklist + "' "
            Dim tabla As DataTable = ListarTablasSQL(Se_envio)
            If tabla.Rows.Count > 0 Then
                Dim sql_act As String = " UPDATE zchecklist SET cl_enviada = '1' WHERE cl_fol='" + codigo_checklist + "'"
                MovimientoSQL(sql_act)

                Return True
                Exit Function
            End If

            smtp.Send(correo)

            Dim sql As String = "INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora, Denv_Dcto)" +
                                "VALUES ('100','" + devuelve_fecha(buscaHoraServidor()) + "','" + DevuelveHora() + "','" + codigo_checklist + "')"

            MovimientoSQL(sql)

            Dim sql_actualiza As String = "UPDATE zchecklist SET cl_enviada='1' WHERE cl_fol='" + codigo_checklist + "'"
            MovimientoSQL(sql_actualiza)

            EnviarCorreoChecklist = True
        Catch ex As Exception
            If (ex.Message.Trim = "No se puede enviar a un destinatario.") Then
                Dim sql As String = "INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora, Denv_Dcto)" +
                                "VALUES ('100','" + devuelve_fecha(buscaHoraServidor()) + "','" + DevuelveHora() + "','" + codigo_checklist + "')"

                MovimientoSQL(sql)

                Dim sql_actualiza As String = "UPDATE zchecklist SET cl_enviada='1' WHERE cl_fol='" + codigo_checklist + "'"
                MovimientoSQL(sql_actualiza)

                EnviarCorreoChecklist = True
            Else
                EnviarCorreoChecklist = False
            End If
        End Try

        Return EnviarCorreoChecklist

    End Function

    Function Retorna_Ruta_Checklist(ByVal codigo_checklist) As String

        Dim x = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Checklist " + codigo_checklist + ".Pdf"
        If Not File.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Checklist " + codigo_checklist + ".Pdf") Then

            Dim report As New Rpt_CheckList

            Dim PictureBox1 As New PictureBox
            Dim PictureBox2 As New PictureBox
            Dim PictureBox3 As New PictureBox
            Dim PictureBox4 As New PictureBox
            Dim PictureBox5 As New PictureBox
            Dim PictureBox6 As New PictureBox

            Dim SqlImagen As String = "SELECT Convert(varchar(max),Convert(Varbinary(MAX),cl_imgtem)), " +
                                      "       Convert(varchar(max),Convert(Varbinary(MAX),cl_imgsel)), " +
                                      "       Convert(varchar(max),Convert(Varbinary(MAX),cl_imgpat)), " +
                                      "       Convert(varchar(max),Convert(Varbinary(MAX),cl_imgtemS)), " +
                                      "       Convert(varchar(max),Convert(Varbinary(MAX),cl_imgselS)), " +
                                      "       Convert(varchar(max),Convert(Varbinary(MAX),cl_imgpatS)) " +
                                      "  FROM chk_images " +
                                      " WHERE id_chk = '" + Convert.ToInt32(codigo_checklist).ToString() + "'"

            Dim tablaimagen As DataTable = ListarTablasSQL(SqlImagen)

            If tablaimagen.Rows.Count > 0 Then

                'imagen1
                If tablaimagen.Rows(0)(0).ToString() <> "" Then
                    PictureBox1.Image = Base64ToImage(tablaimagen.Rows(0)(0).ToString())
                    PictureBox1.Size = New Size(100, 100)
                Else
                    PictureBox1.Image = My.Resources.blanco
                End If

                'imagen2
                If tablaimagen.Rows(0)(1).ToString() <> "" Then

                    PictureBox2.Image = Base64ToImage(tablaimagen.Rows(0)(1).ToString())
                    PictureBox2.Size = New Size(100, 100)
                Else
                    PictureBox2.Image = My.Resources.blanco
                End If

                'imagen3
                If tablaimagen.Rows(0)(2).ToString() <> "" Then

                    PictureBox3.Image = Base64ToImage(tablaimagen.Rows(0)(2).ToString())
                    PictureBox3.Size = New Size(100, 100)
                Else
                    PictureBox3.Image = My.Resources.blanco
                End If

                'imagen4
                If tablaimagen.Rows(0)(3).ToString() <> "" Then

                    PictureBox4.Image = Base64ToImage(tablaimagen.Rows(0)(3).ToString())
                    PictureBox4.Size = New Size(100, 100)
                Else
                    PictureBox4.Image = My.Resources.blanco
                End If

                'imagen5
                If tablaimagen.Rows(0)(4).ToString() <> "" Then

                    PictureBox5.Image = Base64ToImage(tablaimagen.Rows(0)(4).ToString())
                    PictureBox5.Size = New Size(100, 100)
                Else
                    PictureBox5.Image = My.Resources.blanco
                End If

                'imagen6
                If tablaimagen.Rows(0)(5).ToString() <> "" Then

                    PictureBox6.Image = Base64ToImage(tablaimagen.Rows(0)(5).ToString())
                    PictureBox6.Size = New Size(100, 100)
                Else
                    PictureBox6.Image = My.Resources.blanco
                End If

            Else
                PictureBox1.Image = My.Resources.blanco
                PictureBox2.Image = My.Resources.blanco
                PictureBox3.Image = My.Resources.blanco
                PictureBox4.Image = My.Resources.blanco
                PictureBox5.Image = My.Resources.blanco
                PictureBox6.Image = My.Resources.blanco
            End If

            Dim Ds As New Ds_Imagenes

            Ds.Imagenes.AddImagenesRow(ImageToByte(PictureBox1.Image), ImageToByte(PictureBox2.Image), ImageToByte(PictureBox3.Image), ImageToByte(PictureBox4.Image), ImageToByte(PictureBox5.Image), ImageToByte(PictureBox6.Image))
            Try
                report.Subreports(0).SetDataSource(Ds)
                report.SetParameterValue("codigo", codigo_checklist)
                report.SetDatabaseLogon(setLogon, setPass)
            Catch ex As CrystalReportsException
                Return ""
            End Try

            Try
                Dim CrExportOptions As ExportOptions
                Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
                Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()

                CrDiskFileDestinationOptions.DiskFileName = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Checklist " + codigo_checklist + ".Pdf"
                CrExportOptions = report.ExportOptions
                With CrExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.PortableDocFormat
                    .DestinationOptions = CrDiskFileDestinationOptions
                    .FormatOptions = CrFormatTypeOptions
                End With
                report.Export()
                report.Dispose()
                Return My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Checklist " + codigo_checklist + ".Pdf"
                Exit Function
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End If

        Return ""
    End Function

    Private Function BodyChecklist(ByVal Chofer As String, ByVal patente As String, ByVal Rampla As String, ByVal nombre_cliente As String) As AlternateView

        Dim archivo As String = ArchivoAString(Application.StartupPath + "\Despacho1.txt")

        Dim Chofer_formato As String = Chofer.Chars(0) + Chofer.Chars(1) + "." + Chofer.Chars(2) + Chofer.Chars(3) + Chofer.Chars(4) + "." + _
            Chofer.Chars(5) + Chofer.Chars(6) + Chofer.Chars(7) + "-" + Chofer.Chars(8)

        Dim editable = "<h1>Estimado Cliente " + nombre_cliente.ToString() + ", " +
                        "<br><br>OrderBy informa que el camion  <spa>se ha retirado</spa> </h1><br><br></div>" +
                        "<div id='solicitud'><div>" +
                        "<table border='0'>" +
                        "<tr>" +
                        "<td>Chofer</td>" +
                        "<td>" + Chofer.Remove(0, 9).ToUpper() + "  " + Chofer_formato + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Patente</td>" +
                        "<td>" + patente.ToUpper() + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Rampla</td>" +
                        "<td>" + Rampla.ToUpper() + "</td>" +
                        "</tr>" +
                        "</table><br>"

        Dim archivo2 As String = ArchivoAString(Application.StartupPath + "\Despacho2.txt")

        archivo = archivo + editable + archivo2

        Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim htmlView2 As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim cabecera As LinkedResource = New LinkedResource(Application.StartupPath + "\tituloRecepcion.jpg", MediaTypeNames.Image.Jpeg)
        cabecera.ContentId = "titulo"

        htmlView.LinkedResources.Add(cabecera)

        Return htmlView

    End Function

    Function Base64ToImage(ByVal base64string As String) As System.Drawing.Image

        Dim img As System.Drawing.Image
        Dim MS As System.IO.MemoryStream = New System.IO.MemoryStream
        Dim b64 As String
        Dim b() As Byte

        Try
            b64 = base64string.Replace(" ", "+").Trim
        Catch ex As OutOfMemoryException
            img = My.Resources.blanco
        End Try

        Try
            b = Convert.FromBase64String(b64)
            MS = New System.IO.MemoryStream(b)

            img = System.Drawing.Image.FromStream(MS)

        Catch ex As Exception
            img = My.Resources.blanco
        End Try

        Return img


    End Function

    Public Function ImageToByte(ByVal bmp As Image) As Byte()
        Dim converter As New ImageConverter()
        Try
            Try
                Dim Bitmap = New Bitmap(bmp.Width, bmp.Height, bmp.PixelFormat)
                Dim g = Graphics.FromImage(Bitmap)

                g.DrawImage(bmp, New Point(0, 0))
                g.Dispose()
                bmp.Dispose()

                bmp = Bitmap
                Return DirectCast(converter.ConvertTo(bmp, GetType(Byte())), Byte())

            Catch ex As OutOfMemoryException
                bmp = My.Resources.blanco
                Dim Bitmap = New Bitmap(bmp.Width, bmp.Height, bmp.PixelFormat)
                Dim g = Graphics.FromImage(Bitmap)

                g.DrawImage(bmp, New Point(0, 0))
                g.Dispose()
                bmp.Dispose()

                bmp = Bitmap
                Return DirectCast(converter.ConvertTo(bmp, GetType(Byte())), Byte())
            End Try

        Catch ex As ArgumentException
            bmp = My.Resources.blanco
            Dim Bitmap = New Bitmap(bmp.Width, bmp.Height, bmp.PixelFormat)
            Dim g = Graphics.FromImage(Bitmap)

            g.DrawImage(bmp, New Point(0, 0))
            g.Dispose()
            bmp.Dispose()

            bmp = Bitmap
            Return DirectCast(converter.ConvertTo(bmp, GetType(Byte())), Byte())
        End Try

    End Function

    '-------------------------------
    'CORREO DE DESPACHO ------------
    '-------------------------------

    Public Function EnviarCorreoDespacho(ByVal Codigo_Despacho As String, ByVal rut_cliente As String, ByVal inf_cod As String) As Boolean

        Dim smtp As New System.Net.Mail.SmtpClient
        Dim correo As New System.Net.Mail.MailMessage
        Dim adjunto As System.Net.Mail.Attachment

        With smtp
            .Port = puerto
            .Host = host_mail
            .Credentials = New System.Net.NetworkCredential(correoenvio, claveenvio)
            .EnableSsl = estadoSSL
        End With

        With correo

            Dim sql As String = "SELECT cli_nomb, cli_cryd FROM clientes WHERE cli_rut='" + rut_cliente + "'"

            Dim tablaCliente As DataTable = ListarTablasSQL(sql)
            Dim NombreCliente As String = ""
            If tablaCliente.Rows.Count > 0 Then
                NombreCliente = tablaCliente.Rows(0)(0).ToString()
            End If

            .From = New System.Net.Mail.MailAddress(correomostrar)

            Dim uu As String = tablaCliente.Rows(0)(1).ToString().Trim

            If QuitarCaracteres(uu.ToString()).Length < uu.ToString().Length Then
                Dim correo_electronico As String = ""
                For i As Integer = 0 To uu.Length - 1
                    If uu.Chars(i) <> ";" Then
                        correo_electronico = correo_electronico + uu.Chars(i)
                    Else
                        .To.Add(correo_electronico)
                        correo_electronico = ""
                    End If

                Next
                .To.Add(correo_electronico)
            End If

            'Llamada a funcion que carga copia a correo

            Dim sqlInterno As String = "SELECT inf_copia FROM informes WHERE inf_cod ='" + inf_cod + "'"

            Dim tablaInformes As DataTable = ListarTablasSQL(sqlInterno)
            Dim NombreInformes As String = ""
            If tablaInformes.Rows.Count > 0 Then

                Dim uuu As String = tablaInformes.Rows(0)(0).ToString().Trim

                If uuu <> "" Then
                    If QuitarCaracteres(uuu.ToString()).Length < uuu.ToString().Length Then
                        Dim correo_electronico As String = ""
                        For i As Integer = 0 To uuu.Length - 1
                            If uuu.Chars(i) <> ";" Then
                                correo_electronico = correo_electronico + uuu.Chars(i)
                            Else
                                .Bcc.Add(correo_electronico)
                                correo_electronico = ""
                            End If

                        Next
                        .Bcc.Add(correo_electronico)
                    End If
                End If
            End If

            .Subject = "Despacho (" + Codigo_Despacho + ") finalizado, transporte liberado"
            .IsBodyHtml = True

            Dim dt As DataTable = ListarTablasSQL("SELECT fdes_rutcond, cho_nombre, cho_patente, cho_patente2 " +
                                                  "  FROM fichdespa, choferes " +
                                                  " WHERE fdes_rutcond = cho_rut " +
                                                  "   AND fdes_codi = '" + Codigo_Despacho.ToString() + "'")
            If dt.Rows.Count = 0 Then
                Return True
                Exit Function
            End If

            .AlternateViews.Add(BodyDespacho(dt.Rows(0)(0).ToString() + "    " + dt.Rows(0)(1).ToString(), dt.Rows(0)(2).ToString(), dt.Rows(0)(3).ToString(), NombreCliente))
            .Priority = System.Net.Mail.MailPriority.Normal
            Dim ruta As String = Retorna_Ruta_ArchivoDespacho(Codigo_Despacho)
            If ruta = "" Then
                Return True
                Exit Function
            End If
            adjunto = New System.Net.Mail.Attachment(ruta)
            .Attachments.Add(adjunto)
        End With

        Threading.Thread.Sleep(10000)

        Try

            'Dim Se_envio As String = " SELECT * FROM DocumentosEnviados WHERE Denv_seccion = '009' AND Denv_Dcto='" + Codigo_Despacho + "' and convert(date,Denv_fecha)=convert(date,GETDATE())"
            Dim Se_envio As String = "select a.* from DocumentosEnviados a with(nolock) inner join fichdespa b on(a.Denv_Dcto=b.fdes_codi and CONVERT(date,a.Denv_fecha)=CONVERT(date,b.fdes_emis) and a.Denv_hora>left(CONVERT(varchar,b.fdes_emis,108),5)) where a.Denv_seccion='009' and a.Denv_Dcto='" + Codigo_Despacho + "'"
            Dim tabla As DataTable = ListarTablasSQL(Se_envio)
            If tabla.Rows.Count > 0 Then
                Return True
                Exit Function
            End If

            smtp.Send(correo)

            Dim sql As String = " INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora, Denv_Dcto)" +
                                " VALUES ('009','" + devuelve_fecha(buscaHoraServidor()) + "',left(CONVERT(varchar,getdate(),108),5),'" + Codigo_Despacho + "')"
            MovimientoSQL(sql)

            Dim sqlActualiza As String = "UPDATE fichdespa SET fdes_enviada='1' WHERE fdes_codi='" + Codigo_Despacho.ToString() + "'"
            MovimientoSQL(sqlActualiza)

            EnviarCorreoDespacho = True
        Catch ex As Exception
            If (ex.Message.Trim = "No se puede enviar a un destinatario.") Then
                Dim sql As String = " INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora, Denv_Dcto)" +
                                " VALUES ('009','" + devuelve_fecha(buscaHoraServidor()) + "',left(CONVERT(varchar,getdate(),108),5),'" + Codigo_Despacho + "')"
                MovimientoSQL(sql)

                Dim sqlActualiza As String = "UPDATE fichdespa SET fdes_enviada='1' WHERE fdes_codi='" + Codigo_Despacho.ToString() + "'"
                MovimientoSQL(sqlActualiza)

                EnviarCorreoDespacho = True
            Else
                EnviarCorreoDespacho = False
            End If
        End Try

        Return EnviarCorreoDespacho

    End Function

    Function Retorna_Ruta_ArchivoDespacho(ByVal codigo As String) As String

        Dim valor As String = ""

        Dim Se_envio1 As String = " SELECT isnull(fdes_enviada,0) as fdes_enviada FROM fichdespa WHERE fdes_codi='" + codigo.ToString + "' "
        Dim tabla1 As DataTable = ListarTablasSQL(Se_envio1)
        If tabla1.Rows.Count > 0 Then
            valor = tabla1.Rows(0)(0).ToString().Trim()
        End If

        If File.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Despacho_" + codigo.ToString() + ".Pdf") Or valor = "1" Then
            File.Delete(My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Despacho_" + codigo.ToString() + ".Pdf")
        End If

        If Not File.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Despacho_" + codigo.ToString() + ".Pdf") Or valor = "1" Then
            Dim report As New Rpt_GuiaDespacho


            Dim PictureBox1 As New PictureBox
            Dim PictureBox2 As New PictureBox
            Dim PictureBox3 As New PictureBox
            Dim PictureBox4 As New PictureBox

            Dim SqlImagen As String = " SELECT MAX(l.dimg_despcodi) id_pallets, " +
                                      "        (SELECT CONVERT(VARCHAR(MAX), CONVERT(VARBINARY(MAX),dimg_imagen2)) FROM despaimagen WHERE dimg_despcodi = '" + codigo + "' AND dimg_num = 1) pic1, " +
                                      "        (SELECT CONVERT(VARCHAR(MAX), CONVERT(VARBINARY(MAX),dimg_imagen2)) FROM despaimagen WHERE dimg_despcodi = '" + codigo + "' AND dimg_num = 2) pic2, " +
                                      "        (SELECT CONVERT(VARCHAR(MAX), CONVERT(VARBINARY(MAX),dimg_imagen2)) FROM despaimagen WHERE dimg_despcodi = '" + codigo + "' AND dimg_num = 3) pic3, " +
                                      "        (SELECT CONVERT(VARCHAR(MAX), CONVERT(VARBINARY(MAX),dimg_imagen2)) FROM despaimagen WHERE dimg_despcodi = '" + codigo + "' AND dimg_num = 4) pic4 " +
                                      "   FROM despaimagen l " +
                                      "  WHERE l.dimg_despcodi = '" + codigo + "'"

            Dim tablaimagen As DataTable = ListarTablasSQL(SqlImagen)

            If tablaimagen.Rows.Count > 0 Then

                'imagen1
                If tablaimagen.Rows(0)(1).ToString() <> "" Then
                    PictureBox1.Image = Base64ToImage(tablaimagen.Rows(0)(1).ToString())
                    PictureBox1.Size = New Size(100, 100)
                    PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
                Else
                    PictureBox1.Image = My.Resources.blanco
                End If

                'imagen2
                If tablaimagen.Rows(0)(2).ToString() <> "" Then

                    PictureBox2.Image = Base64ToImage(tablaimagen.Rows(0)(2).ToString())
                    PictureBox2.Size = New Size(100, 100)
                    PictureBox2.SizeMode = PictureBoxSizeMode.StretchImage
                Else
                    PictureBox2.Image = My.Resources.blanco
                End If

                'imagen3
                If tablaimagen.Rows(0)(3).ToString() <> "" Then

                    PictureBox3.Image = Base64ToImage(tablaimagen.Rows(0)(3).ToString())
                    PictureBox3.Size = New Size(100, 100)
                    PictureBox3.SizeMode = PictureBoxSizeMode.StretchImage
                Else
                    PictureBox3.Image = My.Resources.blanco
                End If

                'imagen4
                If tablaimagen.Rows(0)(4).ToString() <> "" Then

                    PictureBox4.Image = Base64ToImage(tablaimagen.Rows(0)(4).ToString())
                    PictureBox4.Size = New Size(100, 100)
                    PictureBox4.SizeMode = PictureBoxSizeMode.StretchImage

                Else
                    PictureBox4.Image = My.Resources.blanco
                End If



            Else
                PictureBox1.Image = My.Resources.blanco
                PictureBox2.Image = My.Resources.blanco
                PictureBox3.Image = My.Resources.blanco
                PictureBox4.Image = My.Resources.blanco

            End If

            Dim Ds As New fotos

            Ds.fotografias.AddfotografiasRow(ImageToByte(PictureBox1.Image), ImageToByte(PictureBox2.Image), ImageToByte(PictureBox3.Image), ImageToByte(PictureBox4.Image))

            Try
                report.Subreports(1).SetDataSource(Ds)
                report.SetDatabaseLogon(setLogon, setPass)
                report.SetParameterValue("codigo", codigo)
            Catch ex As CrystalReportsException
                Return ""
                Exit Function
            End Try
            Try
                Dim CrExportOptions As ExportOptions
                Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
                Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()

                CrDiskFileDestinationOptions.DiskFileName = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Despacho_" + codigo.ToString() + ".Pdf"
                CrExportOptions = report.ExportOptions
                With CrExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.PortableDocFormat
                    .DestinationOptions = CrDiskFileDestinationOptions
                    .FormatOptions = CrFormatTypeOptions
                End With
                report.Export()
                report.Dispose()
                Return My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Despacho_" + codigo.ToString() + ".Pdf"
                Exit Function
            Catch ex As Exception
                Return ""
                Exit Function
            End Try
        Else
            Return My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Despacho_" + codigo.ToString() + ".Pdf"
            Exit Function
        End If

        Return ""
    End Function

    Private Function BodyDespacho(ByVal Chofer As String, ByVal patente As String, ByVal Rampla As String, ByVal nombre_cliente As String) As AlternateView

        Dim archivo As String = ArchivoAString(Application.StartupPath + "\Despacho1.txt")

        Dim Chofer_formato As String = Chofer.Chars(0) + Chofer.Chars(1) + "." + Chofer.Chars(2) + Chofer.Chars(3) + Chofer.Chars(4) + "." + _
            Chofer.Chars(5) + Chofer.Chars(6) + Chofer.Chars(7) + "-" + Chofer.Chars(8)

        Dim editable = "<h1>Estimado Cliente " + nombre_cliente.ToString() + ", " +
                        "<br><br>OrderBy informa que el camion enviado a <spa>CARGAR</spa>  ya ha sido despachado y  <span> Liberado!!!</span> de anden</h1><br><br></div>" +
                        "<div id='solicitud'><div>" +
                        "<table border='0'>" +
                        "<tr>" +
                        "<td>Chofer</td>" +
                        "<td>" + Chofer.Remove(0, 9).ToUpper() + "  " + Chofer_formato + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Patente</td>" +
                        "<td>" + patente.ToUpper() + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + "</td>" +
                        "</tr>" +
                        "<tr>" +
                        "<td>Rampla</td>" +
                        "<td>" + Rampla.ToUpper() + "</td>" +
                        "</tr>" +
                        "</table><br>"

        Dim archivo2 As String = ArchivoAString(Application.StartupPath + "\Despacho2.txt")

        archivo = archivo + editable + archivo2

        Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)

        Dim htmlView2 As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim cabecera As LinkedResource = New LinkedResource(Application.StartupPath + "\tituloRecepcion.jpg", MediaTypeNames.Image.Jpeg)
        cabecera.ContentId = "titulo"

        htmlView.LinkedResources.Add(cabecera)

        Return htmlView

    End Function

    '-------------------------------
    'CORREO PEDIDOS DIARIO ---------
    '-------------------------------

    Public Function EnviarCorreoPedidosHora(ByVal inf_cod As String) As Boolean

        Dim smtp As New System.Net.Mail.SmtpClient
        Dim correo As New System.Net.Mail.MailMessage
        Dim adjunto As System.Net.Mail.Attachment

        With smtp
            .Port = puerto
            .Host = host_mail
            .Credentials = New System.Net.NetworkCredential(correoenvio, claveenvio)
            .EnableSsl = estadoSSL
        End With

        With correo

            Dim sql As String = "SELECT prg_mail FROM informes_programa WHERE prg_inf_cod='" + inf_cod + "' AND prg_emp='0'"

            Dim tablaInterno As DataTable = ListarTablasSQL(sql)

            .From = New System.Net.Mail.MailAddress(correomostrar)

            Dim uu As String = tablaInterno.Rows(0)(0).ToString().Trim

            If QuitarCaracteres(uu.ToString()).Length < uu.ToString().Length Then
                Dim correo_electronico As String = ""
                For i As Integer = 0 To uu.Length - 1
                    If uu.Chars(i) <> ";" Then
                        correo_electronico = correo_electronico + uu.Chars(i)
                    Else
                        .To.Add(correo_electronico)
                        correo_electronico = ""
                    End If
                Next
                .To.Add(correo_electronico)
            End If

            .Subject = "Tiempos de pedidos solicitados a la fecha " + devuelve_fecha2(buscaHoraServidor())
            .IsBodyHtml = True
            .AlternateViews.Add(BodyPedidosHora())
            .Priority = System.Net.Mail.MailPriority.Normal

            Dim ruta As String = Retorna_Ruta_ArchivoPedidosHora()
            If ruta = "" Then
                Return True
                Exit Function
            End If
            adjunto = New System.Net.Mail.Attachment(ruta)
            .Attachments.Add(adjunto)
        End With

        Try
            smtp.Send(correo)

            Dim sql As String = "INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora)" +
                                "VALUES ('PDIARIO','" + devuelve_fecha(buscaHoraServidor()) + "','" + DevuelveHora() + "')"
            MovimientoSQL(sql)

            EnviarCorreoPedidosHora = True
        Catch ex As Exception
            If (ex.Message.Trim = "No se puede enviar a un destinatario.") Then
                Dim sql As String = "INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora)" +
                                "VALUES ('PDIARIO','" + devuelve_fecha(buscaHoraServidor()) + "','" + DevuelveHora() + "')"
                MovimientoSQL(sql)

                EnviarCorreoPedidosHora = True
            Else
                EnviarCorreoPedidosHora = False
            End If
        End Try

        Return EnviarCorreoPedidosHora

    End Function

    Function Retorna_Ruta_ArchivoPedidosHora() As String

        Dim Fecha = devuelve_fecha(buscaHoraServidor()).Replace("/", "_")
        Dim x = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Listado " + Fecha + ".Pdf"
        If Not File.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Listado " + Fecha + ".Pdf") Then
            Dim report As New Rpt_Pedidos
            Try
                report.SetDatabaseLogon(setLogon, setPass)
            Catch ex As CrystalReportsException
                Return ""
                Exit Function
            End Try

            Try
                Dim CrExportOptions As ExportOptions
                Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
                Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()

                CrDiskFileDestinationOptions.DiskFileName = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Listado " + Fecha + ".Pdf"
                CrExportOptions = report.ExportOptions
                With CrExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.PortableDocFormat
                    .DestinationOptions = CrDiskFileDestinationOptions
                    .FormatOptions = CrFormatTypeOptions
                End With
                report.Export()
                report.Dispose()
                Return My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Listado " + Fecha + ".Pdf"
                Exit Function
            Catch ex As Exception
                Return ""
                Exit Function
            End Try
        End If
        Return ""

    End Function

    Private Function BodyPedidosHora() As AlternateView

        Dim archivo As String = ArchivoAString(Application.StartupPath + "\PedidosHora.txt")

        Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim htmlView2 As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim cabecera As LinkedResource = New LinkedResource(Application.StartupPath + "\tituloph.jpg", MediaTypeNames.Image.Jpeg)
        cabecera.ContentId = "titulo"

        htmlView.LinkedResources.Add(cabecera)


        Return htmlView

    End Function

    Function Retorna_Ruta_Archivo_StockComercialAgrosuper() As String
        Try
            Dim NomArch As String = "stockdiarioprecisa" & DateAdd(DateInterval.Day, 1, Now).ToString("ddMMyy") & ".CSV"
            Dim x = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\" & NomArch

            If Not File.Exists(x) Then
                Dim sql As String = "select * from V_Stock_Comercial_Agrosuper with(nolock)"
                Dim dt As New DataTable
                dt = ListarTablasSQL(sql)

                If (dt.Rows.Count > 0) Then
                    Dim file As StreamWriter
                    file = My.Computer.FileSystem.OpenTextFileWriter(x, True)

                    Dim LinTit As String = "FECHA;CODIGO_CLIENTE;PRODUCTO;LOTE;EMBARQUE;ENVASES;NETO;PLANTA;FOLIO_INTERNO;FECHA_RECEPCION;FECHA_PRODUCCION;STATUS_FRIGORIFICO"

                    file.WriteLine(LinTit)

                    For i = 0 To dt.Rows.Count - 1
                        Dim LinDat As String = ""

                        For j = 0 To dt.Rows(i).ItemArray.Count - 1
                            If (LinDat = "") Then
                                LinDat = dt.Rows(i).Item(j).ToString.Trim
                            Else
                                LinDat &= ";" & dt.Rows(i).Item(j).ToString.Trim
                            End If
                        Next

                        'LinDat &= ";"

                        file.WriteLine(LinDat)
                    Next

                    file.Close()
                Else
                    Return ""
                End If

                Return x
            Else
                Return ""
            End If
        Catch ex As Exception
            Return ""
            Exit Function
        End Try
    End Function

    '-------------------------------
    'CORREO DIARIO STOCK COMERCIAL AGROSUPER
    '-------------------------------

    Public Function EnviarCorreoStockComercialAgrosuper() As Boolean

        Dim smtp As New System.Net.Mail.SmtpClient
        Dim correo As New System.Net.Mail.MailMessage
        Dim adjunto As System.Net.Mail.Attachment

        With smtp
            .Port = puerto
            .Host = host_mail
            .Credentials = New System.Net.NetworkCredential(correoenvio, claveenvio)
            .EnableSsl = estadoSSL
        End With

        With correo

            Dim sql As String = "SELECT prg_mail FROM informes_programa WHERE prg_inf_cod='22' AND prg_emp='0'"

            Dim tablaInterno As DataTable = ListarTablasSQL(sql)

            .From = New System.Net.Mail.MailAddress(correomostrar)

            Dim uu As String = tablaInterno.Rows(0)(0).ToString().Trim

            If QuitarCaracteres(uu.ToString()).Length < uu.ToString().Length Then
                Dim correo_electronico As String = ""
                For i As Integer = 0 To uu.Length - 1
                    If uu.Chars(i) <> ";" Then
                        correo_electronico = correo_electronico + uu.Chars(i)
                    Else
                        .To.Add(correo_electronico)
                        correo_electronico = ""
                    End If
                Next
                .To.Add(correo_electronico)
            End If

            .Subject = "Informe stock diario Comercial Agrosuper " + devuelve_fecha2(buscaHoraServidor())
            .IsBodyHtml = True
            .AlternateViews.Add(BodyStockComercialAgrosuper())
            .Priority = System.Net.Mail.MailPriority.Normal

            Dim ruta As String = Retorna_Ruta_Archivo_StockComercialAgrosuper()
            If ruta = "" Then
                Return True
                Exit Function
            End If
            adjunto = New System.Net.Mail.Attachment(ruta)
            .Attachments.Add(adjunto)
        End With

        Try
            smtp.Send(correo)

            Dim sql As String = "INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora)" +
                                "VALUES ('STOCK COMERCIAL AGROSUPER','" + devuelve_fecha(buscaHoraServidor()) + "','" + DevuelveHora() + "')"
            MovimientoSQL(sql)

            EnviarCorreoStockComercialAgrosuper = True
        Catch ex As Exception
            Dim sql As String = ""

            If (ex.Message.Trim = "No se puede enviar a un destinatario.") Then
                sql = "INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora)" +
                                "VALUES ('STOCK COMERCIAL AGROSUPER','" + devuelve_fecha(buscaHoraServidor()) + "','" + DevuelveHora() + "')"
                MovimientoSQL(sql)

                EnviarCorreoStockComercialAgrosuper = True
            Else
                sql = "INSERT INTO DocumentosNoEnviados(CODIGO,FECHA,ERROR) VALUES ('22','" + devuelve_fecha(buscaHoraServidor()) + "','Error al enviar informe de stock diario Comercial Agrosuper')"
                MovimientoSQL(sql)

                EnviarCorreoStockComercialAgrosuper = False
            End If
        End Try

        Return EnviarCorreoStockComercialAgrosuper()

    End Function

    '-------------------------------
    'CORREO PALLET VENCIDOS --------
    '-------------------------------

    Public Function EnviarCorreoVencidos(ByVal cliente As String, ByVal inf_cod As String) As Boolean

        Dim smtp As New System.Net.Mail.SmtpClient
        Dim correo As New System.Net.Mail.MailMessage
        Dim adjunto As System.Net.Mail.Attachment

        With smtp
            .Port = puerto
            .Host = host_mail
            .Credentials = New System.Net.NetworkCredential(correoenvio, claveenvio)
            .EnableSsl = estadoSSL
        End With

        With correo
            Dim sql As String = "SELECT cli_nomb, cli_pvenc FROM clientes WHERE cli_rut='" + cliente.ToString() + "'"
            Dim tablaCliente As DataTable = ListarTablasSQL(sql)
            Dim NombreCliente As String = ""
            If tablaCliente.Rows.Count > 0 Then
                NombreCliente = tablaCliente.Rows(0)(0).ToString()
            End If

            .From = New System.Net.Mail.MailAddress(correomostrar)

            Dim uu As String = tablaCliente.Rows(0)(1).ToString().Trim

            If QuitarCaracteres(uu.ToString()).Length < uu.ToString().Length Then
                Dim correo_electronico As String = ""
                For i As Integer = 0 To uu.Length - 1
                    If uu.Chars(i) <> ";" Then
                        correo_electronico = correo_electronico + uu.Chars(i)
                    Else
                        .To.Add(correo_electronico)
                        correo_electronico = ""
                    End If
                Next
                .To.Add(correo_electronico)
            End If

            'Llamada a funcion que carga copia a correo

            Dim sqlInterno As String = "SELECT inf_copia FROM informes WHERE inf_cod ='" + inf_cod + "'"

            Dim tablaInformes As DataTable = ListarTablasSQL(sqlInterno)
            Dim NombreInformes As String = ""
            If tablaInformes.Rows.Count > 0 Then

                Dim uuu As String = tablaInformes.Rows(0)(0).ToString().Trim

                If uuu <> "" Then
                    If QuitarCaracteres(uuu.ToString()).Length < uuu.ToString().Length Then
                        Dim correo_electronico As String = ""
                        For i As Integer = 0 To uuu.Length - 1
                            If uuu.Chars(i) <> ";" Then
                                correo_electronico = correo_electronico + uuu.Chars(i)
                            Else
                                .Bcc.Add(correo_electronico)
                                correo_electronico = ""
                            End If

                        Next
                        .Bcc.Add(correo_electronico)
                    End If
                End If
            End If

            .Subject = "Informe de soportantes vencidos al dia " + devuelve_fecha2(buscaHoraServidor())
            .IsBodyHtml = True
            .AlternateViews.Add(BodyVencidos())
            .Priority = System.Net.Mail.MailPriority.Normal
            Dim ruta As String = Retorna_Ruta_ArchivoVencidos(cliente)
            If ruta = "" Then
                Return True
                Exit Function
            End If
            adjunto = New System.Net.Mail.Attachment(ruta)
            .Attachments.Add(adjunto)
        End With

        Try
            smtp.Send(correo)

            Dim sql As String = "INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora, denv_dcto)" +
                                "VALUES ('PVENC','" + devuelve_fecha(buscaHoraServidor()) + "','" + DevuelveHora() + "','" + cliente + "')"
            MovimientoSQL(sql)

            EnviarCorreoVencidos = True
        Catch ex As Exception
            If (ex.Message.Trim = "No se puede enviar a un destinatario.") Then
                Dim sql As String = "INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora, denv_dcto)" +
                                "VALUES ('PVENC','" + devuelve_fecha(buscaHoraServidor()) + "','" + DevuelveHora() + "','" + cliente + "')"
                MovimientoSQL(sql)

                EnviarCorreoVencidos = True
            Else
                EnviarCorreoVencidos = False
            End If
        End Try

        Return EnviarCorreoVencidos

    End Function

    Function Retorna_Ruta_ArchivoVencidos(ByVal cliente As String) As String

        Dim Fecha = devuelve_fecha(buscaHoraServidor()).Replace("/", "_")

        If Not File.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Vencidos " + cliente + "_" + Fecha + ".Pdf") Then

            Dim report As New Rpt_PalletVencidos

            Try
                report.SetParameterValue("codigo", cliente)
                report.SetDatabaseLogon(setLogon, setPass)
            Catch ex As CrystalReportsException
                Return ""
            End Try

            Try
                Dim CrExportOptions As ExportOptions
                Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
                Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()

                CrDiskFileDestinationOptions.DiskFileName = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Vencidos " + cliente + "_" + Fecha + ".Pdf"
                CrExportOptions = report.ExportOptions
                With CrExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.PortableDocFormat
                    .DestinationOptions = CrDiskFileDestinationOptions
                    .FormatOptions = CrFormatTypeOptions
                End With
                report.Export()
                report.Dispose()
                Return My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Vencidos " + cliente + "_" + Fecha + ".Pdf"
                Exit Function
            Catch ex As Exception
                Return ""
            End Try

        End If

        Return ""
    End Function

    Private Function BodyVencidos() As AlternateView

        Dim archivo As String = ArchivoAString(Application.StartupPath + "\Vencidos.txt")

        Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim htmlView2 As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim cabecera As LinkedResource = New LinkedResource(Application.StartupPath + "\tituloV.jpg", MediaTypeNames.Image.Jpeg)
        cabecera.ContentId = "titulo"

        htmlView.LinkedResources.Add(cabecera)


        Return htmlView

    End Function

    '-------------------------------
    'CORREO DON RAUL SEMANAL -------
    '-------------------------------

    Public Function EnviarCorreoPedidosSemanal(ByVal inf_cod As String) As Boolean

        Dim smtp As New System.Net.Mail.SmtpClient
        Dim correo As New System.Net.Mail.MailMessage
        Dim adjunto As System.Net.Mail.Attachment

        With smtp
            .Port = puerto
            .Host = host_mail
            .Credentials = New System.Net.NetworkCredential(correoenvio, claveenvio)
            .EnableSsl = estadoSSL
        End With

        With correo

            Dim sql As String = "SELECT prg_mail FROM informes_programa WHERE prg_inf_cod='" + inf_cod + "' AND prg_emp='0'"

            Dim tablaInterno As DataTable = ListarTablasSQL(sql)


            .From = New System.Net.Mail.MailAddress(correomostrar)

            Dim uu As String = tablaInterno.Rows(0)(0).ToString().Trim

            If QuitarCaracteres(uu.ToString()).Length < uu.ToString().Length Then
                Dim correo_electronico As String = ""
                For i As Integer = 0 To uu.Length - 1
                    If uu.Chars(i) <> ";" Then
                        correo_electronico = correo_electronico + uu.Chars(i)
                    Else
                        .To.Add(correo_electronico)
                        correo_electronico = ""
                    End If

                Next
                .To.Add(correo_electronico)
            End If

            .Subject = "Listado de tiempos de pedidos realizados la semana anterior "
            .IsBodyHtml = True
            .AlternateViews.Add(BodyPedidosSemanal())
            .Priority = System.Net.Mail.MailPriority.Normal

            Dim ruta As String = Retorna_Ruta_ArchivoPedidosSemanal()
            If ruta = "" Then
                Return True
                Exit Function
            End If
            adjunto = New System.Net.Mail.Attachment(ruta)
            .Attachments.Add(adjunto)
        End With

        Try
            smtp.Send(correo)

            Dim sql As String = "INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora)" +
                                "VALUES ('PSEMANAL','" + devuelve_fecha(buscaHoraServidor()) + "','" + DevuelveHora() + "')"

            MovimientoSQL(sql)

            EnviarCorreoPedidosSemanal = True
        Catch ex As Exception
            If (ex.Message.Trim = "No se puede enviar a un destinatario.") Then
                Dim sql As String = "INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora)" +
                                "VALUES ('PSEMANAL','" + devuelve_fecha(buscaHoraServidor()) + "','" + DevuelveHora() + "')"

                MovimientoSQL(sql)

                EnviarCorreoPedidosSemanal = True
            Else
                EnviarCorreoPedidosSemanal = False
            End If
        End Try

        Return EnviarCorreoPedidosSemanal

    End Function

    Function Retorna_Ruta_ArchivoPedidosSemanal() As String

        Dim Fecha = devuelve_fecha(buscaHoraServidor()).Replace("/", "_")
        Dim x = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Listado Semanal " + Fecha + ".Pdf"
        If Not File.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Listado Semanal " + Fecha + ".Pdf") Then
            Dim report As New Rpt_PedidosSemanal
            report.SetDatabaseLogon(setLogon, setPass)

            Try
                Dim CrExportOptions As ExportOptions
                Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
                Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()

                CrDiskFileDestinationOptions.DiskFileName = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Listado Semanal " + Fecha + ".Pdf"
                CrExportOptions = report.ExportOptions
                With CrExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.PortableDocFormat
                    .DestinationOptions = CrDiskFileDestinationOptions
                    .FormatOptions = CrFormatTypeOptions
                End With
                report.Export()
                report.Dispose()
                Return My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\Listado Semanal " + Fecha + ".Pdf"
                Exit Function
            Catch ex As Exception

            End Try

        End If
        Return ""

    End Function

    Private Function BodyPedidosSemanal() As AlternateView

        Dim archivo As String = ArchivoAString(Application.StartupPath + "\PedidosSemanal.txt")

        Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim htmlView2 As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim cabecera As LinkedResource = New LinkedResource(Application.StartupPath + "\tituloph.jpg", MediaTypeNames.Image.Jpeg)
        cabecera.ContentId = "titulo"

        htmlView.LinkedResources.Add(cabecera)

        Return htmlView

    End Function

    '-------------------------------
    'CORREO ESPACIO CAMARAS --------
    '-------------------------------

    Public Function EnviarPosicionesCamaras(ByVal inf_cod As String) As Boolean

        Dim smtp As New System.Net.Mail.SmtpClient
        Dim correo As New System.Net.Mail.MailMessage
        Dim adjunto As System.Net.Mail.Attachment

        With smtp
            .Port = puerto
            .Host = host_mail
            .Credentials = New System.Net.NetworkCredential(correoenvio, claveenvio)
            .EnableSsl = estadoSSL
        End With

        With correo

            Dim sql As String = "SELECT prg_mail FROM informes_programa WHERE prg_inf_cod='" + inf_cod + "' AND prg_emp='0'"

            Dim tablaInterno As DataTable = ListarTablasSQL(sql)


            .From = New System.Net.Mail.MailAddress(correomostrar)

            Dim uu As String = tablaInterno.Rows(0)(0).ToString().Trim

            If QuitarCaracteres(uu.ToString()).Length < uu.ToString().Length Then
                Dim correo_electronico As String = ""
                For i As Integer = 0 To uu.Length - 1
                    If uu.Chars(i) <> ";" Then
                        correo_electronico = correo_electronico + uu.Chars(i)
                    Else
                        .To.Add(correo_electronico)
                        correo_electronico = ""
                    End If

                Next
                .To.Add(correo_electronico)
            End If

            .Subject = "Disponibilidad en camaras"

            .IsBodyHtml = True
            .AlternateViews.Add(BodyDisponibilidad())
            .BodyEncoding = System.Text.Encoding.UTF8

            Dim ruta As String = Retorna_Ruta_ArchivoDisponibilidad()
            If ruta = "" Then
                Return True
                Exit Function
            End If
            adjunto = New System.Net.Mail.Attachment(ruta)
            .Attachments.Add(adjunto)
        End With
        Try
            Try
                smtp.Send(correo)

                Dim sql As String = " INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora) " +
                                    " VALUES ('POSICIONES','" + devuelve_fecha(buscaHoraServidor()) + "','" + DevuelveHora() + "')"
                MovimientoSQL(sql)

                EnviarPosicionesCamaras = True
            Catch ex As Exception
                If (ex.Message.Trim = "No se puede enviar a un destinatario.") Then
                    Dim sql As String = " INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora) " +
                                    " VALUES ('POSICIONES','" + devuelve_fecha(buscaHoraServidor()) + "','" + DevuelveHora() + "')"
                    MovimientoSQL(sql)

                    EnviarPosicionesCamaras = True
                Else
                    Dim sqlNo As String = "INSERT INTO DocumentosNoEnviados(CODIGO,FECHA,ERROR) VALUES ('" + inf_cod + "',GETDATE(),'" + ex.ToString + "')"
                    MovimientoSQL(sqlNo)
                    EnviarPosicionesCamaras = False
                End If
            End Try
        Catch ex As SmtpFailedRecipientException
            If (ex.Message.Trim = "No se puede enviar a un destinatario.") Then
                Dim sql As String = " INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora) " +
                                " VALUES ('POSICIONES','" + devuelve_fecha(buscaHoraServidor()) + "','" + DevuelveHora() + "')"
                MovimientoSQL(sql)

                EnviarPosicionesCamaras = True
            Else
                Dim sqlNo As String = "INSERT INTO DocumentosNoEnviados(CODIGO,FECHA,ERROR) VALUES ('" + inf_cod + "',GETDATE(),'" + ex.ToString + "')"
                MovimientoSQL(sqlNo)
                EnviarPosicionesCamaras = False
            End If
        End Try


        Return EnviarPosicionesCamaras
    End Function

    Function Retorna_Ruta_ArchivoDisponibilidad() As String

        Dim Fecha = devuelve_fecha(buscaHoraServidor()).Replace("/", "_")
        Dim x = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\dispcamaras " + Fecha + ".Pdf"
        If Not File.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\dispcamaras " + Fecha + ".Pdf") Then
            Dim report As New Rpt_Posiciones
            Try
                report.SetDatabaseLogon(setLogon, setPass)
            Catch ex As CrystalReportsException
                Return ""
                Exit Function
            End Try

            Try
                Dim CrExportOptions As ExportOptions
                Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
                Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()

                CrDiskFileDestinationOptions.DiskFileName = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\dispcamaras " + Fecha + ".Pdf"
                CrExportOptions = report.ExportOptions
                With CrExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.PortableDocFormat
                    .DestinationOptions = CrDiskFileDestinationOptions
                    .FormatOptions = CrFormatTypeOptions
                End With
                report.Export()
                report.Dispose()
                Return My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\dispcamaras " + Fecha + ".Pdf"
                Exit Function
            Catch ex As Exception
                Return ""
                Exit Function
            End Try
        End If
        Return ""

    End Function

    Private Function BodyDisponibilidad() As AlternateView

        Dim archivo As String = ArchivoAString(Application.StartupPath + "\Disponibles.txt")

        Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim htmlView2 As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim cabecera As LinkedResource = New LinkedResource(Application.StartupPath + "\titulocamaras.jpg", MediaTypeNames.Image.Jpeg)
        cabecera.ContentId = "titulo"

        htmlView.LinkedResources.Add(cabecera)

        Return htmlView

    End Function

    '--------------------------------
    'CORREO MOVIMIENTOS SOPORTANTES -
    '--------------------------------

    Public Function EnviarInformeSoportantes(ByVal inf_cod As String) As Boolean

        Dim smtp As New System.Net.Mail.SmtpClient
        Dim correo As New System.Net.Mail.MailMessage
        Dim adjunto As System.Net.Mail.Attachment


        With smtp
            .Port = puerto
            .Host = host_mail
            .Credentials = New System.Net.NetworkCredential(correoenvio, claveenvio)
            .EnableSsl = estadoSSL
        End With

        With correo

            Dim sql As String = "SELECT prg_mail FROM informes_programa WHERE prg_inf_cod='" + inf_cod + "' AND prg_emp='0'"

            Dim tablaInterno As DataTable = ListarTablasSQL(sql)


            .From = New System.Net.Mail.MailAddress(correomostrar)

            Dim uu As String = tablaInterno.Rows(0)(0).ToString().Trim

            If QuitarCaracteres(uu.ToString()).Length < uu.ToString().Length Then
                Dim correo_electronico As String = ""
                For i As Integer = 0 To uu.Length - 1
                    If uu.Chars(i) <> ";" Then
                        correo_electronico = correo_electronico + uu.Chars(i)
                    Else
                        .To.Add(correo_electronico)
                        correo_electronico = ""
                    End If

                Next
                .To.Add(correo_electronico)
            End If

            .Subject = "Informe de Soportantes por Usuarios"
            .IsBodyHtml = True
            .AlternateViews.Add(BodyMovSoportantes())
            .BodyEncoding = System.Text.Encoding.UTF8

            Dim ruta As String = Retorna_Ruta_ArchivoMovSoportantes()

            If ruta = "" Then
                Return True
                Exit Function
            End If
            adjunto = New System.Net.Mail.Attachment(ruta)
            .Attachments.Add(adjunto)
        End With
        Try
            Try
                smtp.Send(correo)

                Dim sql As String = " INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora) " +
                                    " VALUES ('SOPORTANTES','" + devuelve_fecha(buscaHoraServidor()) + "','" + DevuelveHora() + "')"
                MovimientoSQL(sql)

                EnviarInformeSoportantes = True
            Catch ex As Exception
                If (ex.Message.Trim = "No se puede enviar a un destinatario.") Then
                    Dim sql As String = " INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora) " +
                                    " VALUES ('SOPORTANTES','" + devuelve_fecha(buscaHoraServidor()) + "','" + DevuelveHora() + "')"
                    MovimientoSQL(sql)

                    EnviarInformeSoportantes = True
                Else
                    Dim sqlNo As String = "INSERT INTO DocumentosNoEnviados(CODIGO,FECHA,ERROR) VALUES ('" + inf_cod + "',GETDATE(),'" + ex.ToString + "')"
                    MovimientoSQL(sqlNo)
                    EnviarInformeSoportantes = False
                End If
            End Try
        Catch ex As SmtpFailedRecipientException
            If (ex.Message.Trim = "No se puede enviar a un destinatario.") Then
                Dim sql As String = " INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora) " +
                                " VALUES ('SOPORTANTES','" + devuelve_fecha(buscaHoraServidor()) + "','" + DevuelveHora() + "')"
                MovimientoSQL(sql)

                EnviarInformeSoportantes = True
            Else
                Dim sqlNo As String = "INSERT INTO DocumentosNoEnviados(CODIGO,FECHA,ERROR) VALUES ('" + inf_cod + "',GETDATE(),'" + ex.ToString + "')"
                MovimientoSQL(sqlNo)
                EnviarInformeSoportantes = False
            End If
        End Try

        Return EnviarInformeSoportantes

    End Function

    Function Retorna_Ruta_ArchivoMovSoportantes() As String

        Dim Fecha = devuelve_fecha(buscaHoraServidor()).Replace("/", "_")
        Dim x = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\MovSoportantes " + Fecha + ".Pdf"
        If Not File.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\MovSoportantes " + Fecha + ".Pdf") Then
            Dim report As New Rpt_MovSoportanteUsuario


            Try
                report.SetDatabaseLogon(setLogon, setPass)
                report.SetParameterValue("fecini", devuelve_fecha(DateAdd(DateInterval.Day, -1, buscaHoraServidor())))
                report.SetParameterValue("fecter", devuelve_fecha(buscaHoraServidor()))
                report.SetParameterValue("horaini", "22:00:00")
                report.SetParameterValue("horater", "08:30:00")
            Catch ex As CrystalReportsException
                Return ""
                Exit Function
            End Try

            Try
                Dim CrExportOptions As ExportOptions
                Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
                Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()

                CrDiskFileDestinationOptions.DiskFileName = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\MovSoportantes " + Fecha + ".Pdf"
                CrExportOptions = report.ExportOptions
                With CrExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.PortableDocFormat
                    .DestinationOptions = CrDiskFileDestinationOptions
                    .FormatOptions = CrFormatTypeOptions
                End With
                report.Export()
                report.Dispose()
                Return My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\MovSoportantes " + Fecha + ".Pdf"
                Exit Function
            Catch ex As Exception
                Return ""
                Exit Function
            End Try
        End If
        Return ""

    End Function

    Private Function BodyMovSoportantes() As AlternateView

        Dim archivo As String = ArchivoAString(Application.StartupPath + "\MovSoportantes.txt")

        Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim htmlView2 As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim cabecera As LinkedResource = New LinkedResource(Application.StartupPath + "\tituloMovSoportantes.jpg", MediaTypeNames.Image.Jpeg)
        cabecera.ContentId = "titulo"

        htmlView.LinkedResources.Add(cabecera)

        Return htmlView

    End Function

    '-------------------------------
    'CORREO DOCUMENTOS EMITIDOS ----
    '-------------------------------

    Public Function EnviarEmitidos(ByVal inf_cod As String) As Boolean

        Dim smtp As New System.Net.Mail.SmtpClient
        Dim correo As New System.Net.Mail.MailMessage
        Dim adjunto As System.Net.Mail.Attachment


        With smtp
            .Port = puerto
            .Host = host_mail
            .Credentials = New System.Net.NetworkCredential(correoenvio, claveenvio)
            .EnableSsl = estadoSSL
        End With

        With correo

            Dim sql As String = "SELECT prg_mail FROM informes_programa WHERE prg_inf_cod='" + inf_cod + "' AND prg_emp='0'"

            Dim tablaInterno As DataTable = ListarTablasSQL(sql)


            .From = New System.Net.Mail.MailAddress(correomostrar)

            Dim uu As String = tablaInterno.Rows(0)(0).ToString().Trim

            If QuitarCaracteres(uu.ToString()).Length < uu.ToString().Length Then
                Dim correo_electronico As String = ""
                For i As Integer = 0 To uu.Length - 1
                    If uu.Chars(i) <> ";" Then
                        correo_electronico = correo_electronico + uu.Chars(i)
                    Else
                        .To.Add(correo_electronico)
                        correo_electronico = ""
                    End If

                Next
                .To.Add(correo_electronico)
            End If

            .Subject = "Documentos Emitidos Día " + devuelve_fecha2(buscaHoraServidor().AddDays(-1))
            .IsBodyHtml = True
            .AlternateViews.Add(BodyEmitidos())
            .BodyEncoding = System.Text.Encoding.UTF8

            Dim ruta As String = Retorna_Ruta_ArchivoEmitidos()
            If ruta = "" Then
                Return True
                Exit Function
            End If
            adjunto = New System.Net.Mail.Attachment(ruta)
            .Attachments.Add(adjunto)
        End With

        Try
            smtp.Send(correo)

            Dim sql As String = " INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora) " +
                                " VALUES ('EMITIDOS','" + devuelve_fecha2(buscaHoraServidor().AddDays(-1)) + "','" + DevuelveHora() + "')"
            MovimientoSQL(sql)

            EnviarEmitidos = True
        Catch ex As Exception
            If (ex.Message.Trim = "No se puede enviar a un destinatario.") Then
                Dim sql As String = " INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora) " +
                                " VALUES ('EMITIDOS','" + devuelve_fecha2(buscaHoraServidor().AddDays(-1)) + "','" + DevuelveHora() + "')"
                MovimientoSQL(sql)

                EnviarEmitidos = True
            Else
                EnviarEmitidos = False
            End If
        End Try

        Return EnviarEmitidos

    End Function

    Function Retorna_Ruta_ArchivoEmitidos() As String

        Dim Fecha = devuelve_fecha2(buscaHoraServidor().AddDays(-1)).Replace("/", "_")
        Dim x = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\DocumentosEmitidos " + Fecha + ".Pdf"
        If Not File.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\DocumentosEmitidos " + Fecha + ".Pdf") Then
            Dim report As New Rpt_DocumentosEmitidos
            Try
                report.SetParameterValue("fecha", devuelve_fecha(buscaHoraServidor().AddDays(-1)))
                report.SetDatabaseLogon(setLogon, setPass)
            Catch ex As CrystalReportsException
                Return ""
                Exit Function
            End Try

            Try
                Dim CrExportOptions As ExportOptions
                Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
                Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()

                CrDiskFileDestinationOptions.DiskFileName = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\DocumentosEmitidos " + Fecha + ".Pdf"
                CrExportOptions = report.ExportOptions
                With CrExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.PortableDocFormat
                    .DestinationOptions = CrDiskFileDestinationOptions
                    .FormatOptions = CrFormatTypeOptions
                End With
                report.Export()
                report.Dispose()
                Return My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\DocumentosEmitidos " + Fecha + ".Pdf"
                Exit Function
            Catch ex As Exception
                Return ""
                Exit Function
            End Try
        End If
        Return ""

    End Function

    Private Function BodyEmitidos() As AlternateView

        Dim archivo As String = ArchivoAString(Application.StartupPath + "\Emitidos.txt")

        Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim htmlView2 As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim cabecera As LinkedResource = New LinkedResource(Application.StartupPath + "\Emitidos.jpg", MediaTypeNames.Image.Jpeg)
        cabecera.ContentId = "titulo"

        htmlView.LinkedResources.Add(cabecera)

        Return htmlView

    End Function

    '-------------------------------
    'CORREO COBROS FRIGORIFICO ----- NO SE OCUPA
    '-------------------------------

    Public Function EnviarVas() As Boolean

        Dim smtp As New System.Net.Mail.SmtpClient
        Dim correo As New System.Net.Mail.MailMessage
        Dim adjunto As System.Net.Mail.Attachment


        With smtp
            .Port = puerto
            .Host = host_mail
            .Credentials = New System.Net.NetworkCredential(correoenvio, claveenvio)
            .EnableSsl = estadoSSL
        End With

        With correo

            .From = New System.Net.Mail.MailAddress(correomostrar)

            Dim sql As String = "SELECT corr_TO, corr_CCO, corr_BCC FROM Correos WHERE corr_codi='2'"
            Dim tb As DataTable = ListarTablasSQL(sql)

            If tb.Rows.Count > 0 Then

                Dim uu As String = tb.Rows(0)(0).ToString().Trim

                If QuitarCaracteres(uu.ToString()).Length < uu.ToString().Length Then
                    Dim correo_electronico As String = ""
                    For i As Integer = 0 To uu.Length - 1
                        If uu.Chars(i) <> ";" Then
                            correo_electronico = correo_electronico + uu.Chars(i)
                        Else
                            Console.WriteLine(correo_electronico)
                            .To.Add(correo_electronico)
                            correo_electronico = ""
                        End If

                    Next
                    .To.Add(correo_electronico)
                    Console.WriteLine(correo_electronico)
                End If

                uu = tb.Rows(0)(1).ToString().Trim

                If uu <> "" Then
                    If QuitarCaracteres(uu.ToString()).Length < uu.ToString().Length Then
                        Dim correo_electronico As String = ""
                        For i As Integer = 0 To uu.Length - 1
                            If uu.Chars(i) <> ";" Then
                                correo_electronico = correo_electronico + uu.Chars(i)
                            Else
                                Console.WriteLine(correo_electronico)
                                .CC.Add(correo_electronico)
                                correo_electronico = ""
                            End If

                        Next
                        .CC.Add(correo_electronico)
                        Console.WriteLine(correo_electronico)
                    End If
                End If

                uu = tb.Rows(0)(2).ToString().Trim
                If uu <> "" Then
                    If QuitarCaracteres(uu.ToString()).Length < uu.ToString().Length Then
                        Dim correo_electronico As String = ""
                        For i As Integer = 0 To uu.Length - 1
                            If uu.Chars(i) <> ";" Then
                                correo_electronico = correo_electronico + uu.Chars(i)
                            Else
                                Console.WriteLine(correo_electronico)
                                .Bcc.Add(correo_electronico)
                                correo_electronico = ""
                            End If

                        Next
                        .Bcc.Add(correo_electronico)
                        Console.WriteLine(correo_electronico)
                    End If
                End If
            Else
                Return False
                Exit Function
            End If

            .Subject = "Disponibilidad en camaras"
            .IsBodyHtml = True
            .AlternateViews.Add(BodyDisponibilidad())
            .BodyEncoding = System.Text.Encoding.UTF8

            Dim ruta As String = Retorna_Ruta_ArchivoDisponibilidad()
            If ruta = "" Then
                Return True
                Exit Function
            End If
            adjunto = New System.Net.Mail.Attachment(ruta)
            .Attachments.Add(adjunto)
        End With

        Try
            smtp.Send(correo)

            Dim sql As String = " INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora) " +
                                " VALUES ('VAS','" + devuelve_fecha(buscaHoraServidor()) + "','" + DevuelveHora() + "')"
            MovimientoSQL(sql)

            EnviarVas = True
        Catch ex As Exception
            EnviarVas = False
        End Try

        Return EnviarVas

    End Function

    Function Retorna_Ruta_ArchivoVas(ByVal codigo As String) As String


        Dim x = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\vas " + codigo + ".Pdf"
        If Not File.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\vas " + codigo + ".Pdf") Then
            Dim report As New Rpt_Posiciones
            report.SetDatabaseLogon(setLogon, setPass)

            Try
                Dim CrExportOptions As ExportOptions
                Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
                Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()

                CrDiskFileDestinationOptions.DiskFileName = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\vas " + codigo + ".Pdf"
                CrExportOptions = report.ExportOptions
                With CrExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.PortableDocFormat
                    .DestinationOptions = CrDiskFileDestinationOptions
                    .FormatOptions = CrFormatTypeOptions
                End With
                report.Export()
                report.Dispose()
                Return My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\vas " + codigo + ".Pdf"
                Exit Function
            Catch ex As Exception
                Return ""
                Exit Function
            End Try
        End If
        Return ""

    End Function

    Private Function BodyVas() As AlternateView

        Dim archivo As String = ArchivoAString(Application.StartupPath + "\Vas.txt")

        Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim htmlView2 As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim cabecera As LinkedResource = New LinkedResource(Application.StartupPath + "\titulovas.jpg", MediaTypeNames.Image.Jpeg)
        cabecera.ContentId = "titulo"

        htmlView.LinkedResources.Add(cabecera)

        Return htmlView

    End Function

    '-------------------------------
    'SAPEO FRIGORIFICO   -----------
    '-------------------------------

    Public Function SinTermino(ByVal inf_cod As String) As Boolean

        Dim smtp As New System.Net.Mail.SmtpClient
        Dim correo As New System.Net.Mail.MailMessage
        Dim adjunto As System.Net.Mail.Attachment

        With smtp
            .Port = puerto
            .Host = host_mail
            .Credentials = New System.Net.NetworkCredential(correoenvio, claveenvio)
            .EnableSsl = estadoSSL
        End With

        With correo

            Dim sql As String = "SELECT prg_mail FROM informes_programa WHERE prg_inf_cod='" + inf_cod + "' AND prg_emp='0'"

            Dim tablaInterno As DataTable = ListarTablasSQL(sql)


            .From = New System.Net.Mail.MailAddress(correomostrar)

            Dim uu As String = tablaInterno.Rows(0)(0).ToString().Trim

            If QuitarCaracteres(uu.ToString()).Length < uu.ToString().Length Then
                Dim correo_electronico As String = ""
                For i As Integer = 0 To uu.Length - 1
                    If uu.Chars(i) <> ";" Then
                        correo_electronico = correo_electronico + uu.Chars(i)
                    Else
                        .To.Add(correo_electronico)
                        correo_electronico = ""
                    End If

                Next
                .To.Add(correo_electronico)
            End If

            .Subject = "Procesos sin finalizar al día " + devuelve_fecha2(buscaHoraServidor().AddDays(-1))
            .IsBodyHtml = True
            .AlternateViews.Add(BodySinTermino())
            .BodyEncoding = System.Text.Encoding.UTF8
            Dim ruta As String = Retorna_Ruta_SinTermino()
            If ruta = "" Then
                Return True
                Exit Function
            End If
            adjunto = New System.Net.Mail.Attachment(ruta)
            .Attachments.Add(adjunto)
        End With

        Try
            smtp.Send(correo)
            ' Elimina_Archivo(pedido_largo)
            Dim sql As String = " INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora) " +
                                " VALUES ('SINTERMINAR','" + devuelve_fecha2(buscaHoraServidor().AddDays(-1)) + "','" + DevuelveHora() + "')"
            MovimientoSQL(sql)

            SinTermino = True
        Catch ex As Exception
            If (ex.Message.Trim = "No se puede enviar a un destinatario.") Then
                Dim sql As String = " INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora) " +
                                " VALUES ('SINTERMINAR','" + devuelve_fecha2(buscaHoraServidor().AddDays(-1)) + "','" + DevuelveHora() + "')"
                MovimientoSQL(sql)

                SinTermino = True
            Else
                SinTermino = False
            End If
        End Try

        Return EnviarEmitidos(inf_cod)

    End Function

    Function Retorna_Ruta_SinTermino() As String

        Dim Fecha = devuelve_fecha2(buscaHoraServidor().AddDays(-1)).Replace("/", "_")
        Dim x = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\frigo " + Fecha + ".Pdf"
        If Not File.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\frigo " + Fecha + ".Pdf") Then
            Dim report As New Rpt_AnormalFrigo
            Try
                report.SetDatabaseLogon(setLogon, setPass)
            Catch ex As CrystalReportsException
                Return ""
                Exit Function
            End Try


            Try
                Dim CrExportOptions As ExportOptions
                Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
                Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions()

                CrDiskFileDestinationOptions.DiskFileName = My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\frigo " + Fecha + ".Pdf"
                CrExportOptions = report.ExportOptions
                With CrExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.PortableDocFormat
                    .DestinationOptions = CrDiskFileDestinationOptions
                    .FormatOptions = CrFormatTypeOptions
                End With
                report.Export()
                report.Dispose()
                Return My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\frigo " + Fecha + ".Pdf"
                Exit Function
            Catch ex As Exception
                Return ""
                Exit Function
            End Try
        End If
        Return ""

    End Function

    Private Function BodySinTermino() As AlternateView

        Dim archivo As String = ArchivoAString(Application.StartupPath + "\SinTermino.txt")

        Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim htmlView2 As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim cabecera As LinkedResource = New LinkedResource(Application.StartupPath + "\SinTermino.jpg", MediaTypeNames.Image.Jpeg)
        cabecera.ContentId = "titulo"

        htmlView.LinkedResources.Add(cabecera)

        Return htmlView

    End Function

    '-------------------------------
    'PICKING------------------------
    '-------------------------------

    Public Function EnviarCorreoPicking(ByVal inf_cod As String) As Boolean

        Dim smtp As New System.Net.Mail.SmtpClient
        Dim correo As New System.Net.Mail.MailMessage
        Dim adjunto As System.Net.Mail.Attachment

        With smtp
            .Port = puerto
            .Host = host_mail
            .Credentials = New System.Net.NetworkCredential(correoenvio, claveenvio)
            .EnableSsl = estadoSSL
        End With

        With correo

            Dim sql As String = "SELECT prg_mail FROM informes_programa WHERE prg_inf_cod='" + inf_cod + "' AND prg_emp='0'"

            Dim tablaInterno As DataTable = ListarTablasSQL(sql)


            .From = New System.Net.Mail.MailAddress(correomostrar)

            Dim uu As String = tablaInterno.Rows(0)(0).ToString().Trim

            If QuitarCaracteres(uu.ToString()).Length < uu.ToString().Length Then
                Dim correo_electronico As String = ""
                For i As Integer = 0 To uu.Length - 1
                    If uu.Chars(i) <> ";" Then
                        correo_electronico = correo_electronico + uu.Chars(i)
                    Else
                        .To.Add(correo_electronico)
                        correo_electronico = ""
                    End If

                Next
                .To.Add(correo_electronico)
            End If

            .Subject = "Saldos no confirmados de picking " + devuelve_fecha2(buscaHoraServidor().AddDays(-1))
            .IsBodyHtml = True
            .AlternateViews.Add(BodyPicking())
            .BodyEncoding = System.Text.Encoding.UTF8
            Dim ruta As String = Retornar_Ruta_PickingExcel()
            If ruta = "" Then
                Return True
                Exit Function
            End If
            adjunto = New System.Net.Mail.Attachment(ruta)
            .Attachments.Add(adjunto)
        End With

        Try
            smtp.Send(correo)

            'Dim sql As String = " INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora) " +
            '                    " VALUES ('PICKING','" + devuelve_fecha2(buscaHoraServidor().AddDays(-1)) + "','" + DevuelveHora() + "')"
            'MovimientoSQL(sql)

            Dim sql As String = "UPDATE SALDO_PICK_CONF SET enviado='1' WHERE ISNULL(enviado,0)='0' AND saldo_estado='NO'"
            MovimientoSQL(sql)

            EnviarCorreoPicking = True
        Catch ex As FileNotFoundException
            If (ex.Message.Trim = "No se puede enviar a un destinatario.") Then
                Dim sql As String = "UPDATE SALDO_PICK_CONF SET enviado='1' WHERE ISNULL(enviado,0)='0' AND saldo_estado='NO'"
                MovimientoSQL(sql)

                EnviarCorreoPicking = True
            Else
                EnviarCorreoPicking = False
            End If
        Catch ex As Exception
            If (ex.Message.Trim = "No se puede enviar a un destinatario.") Then
                Dim sql As String = "UPDATE SALDO_PICK_CONF SET enviado='1' WHERE ISNULL(enviado,0)='0' AND saldo_estado='NO'"
                MovimientoSQL(sql)

                EnviarCorreoPicking = True
            Else
                EnviarCorreoPicking = False
            End If
        End Try
    End Function

    Private Function BodyPicking() As AlternateView

        Dim archivo As String = ArchivoAString(Application.StartupPath + "\Picking.txt")

        Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim htmlView2 As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim cabecera As LinkedResource = New LinkedResource(Application.StartupPath + "\Picking.jpg", MediaTypeNames.Image.Jpeg)
        cabecera.ContentId = "titulo"

        htmlView.LinkedResources.Add(cabecera)

        Return htmlView

    End Function

    Public Function Retornar_Ruta_PickingExcel() As String

        Dim exApp As New Microsoft.Office.Interop.Excel.Application
        Dim exLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim exHoja As Microsoft.Office.Interop.Excel.Worksheet
        Dim filacol As Integer = 0
        Dim ColorIndex As Integer = 40
        Try

            Dim sqlConsulta As String = ""
            Dim tablaConsulta As DataTable

            exLibro = exApp.Workbooks.Add
            Dim fechahora As String = "" + devuelve_fecha2(buscaHoraServidor().AddDays(-1)) + "-" + DevuelveHora() + ""

            '********************************************************************'
            '***************INICIO DE DETALLE DE SOPORTANTES*********************'
            '********************************************************************'

            sqlConsulta = "SELECT PALLET,RACD_UNIDADES,CAJAS,SALDOCONFI FROM VG_SALDO_CONFIRMAR2 WHERE ISNULL(enviado,0)='0' AND saldo_estado='NO'"

            tablaConsulta = ListarTablasSQL(sqlConsulta)


            exHoja = exLibro.ActiveSheet

            Dim ImageFileName As String = IO.Path.Combine(Application.StartupPath, "Imagen.JPG")
            My.Resources.precisa.Save(ImageFileName)
            exHoja.Shapes.AddPicture(ImageFileName, Microsoft.Office.Core.MsoTriState.msoCTrue, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 0, 175, 52)

            exHoja.Cells.Item(5, 1) = "PLANILLA DE SALDOS NO CONFIRMADOS DE PICKING"
            exHoja.Cells.Item(6, 1) = fechahora

            exHoja.Cells.Item(7, 1) = "PALLET"
            exHoja.Cells.Item(7, 2) = "TIENE"
            exHoja.Cells.Item(7, 3) = "PEDIDAS"
            exHoja.Cells.Item(7, 4) = "SALDO"


            exHoja.Cells.Item(7, 1).Interior.ColorIndex = ColorIndex
            exHoja.Cells.Item(7, 2).Interior.ColorIndex = ColorIndex
            exHoja.Cells.Item(7, 3).Interior.ColorIndex = ColorIndex
            exHoja.Cells.Item(7, 4).Interior.ColorIndex = ColorIndex


            For i = 0 To tablaConsulta.Rows.Count - 1
                filacol = i + 8
                exHoja.Cells.Item(filacol, 1) = "'" + tablaConsulta.Rows(i)(0).ToString()
                exHoja.Cells.Item(filacol, 2) = "'" + tablaConsulta.Rows(i)(1).ToString()
                exHoja.Cells.Item(filacol, 3) = tablaConsulta.Rows(i)(2).ToString()
                exHoja.Cells.Item(filacol, 4) = tablaConsulta.Rows(i)(3).ToString()
                exHoja.Rows.Item(filacol).HorizontalAlignment = 3
            Next

            filacol = filacol + 1

            exHoja.Cells.Item(filacol, 1).Interior.ColorIndex = ColorIndex
            exHoja.Cells.Item(filacol, 2).Interior.ColorIndex = ColorIndex
            exHoja.Cells.Item(filacol, 3).Interior.ColorIndex = ColorIndex
            exHoja.Cells.Item(filacol, 4).Interior.ColorIndex = ColorIndex

            exHoja.Rows.Item(7).Font.Bold = 1

            exHoja.Columns.AutoFit()


            '********************************************************************'
            '*****************FIN DE DETALLE DE SOPORTANTES**********************'
            '********************************************************************'

            'CERRAR PROCESOS DE EXCEL'
            Dim fh = fechahora.Replace("/", "").Replace(":", "").Replace(".", "").Replace("-", "")
            exApp.Application.ActiveWorkbook.SaveCopyAs(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\SaldoPicking_" + fh.ToString() + ".xlsx")
            exHoja = Nothing
            exLibro.Close(SaveChanges:=False)
            exLibro = Nothing
            exApp.Quit()
            exApp = Nothing

            'DESTRUIR PROCESOS DE EXCEL'
            Dim p As Process
            For Each p In Process.GetProcesses()
                If Not p Is Nothing And p.ProcessName = "EXCEL" Then
                    p.Kill() 'Cierra el proceso
                End If
            Next

            Return My.Computer.FileSystem.SpecialDirectories.MyDocuments + "\SaldoPicking_" + fh.ToString() + ".xlsx"
        Catch ex As FileNotFoundException
            Return ""
        Catch ex As Exception
            Return ""
        End Try
        Return ""
    End Function


    '       
    '       VES OCT 2019
    '       ENVIO DEL CORREO CON EL INFORME DE ESTADO DE TUNELES
    '
    Public Function EnviarInformeTuneles(ByVal inf_cod As String) As Boolean

        Dim smtp As New System.Net.Mail.SmtpClient
        Dim correo As New System.Net.Mail.MailMessage
        Dim adjunto As System.Net.Mail.Attachment
        Dim EnviarEmitidos As Boolean


        With smtp
            .Port = puerto
            .Host = host_mail
            .Credentials = New System.Net.NetworkCredential(correoenvio, claveenvio)
            .EnableSsl = estadoSSL
        End With

        With correo

            Dim sql As String = "SELECT prg_mail FROM informes_programa WHERE prg_inf_cod='" + inf_cod + "' AND prg_emp='0'"

            Dim tablaInterno As DataTable = ListarTablasSQL(sql)


            .From = New System.Net.Mail.MailAddress(correomostrar)

            Dim uu As String = tablaInterno.Rows(0)(0).ToString().Trim

            If QuitarCaracteres(uu.ToString()).Length < uu.ToString().Length Then
                Dim correo_electronico As String = ""
                For i As Integer = 0 To uu.Length - 1
                    If uu.Chars(i) <> ";" Then
                        correo_electronico = correo_electronico + uu.Chars(i)
                    Else
                        .To.Add(correo_electronico)
                        correo_electronico = ""
                    End If

                Next
                .To.Add(correo_electronico)
            End If

            .Subject = "Documentos Emitidos Día " + devuelve_fecha2(buscaHoraServidor().AddDays(-1))
            .IsBodyHtml = True
            .AlternateViews.Add(BodyEstadoTuneles())
            .BodyEncoding = System.Text.Encoding.UTF8

            Dim ruta As String = Retorna_Ruta_ArchivoEmitidos()
            If ruta = "" Then
                Return True
                Exit Function
            End If
            adjunto = New System.Net.Mail.Attachment(ruta)
            .Attachments.Add(adjunto)
        End With

        Try
            smtp.Send(correo)

            Dim sql As String = " INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora) " +
                                " VALUES ('EMITIDOS','" + devuelve_fecha2(buscaHoraServidor().AddDays(-1)) + "','" + DevuelveHora() + "')"
            MovimientoSQL(sql)

            EnviarEmitidos = True
        Catch ex As Exception
            If (ex.Message.Trim = "No se puede enviar a un destinatario.") Then
                Dim sql As String = " INSERT INTO DocumentosEnviados(Denv_seccion, Denv_fecha, denv_hora) " +
                                " VALUES ('EMITIDOS','" + devuelve_fecha2(buscaHoraServidor().AddDays(-1)) + "','" + DevuelveHora() + "')"
                MovimientoSQL(sql)

                EnviarEmitidos = True
            Else
                EnviarEmitidos = False
            End If
        End Try

        Return EnviarEmitidos

    End Function

    '
    '       VES OCT 2019
    '       CORREO DE ESTADO DE TUNELES
    '
    Private Function BodyEstadoTuneles() As AlternateView

        Dim archivo As String = ArchivoAString(Application.StartupPath + "\estadoTuneles.txt")
        Dim hora As String = ""
        Dim fecha As String = ""
        Dim observacion As String = ""
        Dim transporte As String = ""
        Dim destino As String = ""
        Dim datos As String = "SELECT cam_descr, informeTuneles FROM vwTueneles WHERE cam_tipo = 2 ORDER BY cam_codi"
        Dim tablaDatos As DataTable = ListarTablasSQL(datos)
        If tablaDatos.Rows.Count > 0 Then
            hora = tablaDatos.Rows(0)(0).ToString()
            fecha = tablaDatos.Rows(0)(1).ToString()
            observacion = tablaDatos.Rows(0)(2).ToString()
            transporte = tablaDatos.Rows(0)(3).ToString()
            destino = tablaDatos.Rows(0)(4).ToString()
        End If

        Dim html = "<h1>ESTADO DE TUNELES</h1>" +
                        "<div id='tuneles'><div>" +
                        "<table border='0'>" +
                        "<tr>" +
                        "<th>Tunel</th>" +
                        "<th>Estado</th>" +
                        "</tr>"
        For Each row As DataRow In tablaDatos.Rows
            html = html + "<tr>" +
                            "<td>" + row("cam_descr") + "</td>" +
                            "<td>" + row("informeTuneles") + "</td>" +
                          "</tr>"
        Next
        html = html + "</table>"

        Dim archivo2 As String = ArchivoAString(Application.StartupPath + "\EstadoTuneles2.txt")

        archivo = archivo + html + archivo2

        Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim htmlView2 As AlternateView = AlternateView.CreateAlternateViewFromString(archivo, Encoding.UTF8, MediaTypeNames.Text.Html)
        Dim cabecera As LinkedResource = New LinkedResource(Application.StartupPath + "\tituloEstadoTUnel.jpg", MediaTypeNames.Image.Jpeg)
        cabecera.ContentId = "informeTuneles"

        htmlView.LinkedResources.Add(cabecera)

        Return htmlView

    End Function
End Class