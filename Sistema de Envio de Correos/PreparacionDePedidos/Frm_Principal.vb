Imports System.Net
Imports Microsoft
Imports Microsoft.Win32
Imports Microsoft.Win32.Registry
Imports System.IO

Public Class Frm_Principal
    Dim segundos As Integer = 0
    Dim minutos As Integer = 0
    Dim horas As Integer = 0
    Dim env As New EnvioCorreos
    Dim horaLimite As String
    Dim horaInicio As String

    Private Function start_Up(ByVal bCrear As Boolean) As String

        ' clave del registro para   
        ' colocar el path del ejecutable para iniciar con windows  
        Const CLAVE As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"

        'ProductName : el nombre del programa.  
        Dim subClave As String = Application.ProductName.ToString
        ' Mensaje para retornar el resultado  
        Dim msg As String = ""

        Try
            ' Abre la clave del usuario actual (CurrentUser) para poder extablecer el dato  
            ' si la clave CurrentVersion\Run no existe la crea  
            Dim Registro As RegistryKey = CurrentUser.CreateSubKey(CLAVE, RegistryKeyPermissionCheck.ReadWriteSubTree)

            With Registro
                .OpenSubKey(CLAVE, True)

                Select Case bCrear
                    ' Crear  
                    ''''''''''''''''''''''  
                    Case True
                        ' Escribe el path con SetValue   
                        'Valores : ProductName el nombre del programa y ExecutablePath : la ruta del exe  
                        .SetValue(subClave, Application.ExecutablePath.ToString)
                        Return "Ok .. clave añadida"
                        ' Eliminar  
                        ''''''''''''''''''''''  
                        'Elimina la entrada con DeleteValue  
                    Case False
                        If .GetValue(subClave, "").ToString <> "" Then
                            .DeleteValue(subClave) ' eliminar  
                            msg = "Ok .. clave eliminada"
                        Else

                            msg = "No se eliminó , por que el programa" & _
                                   " no iniciaba con windows"
                        End If
                End Select
            End With
            ' Error  
            ''''''''''''''''''''''  
        Catch ex As Exception
            msg = ex.Message.ToString
        End Try
        'retorno  
        Return msg
    End Function

    'Declare our API functions 
    Declare Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, _
    ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long
    'Declare our constants 

    'SWP stands for SetWindowPos 

    'SWP_NoSize tells SetWindowPos to ignore the cx and cy arguments 
    Private Const SWP_NOSIZE = &H1

    'SWP_NoMove tells SetWindowPos to ignore the x and y arguments. 
    Private Const SWP_NOMOVE = &H2

    'HWND_TOPMOST is passed to SetWindowPos to set the target window Always On Top. 
    Private Const HWND_TOPMOST = -1

    'HWN_NOTOPMOST is passed to SetWindowPos to remove Always on Top 
    Private Const HWND_NOTOPMOST = -2

    'Declare our variables 
    Public Sub SetFormOnTop(ByVal myForm As Object)
        SetWindowPos(myForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    End Sub

    Private Sub Frm_Principal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'env.probarSMTP()
        'env.EnviarCorreoPedidoWeb("36412", "764075056", "4")
        'env.EnviarCorreoPedidoWeb("36413", "799842408", "4")
        'env.EnviarCorreoPedidosHora("12")
        'CorreosPedidos("7", "Mail2", "Clientes", "cli_rut")
        'env.EnviarPosicionesCamaras("10")
        'env.EnviarCorreoPedidosSemanal("11")
        'HABILITADOS
        Timer_contar_segundos.Start()
        Timer_Envia_Correos.Start() 'Inicia el cronometro en 1 minuto
        barra_envio.Maximum = 100
        comprobarExistencia()
        'prueba()
        'env.probarSMTP()
        'HABILITADOS
    End Sub

    Sub prueba()
        Try
            'CorreoPedidosWeb("4", "cli_mail", "Clientes", "cli_rut")
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Timer_Envia_Correos_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer_Envia_Correos.Tick
        'comprobarExistencia()
    End Sub

    Private Sub comprobarExistencia()


        '
        '       VES OCT 2019
        '       SE REPROGRAMO COMPLETAMENTE ESTE METODO PARA ESTANDARIZAR
        '       EL MAJEJO DE LA FRECUENCIA TANTO EN INFORMES DE CLIENTES
        '       COMO EN INFORMES INTERNOS
        '

        'env.EnviarInformeTuneles("22")
        'Return


        '
        '       VES OCT 2019
        '       OBTENEMOS UNA LISTA DE INFORMES EJECUTABLES, SEGUN LA FECHA
        '       Y HORA ACTUAL, CONFIGURACION DE FRECUENCIA DE LOS INFORMES
        '       Y LA FECHA DE ULTIMA EJECUCION DE CADA PROGRAMA DE INFORME.
        '

        Dim informes As DataTable = ListarTablasSQL("SELECT * FROM vwInformesPrograma ORDER BY prg_emp, prg_inf_cod, inf_pro_cod")
        If informes.Rows.Count = 0 Then  ' NADA QUE EJECUTAR?  FINALIZAMOS
            lblenvio.Visible = False
            barra_envio.Visible = False
            Console.WriteLine(DateTime.Now())
            Application.Exit()
            Me.Close()
        End If


        '
        '       VES OCT 2019
        '       RECORREMOS LA LISTA DE INFORMES A EJECUTAR Y SE
        '       VAN EJECUTANDO UNO POR UNO. LOS METODOS ORIGINALES
        '       enviarCorreosDiarios y  enviarCorreosAutomaticos
        '       SE INTEGRARON EN UN SOLO METODO procesarInforme.
        '
        lblenvio.Visible = True
        barra_envio.Visible = True

        barra_envio.Value = 0
        barra_envio.Maximum = informes.Rows.Count + 1
        For Each row As DataRow In informes.Rows
            Dim prg_cod As Integer = CInt(row("inf_pro_cod"))
            Dim inf_nom As String = row("inf_nom").ToString().Trim()
            Dim prg_interv As Integer = CInt(row("prg_interv"))
            Dim ipm_params As String = ""
            Dim ipm_id As Integer = 0

            '
            '       SI SE TRATA DE UN INFORME MANUAL, SE OBTIENE
            '       EL ID DE LA SIGUUIENTE ENTRADA EN LA COLA
            '       PARA DICHO INFORME
            '
            If prg_interv = 4 Then
                Dim ipm As DataTable = ListarTablasSQL("SELECT TOP 1 ipm_id, ipm_params FROM informes_manual WHERE inf_pro_cod = " + prg_cod.ToString() + " AND ipm_status = 'PENDIENTE' ORDER BY ipm_id")
                If ipm.Rows.Count > 0 Then
                    ipm_id = CInt(ipm.Rows(0)("ipm_id"))
                    ipm_params = ipm.Rows(0)("ipm_params").ToString()
                End If
            End If


            '
            '       SE PROCESA EL INFORME A EJECUTAR, ACTUALIZANDO
            '       LA FECHA/HORA DE ULTIMA EJECUCION.
            '
            '       SI EL INFORME ERA DE UN PROGRAMA MANUAL, SE MARCA
            '       LA ENTRADA EN LA COLA COMO PROCESADA
            '
            lblenvio.Text = inf_nom
            Try
                If (prg_interv <> 4 Or ipm_id > 0) Then
                    '
                    '       EJECUTAMOS EL INFORME
                    '
                    procesarInforme(row, ipm_id, ipm_params)

                    '
                    '       ACTUALIZAMOS LA FECHA DE ULTIMA
                    '       EJECUCION DEL INFORME
                    '
                    actualizarUltEjec(prg_cod)

                    '
                    '       SI SE TRATA DE UN PROGRAMA MANUAL SE
                    '       ACTUALIZA EL ESTATUS DE LA ENTRADA 
                    '       EB LA COLA 
                    '
                    If prg_interv = 4 And ipm_id > 0 Then
                        actualizarIPM(ipm_id)
                    End If
                End If


            Catch ex As Exception
            End Try
        Next


        'Modificacion HAmestica 211218. Pedidos Locales saldo informado incorrecto
        Dim sqlExiPedLocSalErr = "select count(ID) from V_Pedidos_Locales_Saldos_Incorrectos with(nolock)"
        Dim dtResp As DataTable = ListarTablasSQL(sqlExiPedLocSalErr)

        If (dtResp.Rows.Count > 0) Then
            enviarCorreosLocalPedidoSaldo()
        End If
        barra_envio.Increment(1)
        'Fin modificacion HAmestica 211218. Pedidos Locales saldo informado incorrecto

        lblenvio.Visible = False
        barra_envio.Visible = False

        Console.WriteLine(DateTime.Now())

        'eliminarArchivosTemporales()
        Application.Exit()
        Me.Close()
    End Sub

    'Modificacion HAmestica 211218. Pedidos Locales saldo informado incorrecto
    Sub enviarCorreosLocalPedidoSaldo()
        Try
            Dim SQLInforme As String
            Dim tabla As DataTable

            lblenvio.Visible = True
            barra_envio.Visible = True
            lblenvio.Text = "Enviando E-mails Pedidos Locales Saldo Incorrectos:"

            SQLInforme = "SELECT ID,Fecha_Pedido,Orden,cliente,cli_nomb,Pallet,Cajas_Pickeadas,Saldo_Informado,Saldo_Real,Trackeo,Codigo_Usuario,Usuario,Fecha_Preparacion FROM V_Pedidos_Locales_Saldos_Incorrectos WHERE Mail_Enviado='0'"
            tabla = ListarTablasSQL(SQLInforme)

            Dim Avance As Double = 0
            Dim saltoAvance As Double = 0

            If tabla.Rows.Count > 0 Then
                saltoAvance = Math.Floor(100 / tabla.Rows.Count)

                For i = 0 To tabla.Rows.Count - 1
                    Dim ID = tabla.Rows(i).Item(0).ToString.Trim
                    Dim FecPed = tabla.Rows(i).Item(1).ToString.Trim
                    Dim Ord = tabla.Rows(i).Item(2).ToString.Trim
                    Dim RutCli = tabla.Rows(i).Item(3).ToString.Trim
                    Dim NomCli = tabla.Rows(i).Item(4).ToString.Trim
                    Dim Pal = tabla.Rows(i).Item(5).ToString.Trim
                    Dim CajsPick = tabla.Rows(i).Item(6).ToString.Trim
                    Dim SalInf = tabla.Rows(i).Item(7).ToString.Trim
                    Dim SalRea = tabla.Rows(i).Item(8).ToString.Trim
                    Dim Tra = tabla.Rows(i).Item(9).ToString.Trim
                    Dim CodUsu = tabla.Rows(i).Item(10).ToString.Trim
                    Dim NomUsu = tabla.Rows(i).Item(11).ToString.Trim
                    Dim FecPre = tabla.Rows(i).Item(12).ToString.Trim

                    env.EnviarCorreoSaldosIncorrectos(ID, FecPed, Ord, RutCli, NomCli, Pal, CajsPick, SalInf, SalRea, Tra, NomUsu, FecPre)

                    Avance = Avance + saltoAvance
                    barra_envio.Increment(saltoAvance)
                Next

                barra_envio.Increment(100 - Avance)
            End If
        Catch ex As Exception

        End Try
    End Sub
    'Fin modificacion HAmestica 211218. Pedidos Locales saldo informado incorrecto

    Private Sub eliminarArchivosTemporales()
        'ELIMINAR TODOS LOS ARCHIVOS TEMPORALES DE LA CARPETA 2 DEL SERVIDOR'
        Dim result As String = Path.GetTempPath()
        result = result + "2\"
        Try
            For Each foundFile As String In My.Computer.FileSystem.GetFiles(result, Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories, "*.rpt")
                My.Computer.FileSystem.DeleteFile(foundFile,
                    Microsoft.VisualBasic.FileIO.UIOption.AllDialogs,
                    Microsoft.VisualBasic.FileIO.RecycleOption.DeletePermanently)
            Next
        Catch ex As DirectoryNotFoundException
            Exit Sub
        Catch ex As ArgumentException
            Exit Sub
        Catch ex As IOException
            Exit Sub
        Catch ex As NotSupportedException
            Exit Sub
        Catch ex As UnauthorizedAccessException
            Exit Sub
        End Try
        'FIN DE ELIMINAR TODOS LOS ARCHIVOS'
    End Sub

    Private Sub retornarHoras()

        Dim fec As Date = DateTime.Now
        Dim horaL As String = Convert.ToString(fec.AddMinutes(1))
        Dim cl As Integer = horaL.Length

        Dim horaI As String = Convert.ToString(fec.AddMinutes(-10))
        Dim ci As Integer = horaI.Length

        If horaL.Length = 18 Then
            horaLimite = horaL.Substring(11, 4)
            horaLimite = "0" + horaLimite
        Else
            horaLimite = horaL.Substring(11, 5)
        End If

        If horaI.Length = 18 Then
            horaInicio = horaI.Substring(11, 4)
            horaInicio = "0" + horaInicio
        Else
            horaInicio = horaI.Substring(11, 5)
        End If

    End Sub

    Private Sub enviarCorreosPrueba()
        Try
            Dim campo As String = verificarDia()
            Dim SQLInforme As String
            Dim tabla As DataTable
            lblenvio.Visible = True
            barra_envio.Visible = True
            lblenvio.Text = "Enviando E-mails Diarios:"

            SQLInforme = "SELECT inf_est, prg_mail FROM informes, informes_programa WHERE inf_cod='13' AND inf_cod= prg_inf_cod AND prg_estado='Activado' AND " + campo + "='si' AND (prg_hora>='" + horaInicio + "' AND prg_hora<='" + horaLimite + "')"
            tabla = ListarTablasSQL(SQLInforme)

            If tabla.Rows.Count > 0 Then
                If tabla.Rows(0)(0) = "Habilitado" And tabla.Rows(0)(1) <> "" Then
                    CorreoInformeSoportantes("13") 'prg_mail codigo:13
                End If
            End If

            barra_envio.Increment(100)
        Catch ex As Exception

        End Try
    End Sub

    '
    '       VES OCT 2019
    '       SE PROCESA EL INFORME SOLICITADO. ESTE METODO SIMPLIFICA
    '       Y REFACTORIZA EL CODIGO DE LOS METODOS ENVIARCORREOSDIARIOS
    '       Y ENVIARCORREOSAUTOMATICOS DEL CODIGO ORIGINAL
    '
    Private Sub procesarInforme(ByRef prog As DataRow, ByVal ipm_id As Integer, ByVal ipm_params As String)
        Dim inf_cod As Integer = CInt(prog("prg_inf_cod"))
        Dim prg_cod As Integer = CInt(prog("inf_pro_cod"))

        Select Case inf_cod
            Case 22
                env.EnviarCorreoStockComercialAgrosuper()

            Case 12
                CorreoPedidosHora("12")

            Case 11
                CorreoSemanal("11")

            Case 10
                CorreoPosiciones("10")

            Case 13
                CorreoInformeSoportantes("13")

            Case 3
                CorreoDocumentosEmitidos("3")

            Case 14
                CorreoSinTermino("14")

            Case 17
                CorreosPicking("17")

            Case 21
                env.EnviarCorreoAlertaRFID()

            Case 1
                CorreoRecepciones("1", "CLI_CRYD", "Clientes", "cli_rut") 'RECEPCION CHEKLIST DESPACHO CLI_CRYD codigo:1
                CorreoDespacho("1", "CLI_CRYD", "Clientes", "cli_rut") 'RECEPCION CHEKLIST DESPACHO CLI_CRYD codigo:1
                CorreoChecklist("1", "CLI_CRYD", "Clientes", "cli_rut") 'RECEPCION CHEKLIST DESPACHO CLI_CRYD codigo:1

            Case 7
                CorreosPedidos("7", "Mail2", "Clientes", "cli_rut")

            Case 5
                CorreoVencidos("5", "CLI_PVENC", "Clientes", "cli_rut")

            Case 4
                CorreoPedidosWeb("4", "cli_mail", "Clientes", "cli_rut")

            Case 18
                CorreoPedidosWeb24Hrs("18")

            Case 23
                'env.EnviarInformeTuneles(prg_cod)

            Case 24
                env.EnviarNotificacionInicioTunel(ipm_id)

            Case 25
                env.EnviarNotificacionFinalTunel(ipm_id)

            Case 26
                env.EnviarNotificacionRecepcionTardiaTunel(prog)

            Case 27
                env.EnviarAlertasInicioTunel(prog, 1)

            Case 28
                env.EnviarAlertasInicioTunel(prog, 2)

            Case 29
                env.EnviarAlertasInicioTunel(prog, 3)
        End Select

    End Sub



    '
    '       VES OCT 2019
    '       ACTUALIZA LA FECHA DE ULTIMA EJECCION PARA EL
    '       PROGRAMA DE INfORME INDICADO
    '
    Private Sub actualizarUltEjec(ByVal prg_cod As Integer)
        Dim sql As String
        sql = "UPDATE informes_programa SET prg_lastrun = GETDATE() WHERE inf_pro_cod = " + prg_cod.ToString()
        MovimientoSQL(sql)
    End Sub


    Private Sub actualizarIPM(ByVal ipm_id As Integer)
        Dim sql As String = "UPDATE informes_manual " +
                            "   SET ipm_status = 'PROCESADO'," +
                            "       ipm_fecsta = GETDATE()," +
                            "       ipm_ususta = '050' " +
                            " WHERE ipm_id = " + ipm_id.ToString() +
                            "   AND ipm_status = 'PENDIENTE'"
        MovimientoSQL(sql)
    End Sub

    Private Function verificarDia()

        '*********************************************************************
        'SE SELECCIONA EL DÍA PARA PODER INDICAR CUAL CAMPO SE CONSULTARA
        '*********************************************************************
        Dim Dia As String = ""
        Dim campo As String = ""

        Dim dateValue As Date = DateTime.Now()
        Dia = dateValue.DayOfWeek.ToString

        If Dia = "Monday" Then
            campo = "prg_lunes"
        End If

        If Dia = "Tuesday" Then
            campo = "prg_martes"
        End If

        If Dia = "Wednesday" Then
            campo = "prg_miercoles"
        End If

        If Dia = "Thursday" Then
            campo = "prg_jueves"
        End If

        If Dia = "Friday" Then
            campo = "prg_viernes"
        End If

        If Dia = "Saturday" Then
            campo = "prg_sabado"
        End If

        If Dia = "Sunday" Then
            campo = "prg_domingo"
        End If

        Return campo
    End Function

    Private Sub Frm_Principal_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        If WindowState = FormWindowState.Minimized Then
            Hide()
        End If
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        Show()
        WindowState = FormWindowState.Maximized
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MsgBox(start_Up(True))
    End Sub

    '***************************************
    Sub CorreosPedidos(ByVal inf_cod As String, ByVal ca As String, ByVal ta As String, ByVal co As String)
        Dim sql_terminados As String = "SELECT TOP 75 pf.pedido, pf.Orden, cliente+ '  '+ cli_nomb AS cliente, pf.fecha, pf.hora" +
            " FROM pedidos_ficha AS pf, informes_programa, clientes AS cl  WHERE cli_rut=cliente AND terminado='1'" +
              " AND cliente = prg_rut_cli" +
              " AND prg_estado = 'Activado'" +
              " AND prg_inf_cod='" + inf_cod + "'" +
              " ORDER BY pf.orden DESC"

        Dim tabla_terminados As DataTable = ListarTablasSQL(sql_terminados)

        If tabla_terminados.Rows.Count > 0 Then
            For i As Integer = 0 To tabla_terminados.Rows.Count - 1

                Dim existe As String = "SELECT * FROM pedidos_enviados WHERE ORDEN='" + tabla_terminados.Rows(i)(1).ToString() + "'"

                Dim tabla As DataTable = ListarTablasSQL(existe)
                Dim rut As String = tabla_terminados.Rows(i)(2).ToString()
                rut = rut.ToString.Remove(9, rut.Length - 9)

                If comprobarCorreos(ca, ta, co, rut) = True Then
                    If tabla.Rows.Count = 0 Then
                        While (env.EnviarCorreoPedidos(tabla_terminados.Rows(i)(0).ToString(), tabla_terminados.Rows(i)(1).ToString(), _
                                                       tabla_terminados.Rows(i)(2).ToString(), tabla_terminados.Rows(i)(3).ToString(), _
                                                       tabla_terminados.Rows(i)(4).ToString(), inf_cod) = False)
                        End While
                    End If
                End If
            Next
        End If
    End Sub
    '***************************************
    Sub CorreoRecepciones(ByVal inf_cod As String, ByVal ca As String, ByVal ta As String, ByVal co As String)

        Dim sql As String = "SELECT frec_codi, frec_rutcli FROM fichrece, informes_programa WHERE isnull(frec_enviada,0)='0' " +
            " AND frec_rutcond<>'222222222' " +
            " AND frec_rutcond<>'111111111' " +
            " AND frec_rutcond<>'000000035' " +
            " AND frec_rutcond<>'333333333' " +
            " AND frec_rutcond<>'444444444' " +
            " AND frec_rutcond<>'555555555' " +
            " AND frec_emis<='" + buscaHoraServidor().AddMinutes(-5) + "'" +
            " AND frec_rutcli = prg_rut_cli" +
            " AND prg_estado = 'Activado'" +
            " AND prg_inf_cod='1'"
        Dim tabla As DataTable = ListarTablasSQL(sql)

        For i As Integer = 0 To tabla.Rows.Count - 1

            ' verifica si esta el detalle

            Dim sql_verifica As String = "SELECT fr.frec_codi, COUNT(*) as count, frec_totsopo  FROM fichrece fr , detarece " +
                "WHERE frec_codi1=fr.frec_codi AND fr.frec_codi='" + tabla.Rows(i)(0).ToString() + "'  GROUP BY fr.frec_codi, fr.frec_totsopo"

            Dim tablaveri As DataTable = ListarTablasSQL(sql_verifica)

            If tablaveri.Rows.Count > 0 Then
                If tablaveri.Rows(0)(1).ToString() <> tablaveri.Rows(0)(2).ToString() Then
                    Exit Sub
                End If
            End If

            If comprobarCorreos(ca, ta, co, tabla.Rows(i)(1).ToString()) = True Then
                env.EnviarCorreoRecepcion(tabla.Rows(i)(0).ToString(), tabla.Rows(i)(1).ToString(), inf_cod)
            End If

        Next

    End Sub
    '***********************************
    Sub CorreoDespacho(ByVal inf_cod As String, ByVal ca As String, ByVal ta As String, ByVal co As String)

        Dim sql As String = "SELECT fdes_codi, fdes_rutcli FROM fichdespa, informes_programa  WHERE isnull(fdes_enviada,0)='0' AND " +
            " fdes_rutcond<>'222222222' " +
            " AND fdes_rutcond<>'111111111' " +
            " AND fdes_rutcond<>'000000035' " +
            " AND fdes_rutcond<>'333333333' " +
            " AND fdes_rutcond<>'444444444' " +
            " AND fdes_rutcond<>'555555555' " +
            " AND convert(date,fdes_emis)=convert(date,GETDATE())" +
            " AND fdes_rutcli = prg_rut_cli" +
            " AND prg_estado = 'Activado'" +
            " AND prg_inf_cod='" + inf_cod + "'"

        '" AND fdes_emis<='" + buscaHoraServidor().AddMinutes(-5) + "'" +

        ' MsgBox("Consulta: " + sql + "")

        Dim tabla As DataTable = ListarTablasSQL(sql)

        For i As Integer = 0 To tabla.Rows.Count - 1

            Dim sql_verifica As String = "SELECT fdes_codi, COUNT(*) as TOT, fdes_totsopo FROM fichdespa , detadespa " +
                "WHERE fdes_codi=LEFT(ddes_codi,7) AND fdes_codi='" + tabla.Rows(i)(0).ToString() + "' GROUP BY fdes_codi, fdes_totsopo "

            Dim tablaveri As DataTable = ListarTablasSQL(sql_verifica)

            If tablaveri.Rows.Count > 0 Then
                If tablaveri.Rows(0)(1).ToString() <> tablaveri.Rows(0)(2).ToString() Then
                    Exit Sub
                End If
            End If

            If comprobarCorreos(ca, ta, co, tabla.Rows(i)(1).ToString()) = True Then
                env.EnviarCorreoDespacho(tabla.Rows(i)(0).ToString(), tabla.Rows(i)(1).ToString(), inf_cod)
            End If

        Next

    End Sub
    '***********************************
    Sub CorreoPedidosHora(ByVal inf_cod As String)


        Dim sql As String = "SELECT * FROM DocumentosEnviados WHERE Denv_Seccion='PDIARIO' AND Denv_fecha='" + devuelve_fecha(buscaHoraServidor()) + "' "
        Dim tabla As DataTable = ListarTablasSQL(sql)

        If tabla.Rows.Count = 0 Then
            While (env.EnviarCorreoPedidosHora(inf_cod) = False)

            End While
        End If




    End Sub
    '***********************************
    Sub CorreoVencidos(ByVal inf_cod As String, ByVal ca As String, ByVal ta As String, ByVal co As String)

        Dim x = buscaHoraServidor().DayOfWeek

        If buscaHoraServidor().DayOfWeek = DayOfWeek.Monday Then


            'Dim SQLClientes As String = " SELECT cli_rut FROM Clientes, informes_programa WHERE cli_venc = '0'" +
            '                            " AND cli_rut = prg_rut_cli" +
            '                            " AND prg_estado = 'Activado'" +
            '                            " AND prg_inf_cod='5' AND cli_rut='774486305'"

            Dim SQLClientes As String = " SELECT cli_rut FROM Clientes, informes_programa WHERE cli_venc = '0'" +
                                        " AND cli_rut = prg_rut_cli" +
                                        " AND prg_estado = 'Activado'" +
                                        " AND prg_inf_cod='5'"

            Dim tablaClientes As DataTable = ListarTablasSQL(SQLClientes)

            For i As Integer = 0 To tablaClientes.Rows.Count - 1

                Dim sql As String = "SELECT * FROM DocumentosEnviados WHERE Denv_Seccion='PVENC' " +
                    "AND Denv_Dcto = '" + tablaClientes.Rows(i)(0).ToString() + "' AND Denv_fecha='" + devuelve_fecha(buscaHoraServidor()) + "' "

                Dim tabla As DataTable = ListarTablasSQL(sql)

                Dim rut As String = tablaClientes.Rows(i)(0).ToString()
                rut = rut.ToString.Remove(9, rut.Length - 9)

                If comprobarCorreos(ca, ta, co, rut) = True Then
                    If tabla.Rows.Count = 0 Then

                        Dim sql_Tiene_soportantes_vencidos As String = "SELECT  cli_nomb, racd_codi, drec_sopocli, racd_codpro,mae_descr,convert(date,racd_fecpro) AS fproduccion, " +
                            "convert(date,dr.fechavencimiento) AS fvencimiento , (DATEDIFF(day,convert(date,GETDATE()),convert(date,dr.fechavencimiento))) AS diasvencidos " +
                            "FROM rackdeta, maeprod , detarece as dr, clientes WHERE racd_codpro=mae_codi AND drec_codi=racd_codi AND cli_rut=drec_rutcli AND " +
                            "DATEDIFF(day,convert(date,GETDATE()),convert(date,dr.fechavencimiento))< 30 AND cli_venc='0' AND cli_rut='" + tablaClientes.Rows(i)(0).ToString() + "'"

                        Dim tablatiene As DataTable = ListarTablasSQL(sql_Tiene_soportantes_vencidos)

                        If tablatiene.Rows.Count > 0 Then
                            While (env.EnviarCorreoVencidos(tablaClientes.Rows(i)(0).ToString(), inf_cod) = False)

                            End While
                        End If
                    End If
                End If
            Next
        End If

    End Sub
    '***********************************
    'CorreoPedidosWeb("4", "cli_mail", "Clientes", "cli_rut")
    Sub CorreoPedidosWeb(ByVal inf_cod As String, ByVal ca As String, ByVal ta As String, ByVal co As String)
        Dim sql As String = "SELECT Orden, cliente FROM Pedidos_ficha, informes_programa" +
                            " WHERE cliente=prg_rut_cli AND ISNULL(ped_enviado,0)=0 AND prg_inf_cod='4' AND codvig='0' AND prg_estado = 'Activado'"

        Dim tabla As DataTable = ListarTablasSQL(sql)

        For i As Integer = 0 To tabla.Rows.Count - 1
            If comprobarCorreos(ca, ta, co, tabla.Rows(i)(1).ToString()) = True Then
                env.EnviarCorreoPedidoWeb(tabla.Rows(i)(0).ToString(), tabla.Rows(i)(1).ToString(), inf_cod)
            End If
        Next
    End Sub
    '***********************************

    '***********************************
    Sub CorreoPedidosWeb24Hrs(ByVal inf_cod As String)
        Dim sql As String = "select distinct Orden,Cliente from Pedidos_24_Horas with(nolock) where Mail_Enviado='0' and CONVERT(varchar,Fecha_Pedido,103)=CONVERT(varchar,GETDATE(),103)"

        Dim tabla As DataTable = ListarTablasSQL(sql)

        For i As Integer = 0 To tabla.Rows.Count - 1
            Dim Ord = tabla.Rows(i)(0).ToString.Trim
            Dim RutCli = tabla.Rows(i)(1).ToString.Trim

            env.EnviarCorreoPedidoWeb24Hrs(Ord, RutCli, inf_cod)
        Next
    End Sub
    '***********************************

    Sub CorreoSemanal(ByVal inf_cod As String)



        Dim sql As String = "SELECT * FROM DocumentosEnviados WHERE Denv_Seccion='PSEMANAL' AND Denv_fecha='" + devuelve_fecha(buscaHoraServidor()) + "' "
        Dim tabla As DataTable = ListarTablasSQL(sql)

        If tabla.Rows.Count = 0 Then
            While (env.EnviarCorreoPedidosSemanal(inf_cod) = False)

            End While
        End If

    End Sub
    '***********************************
    Sub CorreoChecklist(ByVal inf_cod As String, ByVal ca As String, ByVal ta As String, ByVal co As String)

        Dim sql As String = "SELECT cl_fol, cl_rutcli FROM zchecklist WHERE isnull(cl_enviada,0)='0' " +
            "AND cl_rutcli<>'116095327' AND cl_estpor='2' and CONVERT(date,Cl_Des)>=DATEADD(MONTH,-1,GETDATE()) ORDER BY cl_fol desc"

        Dim tabla As DataTable = ListarTablasSQL(sql)

        For i As Integer = 0 To tabla.Rows.Count - 1
            If comprobarCorreos(ca, ta, co, tabla.Rows(i)(1).ToString()) = True Then
                env.EnviarCorreoChecklist(tabla.Rows(i)(0).ToString(), tabla.Rows(i)(1).ToString(), inf_cod)
            End If
        Next

    End Sub
    '***********************************
    Sub CorreoPosiciones(ByVal inf_cod As String)


        Dim sql As String = "SELECT * FROM DocumentosEnviados WHERE Denv_Seccion='POSICIONES' AND Denv_fecha='" + devuelve_fecha(buscaHoraServidor()) + "' "
        Dim tabla As DataTable = ListarTablasSQL(sql)

        If tabla.Rows.Count = 0 Then
            While (env.EnviarPosicionesCamaras(inf_cod) = False)

            End While
        End If

    End Sub
    '***********************************
    Sub CorreoInformeSoportantes(ByVal inf_cod As String)

        Dim sql As String = "SELECT * FROM DocumentosEnviados WHERE Denv_Seccion='SOPORTANTES' AND Denv_fecha='" + devuelve_fecha(buscaHoraServidor()) + "' "
        Dim tabla As DataTable = ListarTablasSQL(sql)

        If tabla.Rows.Count = 0 Then
            While (env.EnviarInformeSoportantes(inf_cod) = False)

            End While
        End If

    End Sub
    '***********************************
    Sub CorreoDocumentosEmitidos(ByVal inf_cod As String)


        Dim sql As String = "SELECT * FROM DocumentosEnviados WHERE Denv_Seccion='EMITIDOS' AND Denv_fecha='" + devuelve_fecha2(buscaHoraServidor().AddDays(-1)) + "' "
        Dim tabla As DataTable = ListarTablasSQL(sql)

        If tabla.Rows.Count = 0 Then
            While (env.EnviarEmitidos(inf_cod) = False)

            End While
        End If

    End Sub
    '***********************************
    Sub CorreosPicking(ByVal inf_cod As String)


        Dim sql As String = "SELECT PALLET FROM VG_SALDO_CONFIRMAR2 WHERE ISNULL(enviado,0)='0' AND saldo_estado='NO'"
        Dim tabla As DataTable = ListarTablasSQL(sql)

        If tabla.Rows.Count > 0 Then
            If comprobarCorreos("prg_mail", "informes_programa", "prg_inf_cod", inf_cod) = True Then
                env.EnviarCorreoPicking(inf_cod)
            End If
        End If

    End Sub
    '***********************************
    Sub CorreoSinTermino(ByVal inf_cod As String)


        Dim sql As String = "SELECT * FROM DocumentosEnviados WHERE Denv_Seccion='SINTERMINAR' AND Denv_fecha='" + devuelve_fecha2(buscaHoraServidor().AddDays(-1)) + "' "
        Dim tabla As DataTable = ListarTablasSQL(sql)

        If tabla.Rows.Count = 0 Then
            While (env.SinTermino(inf_cod) = False)

            End While
        End If
    End Sub

    Private Sub SalirToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalirToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub Frm_Principal_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        NotifyIcon1.Dispose()
    End Sub

    Private Sub Timer_contar_segundos_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer_contar_segundos.Tick
        segundos = segundos + 1

        If segundos > 59 Then
            minutos = minutos + 1
            segundos = 0
        End If

        If minutos > 59 Then
            horas = horas + 1
            minutos = 0
        End If

        lbltiempo.Text = Format(horas, "00") & ":" & Format(minutos, "00") & ":" & Format(segundos, "00")
    End Sub
End Class
