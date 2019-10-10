Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System.Data

Public Module Funciones

    Dim con As New Conexion

    '-------------------- ÇRUTE SERVIDOR SQL SERVER ---------------------->

    Public Function ListarTablasSQL(ByVal Consulta_sql As String)
        Dim tabla As DataTable = New DataTable
        Try
            con.conectar()

            Dim Listado As SqlDataAdapter = New SqlDataAdapter(Consulta_sql, con.con)

            Listado.Fill(tabla)

            con.cerrar()
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
        'retornar tabla con datos
        Return tabla

    End Function

    Public Function MovimientoSQL(ByVal Consulta_sql)
        Dim retorno As Integer = 0

        Try
            con.conectar()

            Dim _cmd As SqlCommand = New SqlCommand(Consulta_sql, con.con)
            _cmd.ExecuteNonQuery()
            retorno = 1
            con.cerrar()

        Catch ex As Exception
            retorno = 0
            MsgBox(ex.ToString())
        End Try

        ' retornar 
        '1 si se ejecuta correctamente
        '0 si no se ejecuta

        Return retorno
    End Function

    Public Function UltimoRegistro(ByVal Consulta_sql)
        Dim _tabla As DataTable = New DataTable
        Dim _ultimoRegistro = 0
        Try
            con.conectar()

            Dim _cmd As SqlDataAdapter = New SqlDataAdapter(Consulta_sql, con.con)
            _cmd.Fill(_tabla)

            If _tabla.Rows.Count = 0 Then
                _ultimoRegistro = 1
            Else
                _ultimoRegistro = Val(_tabla.Rows(_tabla.Rows.Count - 1)(0) + 1)
            End If
            con.cerrar()


        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
        'retornar ultimo registro de la tabla 
        Return _ultimoRegistro
    End Function

    Public Function verificaExistencia(ByVal Tabla As String,
                                       ByVal Campo As String, ByVal valor As String) As Boolean

        Dim _tabla As DataTable = New DataTable
        Dim _RegistroExiste As Boolean = False
        Try
            con.conectar()

            Dim _cmd As SqlDataAdapter = New SqlDataAdapter("SELECT * FROM " + Tabla + " WHERE " + Campo + "='" + valor + "'", con.con)
            _cmd.Fill(_tabla)

            If _tabla.Rows.Count = 0 Then
                _RegistroExiste = False
            Else
                _RegistroExiste = True
            End If

            con.cerrar()


        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
        'retornar ultimo registro de la tabla 
        Return _RegistroExiste


    End Function

    Public Function ValorMaximo(ByVal Tabla As String, ByVal Campo As String, ByVal ValorSumar As Integer) As String

        Dim _tabla As DataTable = New DataTable
        Dim _valorMaximo As Integer = 0

        Try
            con.conectar()

            Dim _cmd As SqlDataAdapter = New SqlDataAdapter("SELECT Max(" + Campo + ") FROM " + Tabla, con.con)
            _cmd.Fill(_tabla)

            If _tabla.Rows.Count > 0 Then
                _valorMaximo = Convert.ToInt32(_tabla.Rows(0)(0))
            Else
                _valorMaximo = 0
            End If
            con.cerrar()


        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try

        Return (Convert.ToInt32(_valorMaximo) + Convert.ToInt32(ValorSumar))
    End Function

    Public Function ValorMaximoCondicional(ByVal Tabla As String, ByVal Campo As String, ByVal CondicionConWhere As String, ByVal ValorSumar As Integer) As String

        Dim _tabla As DataTable = New DataTable
        Dim _valorMaximo As Integer = 0

        Try
            con.conectar()

            Dim _cmd As SqlDataAdapter = New SqlDataAdapter("SELECT Max(" + Campo + ") FROM " + Tabla + " " + CondicionConWhere, con.con)
            _cmd.Fill(_tabla)

            If _tabla.Rows.Count > 0 Then
                _valorMaximo = Convert.ToInt32(_tabla.Rows(0)(0))
            Else
                _valorMaximo = 0
            End If
            con.cerrar()


        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try

        Return (Convert.ToInt32(_valorMaximo) + Convert.ToInt32(ValorSumar))
    End Function

    Public Function verificaExistenciaCondicional(ByVal Tabla As String,
                                  ByVal Condicion As String) As Boolean

        Dim _tabla As DataTable = New DataTable
        Dim _RegistroExiste As Boolean = False
        Try
            con.conectar()
            Dim sql As String = "SELECT * FROM " + Tabla + " WHERE " + Condicion + ""
            Dim _cmd As SqlDataAdapter = New SqlDataAdapter(sql, con.con)
            _cmd.Fill(_tabla)

            If _tabla.Rows.Count = 0 Then
                _RegistroExiste = False
            Else
                _RegistroExiste = True
            End If

            con.cerrar()


        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
        'retornar ultimo registro de la tabla 
        Return _RegistroExiste


    End Function

    Public Function DevuelveCorrelativo(ByVal codigo As String) As String
        DevuelveCorrelativo = ""
        Dim sql As String = "SELECT cor_correact FROM correlat WHERE cor_codi='" + codigo.ToString() + "'"
        Dim tabla As DataTable = ListarTablasSQL(sql)
        If tabla.Rows.Count > 0 Then
            DevuelveCorrelativo = tabla.Rows(0)(0).ToString()
        Else
            MsgBox("Error al rescatar correlativo")
        End If
        Return DevuelveCorrelativo
    End Function

    Public Function buscaHoraServidor() As DateTime

        Dim tabla As DataTable = New DataTable
        Try
            con.conectar()
            Dim Listado As SqlDataAdapter = New SqlDataAdapter("SELECT GETDATE()", con.con)
            Listado.Fill(tabla)
            con.cerrar()
        Catch ex As Exception
            '  MsgBox(ex.Message())
        End Try
        'retornar tabla con datos
        Dim fecha As DateTime
        If tabla.Rows.Count > 0 Then
            fecha = Convert.ToDateTime(tabla.Rows(0)(0).ToString())
        Else
            Return DateTime.Now
        End If

        'Dim fecha As DateTime = Convert.ToDateTime("20-03-2015 14:25:00.000")
        Return fecha

    End Function

    Public Function comprobarCorreos(ByVal campo As String, ByVal tabla As String, ByVal codigo As String, ByVal id As String) As Boolean
        Try

            Dim SQLInforme As String
            Dim registros As DataTable

            SQLInforme = "SELECT " + campo + " FROM " + tabla + " WHERE " + codigo + "='" + id + "'"
            registros = ListarTablasSQL(SQLInforme)

            If registros.Rows.Count > 0 Then
                Try
                    If Trim(registros.Rows(0)(0)) <> "" Then
                        Return True
                    Else
                        Return False
                    End If
                Catch ex As InvalidCastException
                    Return False
                End Try
            Else
                Return False
            End If


        Catch ex As IndexOutOfRangeException
            Return False
        End Try
    End Function


End Module
