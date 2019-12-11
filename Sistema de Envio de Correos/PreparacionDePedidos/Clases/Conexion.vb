Imports System.Data.SqlClient

Public Class Conexion

    Public con As New SqlConnection

    'CONEXION ORIGINAL
    Public Ip As String
    Public Base_wms As String
    Public Base_fac As String
    Public Usuario As String
    Public Clave As String


    
    Sub conexion()

        'CONEXION ORIGINAL
        Ip = "192.168.1.90\PRECISABD"
        Base_wms = "precisa"
        Base_fac = "precisabd"
        Usuario = "sa"
        Clave = "precisa"

        'CONEXTION VICTOR
        'Ip = "localhost"
        'Base_wms = "precisa"
        'Base_fac = "precisa"
        'Usuario = "sa"
        'Clave = "sa*2019"

        'CONEXION EXTERNA
        'Ip = "192.168.10.150"
        'Ip = "186.64.109.246"
        'Base_wms = "precisa"
        'Base_fac = "precisa_bd"
        'Usuario = "sa"
        'Clave = "Clave01*"

        'PRUEBAS
        'Base_wms = "precisa_backup"
        
    End Sub

    Sub conectar()
        conexion()
        Dim cadena As String = "Data Source=" & Ip & "; initial catalog=" & Base_wms & "; USER=" & Usuario & "; PWD=" & Clave & ";timeout=100;"
        Try
            If con.State = 0 Then

                con.ConnectionString = (cadena)
                con.Open()

            End If
        Catch ex As Exception
            'MsgBox(cadena)
            Application.Exit()
        End Try

    End Sub

    Sub cerrar()
        con.Close()
    End Sub

    Sub ConectarFact()
        conexion()
        Dim cadena As String = "Data Source=" & Ip & "; initial catalog=" & Base_fac & "; USER=" & Usuario & "; PWD=" & Clave & ";"
        Try
            If con.State = 0 Then
                con.ConnectionString = (cadena)
                con.Open()

            End If
        Catch ex As Exception
            ''MsgBox(ex.Message)
            Application.Exit()
        End Try

    End Sub

    Sub CerrarFact()
        con.Close()
    End Sub
End Class
