Public Class Form2

    Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long

    'Función Api GetShortPathName para obtener los paths de los archivos en formato corto para visual Basic 6.0
    'Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
    'Función Api GetShortPathName para obtener los paths de los archivos en formato corto para visual Basic Net

    Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" _
        (ByVal longPath As String, ByVal shortPath As String, ByVal shortBufferSize As Int32) As Int32
 
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            'Variable para obtener el resultado cuando disponemos la ejecución del mp3
            Dim Sound As Long
            'Asignamos a esta variable el path y nombre y extensión del mp3 que ejecutaremos
            Dim archivo = "C:\Users\Victor alejandro\Desktop\My Shared Folder\bon jovi - always.mp3"
            'Si es menor a la letra de la unidad mas los dos puntos, mas la barra invertida
            'mas las ocho letras de un archivo con nombre corto(Ver correspondencia con MSDOS)
            'mas el punto, mas la extensión que son tres letras, entonces...
            If archivo.Length > 13 Then
                'ejecutamos la api para reducir el path
                archivo = PathCorto(archivo)
            End If
            'ejecutamos la canción
            Sound = mciExecute("Play " & archivo)
        Catch ex As Exception
            MsgBox("Ha ocurrido un error!", MsgBoxStyle.Critical, "Nombre de tu programa")
        End Try
    End Sub

    'Sub que obtiene el path corto del archivo a reproducir
    Function PathCorto(ByVal archivo As String) As String
        'Obtenemos el tamaño del string para crear el buffer correcto
        Dim longPathLength As Int32 = archivo.Length
        'Aquí pondremos el path corto, le asignamos el espacio necesario
        Dim shortPathName As String = Space(longPathLength)
        'Variable por si queremos controlar que nos devuelve como valor la api
        Dim returnValue As Int32
        'Llamamosa la función para la conversión
        returnValue = GetShortPathName(archivo, shortPathName, longPathLength)
        'Cargamos el valor al nombre de la función para que devuelva el resultado.
        PathCorto = shortPathName.ToString
    End Function

    Private Sub Form2_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            'Cerramos todos los dispositivos abiertos mediante esta API
            mciExecute("Close All")
        Catch ex As Exception
            MsgBox("Ha ocurrido un error!", MsgBoxStyle.Critical, "Nombre de tu programa")
        End Try
    End Sub

End Class
