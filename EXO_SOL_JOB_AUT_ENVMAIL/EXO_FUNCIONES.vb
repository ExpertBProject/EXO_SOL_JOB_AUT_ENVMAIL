Imports System.IO
Imports System.Reflection
Public Class EXO_FUNCIONES
    Public Shared Function LeerEmbebido(ByVal File As String) As String

        'Dim x As String = New IO.StreamReader(System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(File), System.Text.Encoding.UTF8).ReadToEnd
        'Return x
        Dim assmbly As Assembly = Assembly.GetExecutingAssembly()
        Dim reader As New StreamReader(assmbly.GetManifestResourceStream(File))
        Return reader.ReadToEnd()
    End Function
End Class
