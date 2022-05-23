Module JOB
    Public Sub Main()
        Dim iCountExeJOB As Integer = 0
        Dim sIDMax As String = "0"
        Dim oLog As EXO_Log.EXO_Log = Nothing
        Dim sError As String
        Dim sPath As String = ""
        Dim oFiles() As String = Nothing
        Dim sToken As String = ""

        Try
            sPath = My.Application.Info.DirectoryPath.ToString

            If Not System.IO.Directory.Exists(sPath & "\Logs") Then
                System.IO.Directory.CreateDirectory(sPath & "\Logs")
            End If
            oLog = New EXO_Log.EXO_Log(sPath & "\Logs\LOG_", 10, EXO_Log.EXO_Log.Nivel.todos, 4, "", EXO_Log.EXO_Log.GestionFichero.dia)
            oLog.escribeMensaje("", EXO_Log.EXO_Log.Tipo.informacion)
            oLog.escribeMensaje("#######################################################", EXO_Log.EXO_Log.Tipo.informacion)
            oLog.escribeMensaje("#######################################################", EXO_Log.EXO_Log.Tipo.informacion)
            oLog.escribeMensaje("#####        INICIO PROCESOS AUTORIZACIONES       #####", EXO_Log.EXO_Log.Tipo.informacion)

            If Conexiones.Datos_Confi("ACTUALIZAR", "CAMPOS") = "Y" Then
                oLog.escribeMensaje("Procedimiento. ACTUALIZAR CAMPO", EXO_Log.EXO_Log.Tipo.informacion)
                oLog.escribeMensaje("###################################################", EXO_Log.EXO_Log.Tipo.informacion)
                Procesos.Actualizar_Campos(oLog)
            ElseIf Conexiones.Datos_Confi("ACTUALIZAR", "VIASPAGO") = "Y" Then
                oLog.escribeMensaje("Procedimiento. ACTUALIZAR VIAS PAGO", EXO_Log.EXO_Log.Tipo.informacion)
                oLog.escribeMensaje("###################################################", EXO_Log.EXO_Log.Tipo.informacion)
                Procesos.Actualizar_ViasPago(oLog)
            Else
                Dim sHora1 As String = Conexiones.Datos_Confi("STATAPROB", "HORA1")
                Dim shora As String = Now.Hour.ToString("00") & ":" & Now.Minute.ToString("00")
                If sHora1 = shora Then
                    oLog.escribeMensaje("Procedimiento. Enviar mails status aprobados y rechazados", EXO_Log.EXO_Log.Tipo.informacion)
                    oLog.escribeMensaje("#######################################################", EXO_Log.EXO_Log.Tipo.informacion)
                    Procesos.EnviarMails_statusaprob(oLog)
                    oLog.escribeMensaje("#######################################################", EXO_Log.EXO_Log.Tipo.informacion)
                    oLog.escribeMensaje(" ", EXO_Log.EXO_Log.Tipo.informacion)
                End If

                oLog.escribeMensaje("Procedimiento. Enviar mails Autorizaciones pendientes", EXO_Log.EXO_Log.Tipo.informacion)
                oLog.escribeMensaje("#######################################################", EXO_Log.EXO_Log.Tipo.informacion)
                Procesos.EnviarMails(oLog)
                oLog.escribeMensaje("#######################################################", EXO_Log.EXO_Log.Tipo.informacion)
                oLog.escribeMensaje(" ", EXO_Log.EXO_Log.Tipo.informacion)


            End If
        Catch ex As Exception
            If ex.InnerException IsNot Nothing AndAlso ex.InnerException.Message <> "" Then
                sError = ex.InnerException.Message
            Else
                sError = ex.Message
            End If
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally
            oLog.escribeMensaje("#####                 FIN PROCESO                 #####", EXO_Log.EXO_Log.Tipo.informacion)
            oLog.escribeMensaje("#######################################################", EXO_Log.EXO_Log.Tipo.informacion)
            oLog.escribeMensaje("#######################################################", EXO_Log.EXO_Log.Tipo.informacion)
            oLog.escribeMensaje("", EXO_Log.EXO_Log.Tipo.informacion)
        End Try
    End Sub
End Module
