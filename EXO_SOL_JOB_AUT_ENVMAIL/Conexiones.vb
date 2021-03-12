Imports System.IO
Imports System.Xml
Imports Sap.Data.Hana

Public Class Conexiones
#Region "Variables globales"

    Public Shared _sSchema As String = ""
    Public Shared sServer As String = ""
    Public Shared sBBDD As String = ""
    Public Shared sUser As String = ""
    Public Shared sPwd As String = ""
    Public Shared sCadena As String = ""

#End Region
#Region "Datos de Configuración"
    Public Shared Function Datos_Confi(ByVal sTipo As String, ByVal sDato As String) As String
        Dim myStream As Stream = Nothing
        Dim Reader As XmlTextReader = Nothing
        Datos_Confi = ""
        Try
            myStream = File.OpenRead(My.Application.Info.DirectoryPath.ToString & "\Connections.xml")
            Reader = New XmlTextReader(myStream)
            myStream = Nothing
            While Reader.Read
                Select Case Reader.NodeType
                    Case XmlNodeType.Element
                        Select Case Reader.Name.ToString.Trim
                            Case sTipo
                                Datos_Confi = Reader.GetAttribute(sDato).ToString.Trim
                                Exit While
                        End Select
                End Select
            End While

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            myStream = Nothing
            Reader.Close()
            Reader = Nothing
        End Try
    End Function
#End Region

#Region "Connect to Company"
    Public Shared Sub Connect_Company(ByRef oCompany As SAPbobsCOM.Company, ByVal sClave As String, ByVal sBBDD As String, ByRef oLog As EXO_Log.EXO_Log)
        Dim myStream As Stream = Nothing
        Dim Reader As XmlTextReader = Nothing

        Try
            'Conectar DI SAP
            myStream = File.OpenRead(My.Application.Info.DirectoryPath.ToString & "\Connections.xml")
            Reader = New XmlTextReader(myStream)
            myStream = Nothing
            oLog.escribeMensaje("Leyendo cadena de conexión...", EXO_Log.EXO_Log.Tipo.advertencia)
            While Reader.Read
                Select Case Reader.NodeType
                    Case XmlNodeType.Element
                        Select Case Reader.Name.ToString.Trim
                            Case sClave
                                oCompany = New SAPbobsCOM.Company

                                oCompany.language = SAPbobsCOM.BoSuppLangs.ln_Spanish
                                oCompany.Server = Reader.GetAttribute("Server").ToString.Trim
                                oCompany.LicenseServer = Reader.GetAttribute("LicenseServer").ToString.Trim
                                oCompany.UserName = Reader.GetAttribute("UserName").ToString.Trim
                                oCompany.Password = Reader.GetAttribute("Password").ToString.Trim
                                oCompany.UseTrusted = False
                                oCompany.DbPassword = Reader.GetAttribute("DbPassword").ToString.Trim
                                oCompany.DbUserName = Reader.GetAttribute("DbUserName").ToString.Trim
                                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
                                'oCompany.CompanyDB = Reader.GetAttribute("CompanyDB").ToString.Trim
                                oCompany.CompanyDB = sBBDD
                                oLog.escribeMensaje("Conectando: " & Reader.GetAttribute("Server").ToString.Trim & " " & Reader.GetAttribute("LicenseServer").ToString.Trim & " " & Reader.GetAttribute("UserName").ToString.Trim & " " & Reader.GetAttribute("Password").ToString.Trim & " " & Reader.GetAttribute("DbPassword").ToString.Trim & " " & Reader.GetAttribute("DbUserName").ToString.Trim & " " & sBBDD, EXO_Log.EXO_Log.Tipo.error)


                                If oCompany.Connect <> 0 Then
                                    oLog.escribeMensaje("Error en la conexión a la compañia:" & oCompany.GetLastErrorDescription.Trim, EXO_Log.EXO_Log.Tipo.error)
                                    Throw New System.Exception("Error en la conexión a la compañia:" & oCompany.GetLastErrorDescription.Trim)
                                End If
                                Exit While
                        End Select
                End Select
            End While
            oLog.escribeMensaje("Conectado a la compañia " & oCompany.CompanyName, EXO_Log.EXO_Log.Tipo.advertencia)

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            myStream = Nothing
            Reader.Close()
            Reader = Nothing
        End Try
    End Sub
    Public Shared Sub Disconnect_Company(ByRef oCompany As SAPbobsCOM.Company)
        Try
            If Not oCompany Is Nothing Then
                If oCompany.Connected = True Then
                    oCompany.Disconnect()
                End If
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If oCompany IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCompany)
            oCompany = Nothing
        End Try
    End Sub

#End Region
#Region "Connect to HANA"

    Public Shared Sub Connect_SQLHANA(ByRef db As HanaConnection, ByVal sClave As String, ByRef oLog As EXO_Log.EXO_Log)
        Dim myStream As Stream = Nothing
        Dim Reader As XmlTextReader = Nothing

        Try
            'Conectar SQL
            myStream = File.OpenRead(My.Application.Info.DirectoryPath.ToString & "\Connections.xml")
            Reader = New XmlTextReader(myStream)
            myStream = Nothing
            While Reader.Read
                Select Case Reader.NodeType
                    Case XmlNodeType.Element
                        Select Case Reader.Name.ToString.Trim
                            Case sClave
                                If db Is Nothing OrElse db.State = ConnectionState.Closed Then
                                    db = New HanaConnection
                                    db.ConnectionString = "Server=" & Reader.GetAttribute("Server").ToString.Trim & ";UserID=" & Reader.GetAttribute("DbUser").ToString & ";Password=" & Reader.GetAttribute("DbPwd").ToString & ";"
                                    db.Open()

                                    'variable a añadir en las sql hana
                                    _sSchema = Reader.GetAttribute("databaseName").ToString.Trim

                                    sServer = Reader.GetAttribute("Server").ToString.Trim
                                    sBBDD = Reader.GetAttribute("databaseName").ToString.Trim
                                    sUser = Reader.GetAttribute("DbUser").ToString.Trim
                                    sPwd = Reader.GetAttribute("DbPwd").ToString.Trim
                                    sCadena = "Server=" & Reader.GetAttribute("Server").ToString.Trim & ";UserID=" & Reader.GetAttribute("DbUser").ToString & ";Password=" & Reader.GetAttribute("DbPwd").ToString & ";"
                                End If
                                Exit While
                        End Select
                End Select
            End While
            oLog.escribeMensaje("Ha conectado a BBDD " & sBBDD, EXO_Log.EXO_Log.Tipo.advertencia)
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            myStream = Nothing
            Reader.Close()
            Reader = Nothing
        End Try
    End Sub

    Public Shared Sub Disconnect_SQLHANA(ByRef db As HanaConnection)
        Try
            If Not db Is Nothing AndAlso db.State = ConnectionState.Open Then
                db.Close()
                db.Dispose()
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            db = Nothing
        End Try
    End Sub

    Public Shared Function GetValueDB(ByRef db As HanaConnection, ByRef sTable As String, ByRef sField As String, ByRef sCondition As String) As String
        Dim cmd As HanaCommand = Nothing
        Dim da As HanaDataAdapter = Nothing
        Dim dt As System.Data.DataTable = Nothing
        Dim sSQL As String = ""

        Try
            If sCondition = "" Then
                sSQL = "SELECT " & sField & " FROM " & sTable
            Else
                sSQL = "SELECT " & sField & " FROM " & sTable & " WHERE " & sCondition
            End If

            cmd = New HanaCommand(sSQL, db)
            cmd.CommandTimeout = 0

            da = New HanaDataAdapter

            da.SelectCommand = cmd
            dt = New DataTable
            da.Fill(dt)

            If dt.Rows.Count <= 0 Then
                Return ""
            Else
                If Not IsDBNull(dt.Rows(0).Item(0).ToString) Then
                    Return dt.Rows(0).Item(0).ToString
                Else
                    Return ""
                End If
            End If

        Catch ex As Exception
            Throw ex
        Finally
            If Not dt Is Nothing Then
                dt.Dispose()
            End If

            If Not cmd Is Nothing Then
                cmd.Dispose()
            End If

            If Not da Is Nothing Then
                da.Dispose()
            End If
        End Try
    End Function

    Public Shared Sub FillDtDB(ByRef db As HanaConnection, ByRef dt As System.Data.DataTable, ByVal sSQL As String)
        Dim cmd As HanaCommand = Nothing
        Dim da As HanaDataAdapter = Nothing

        Try
            cmd = New HanaCommand(sSQL, db)
            cmd.CommandTimeout = 0

            da = New HanaDataAdapter

            da.SelectCommand = cmd
            da.Fill(dt)

        Catch ex As Exception
            Throw ex
        Finally
            If Not cmd Is Nothing Then
                cmd.Dispose()
            End If

            If Not da Is Nothing Then
                da.Dispose()
            End If
        End Try
    End Sub

    Public Shared Function ExecuteSqlDB(ByRef db As HanaConnection, ByVal sSQL As String) As Boolean
        Dim cmd As HanaCommand = Nothing

        ExecuteSqlDB = False

        Try
            cmd = New HanaCommand(sSQL, db)
            cmd.ExecuteNonQuery()

            ExecuteSqlDB = True

        Catch ex As Exception
            Throw ex
        Finally
            If Not cmd Is Nothing Then
                cmd.Dispose()
            End If
        End Try
    End Function

#End Region
End Class
