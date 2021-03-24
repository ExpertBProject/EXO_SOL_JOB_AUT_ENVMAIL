Imports System.IO
Imports System.Text
Imports System.Xml
Imports EXO_DIAPI
Imports EXO_Log
Imports Sap.Data.Hana
Imports SAPbobsCOM

Public Class Procesos
    Public Shared Function FormateaString(ByVal dato As Object, ByVal tam As Integer) As String
        Dim retorno As String = String.Empty

        If dato IsNot Nothing Then
            retorno = dato.ToString
        End If

        If retorno.Length > tam Then
            retorno = retorno.Substring(0, tam)
        End If

        Return retorno.PadRight(tam, CChar(" "))
    End Function

#Region "Actualizar VIAS PAGO"
    Public Shared Sub Actualizar_ViasPago(ByRef oLog As EXO_Log.EXO_Log)
        Dim oDBSAP As HanaConnection = Nothing
        Dim oCompany As SAPbobsCOM.Company = Nothing
        Dim sError As String = ""
        Dim sSQL As String = ""
        Dim OdtDatos As System.Data.DataTable = Nothing
        Dim sPass As String = ""
        Dim refDI As EXO_DIAPI.EXO_DIAPI = Nothing
        Dim oXML As String = ""
        Dim sDir As String = Application.StartupPath
        Try
            sPass = Conexiones.Datos_Confi("DI", "Password")
            Dim tipoServidor As SAPbobsCOM.BoDataServerTypes = SAPbobsCOM.BoDataServerTypes.dst_HANADB
            Dim tipocliente As EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente = EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente.Clasico
            Conexiones.Connect_SQLHANA(oDBSAP, "HANA", oLog)
            sSQL = " SELECT ""EXO_BD"" FROM ""SOL_AUTORIZ"".""EXO_SOCIEDADES"" "
            OdtDatos = New System.Data.DataTable
            Conexiones.FillDtDB(oDBSAP, OdtDatos, sSQL)
            If OdtDatos.Rows.Count > 0 Then
                oLog.escribeMensaje("Se va a proceder a recorrer las SOCIEDADES...", EXO_Log.EXO_Log.Tipo.advertencia)
                For Each dr In OdtDatos.Rows
                    Conexiones.Connect_Company(oCompany, "DI", dr("EXO_BD").ToString, oLog)
                    oLog.escribeMensaje("Sociedad " + dr("EXO_BD").ToString, EXO_Log.EXO_Log.Tipo.advertencia)
                    Try
                        refDI = New EXO_DIAPI.EXO_DIAPI(oCompany, oLog)
                    Catch ex As Exception
                        refDI = New EXO_DIAPI.EXO_DIAPI(tipoServidor, oCompany.Server.ToString, oCompany.LicenseServer.ToString, oCompany.CompanyDB.ToString, oCompany.UserName.ToString, sPass, tipocliente)
                    End Try

                    'para cada cliente/proveedor activar todas las vias de pago
                    ActualizarViasClientes(oCompany, refDI, oLog)

                    Conexiones.Disconnect_Company(oCompany)
                Next
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            sError = exCOM.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            sError = ex.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally
            Conexiones.Disconnect_SQLHANA(oDBSAP)
            refDI = Nothing
        End Try

    End Sub

    Private Shared Sub ActualizarViasClientes(oCompany As Company, refDI As EXO_DIAPI.EXO_DIAPI, oLog As EXO_Log.EXO_Log)

        Dim dtClientes As System.Data.DataTable = Nothing
        Dim sSQL As String = ""

        Dim dtVias As System.Data.DataTable = Nothing

        Try

            sSQL = "SELECT T1.""PayMethCod"",T1.""Descript""  FROM  ""OPYM"" T1  " +
                " where T1.""Active""='Y' AND ""Type""='O'"
            dtVias = refDI.SQL.executeQuery(sSQL)

            sSQL = " SELECT ""CardCode"" from ""OCRD"" where ""frozenFor""='N' and ""CardType""='S' order by ""CardCode"""
            dtClientes = New System.Data.DataTable
            dtClientes = refDI.SQL.executeQuery(sSQL)

            If dtClientes.Rows.Count > 0 Then
                oLog.escribeMensaje("Se va a proceder a recorrer los clientes...", EXO_Log.EXO_Log.Tipo.advertencia)

                For Each dr In dtClientes.Rows

                    'recorremos las vias del cliente y si no esta seleccionada se selecciona
                    Dim oBP As SAPbobsCOM.BusinessPartners = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                    oBP.GetByKey(dr("CardCode").ToString)

                    Dim bEncontrado = False
                    Dim bModificado As Boolean = False
                    Dim bEsPrimero As Boolean = False
                    For Each dr2 In dtVias.Rows
                        bEncontrado = False
                        For i = 0 To oBP.BPPaymentMethods.Count - 1
                            oBP.BPPaymentMethods.SetCurrentLine(i)

                            If oBP.BPPaymentMethods.PaymentMethodCode = "" Then
                                bEsPrimero = True
                            End If

                            If oBP.BPPaymentMethods.PaymentMethodCode = dr2("PayMethCod").ToString Then
                                bEncontrado = True
                                Exit For
                            End If
                        Next

                        If bEncontrado = False Then
                            bModificado = True
                            If bEsPrimero = True Then
                                bEsPrimero = False
                            Else
                                oBP.BPPaymentMethods.Add()
                            End If

                            oBP.BPPaymentMethods.PaymentMethodCode = dr2("PayMethCod").ToString
                        End If

                    Next

                    If bModificado Then
                        If oBP.Update() <> 0 Then
                            oLog.escribeMensaje(dr("CardCode").ToString + " - Error:" + oCompany.GetLastErrorDescription)
                        Else
                            oLog.escribeMensaje(dr("CardCode").ToString + " - Ok:" + oCompany.GetLastErrorDescription)
                        End If
                    End If

                Next
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            oLog.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            oLog.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        End Try


    End Sub
#End Region

#Region "Actualizar campos"
    Public Shared Sub Actualizar_Campos(ByRef oLog As EXO_Log.EXO_Log)
        Dim oDBSAP As HanaConnection = Nothing
        Dim oCompany As SAPbobsCOM.Company = Nothing
        Dim sError As String = ""
        Dim sSQL As String = ""
        Dim OdtDatos As System.Data.DataTable = Nothing
        Dim sPass As String = ""
        Dim refDI As EXO_DIAPI.EXO_DIAPI = Nothing
        Dim oXML As String = ""
        Dim sDir As String = Application.StartupPath
        Try
            sPass = Conexiones.Datos_Confi("DI", "Password")
            Dim tipoServidor As SAPbobsCOM.BoDataServerTypes = SAPbobsCOM.BoDataServerTypes.dst_HANADB
            Dim tipocliente As EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente = EXO_DIAPI.EXO_DIAPI.EXO_TipoCliente.Clasico
            Conexiones.Connect_SQLHANA(oDBSAP, "HANA", oLog)
            sSQL = " SELECT ""EXO_BD"" FROM ""SOL_AUTORIZ"".""EXO_SOCIEDADES"" "
            OdtDatos = New System.Data.DataTable
            Conexiones.FillDtDB(oDBSAP, OdtDatos, sSQL)
            If OdtDatos.Rows.Count > 0 Then
                oLog.escribeMensaje("Se va a proceder a recorrer las SOCIEDADES...", EXO_Log.EXO_Log.Tipo.advertencia)
                For Each dr In OdtDatos.Rows
                    Conexiones.Connect_Company(oCompany, "DI", dr("EXO_BD").ToString, oLog)
#Region "Creamos campos"
                    Try
                        refDI = New EXO_DIAPI.EXO_DIAPI(oCompany, oLog)
                    Catch ex As Exception
                        refDI = New EXO_DIAPI.EXO_DIAPI(tipoServidor, oCompany.Server.ToString, oCompany.LicenseServer.ToString, oCompany.CompanyDB.ToString, oCompany.UserName.ToString, sPass, tipocliente)
                    End Try

                    Dim fsXML As New FileStream(sDir & "\XML_BD\UDFs_EXO_DOCs.xml", FileMode.Open, FileAccess.Read)
                    Dim xmldoc As New XmlDocument()
                    xmldoc.Load(fsXML)
                    oXML = xmldoc.InnerXml.ToString
                    oLog.escribeMensaje(oXML)
                    refDI.comunes.LoadBDFromXML(oXML, sError)
                    oLog.escribeMensaje("Validado: UDFs_EXO_DOCs - " & sError, EXO_Log.EXO_Log.Tipo.advertencia)

                    fsXML = New FileStream(sDir & "\XML_BD\UDFs_EXO_OVPM.xml", FileMode.Open, FileAccess.Read)
                    xmldoc.Load(fsXML)
                    oXML = xmldoc.InnerXml.ToString
                    refDI.comunes.LoadBDFromXML(oXML, sError)
                    oLog.escribeMensaje("Validado: UDFs_EXO_OVPM - " & sError, EXO_Log.EXO_Log.Tipo.advertencia)
#End Region
                    Conexiones.Disconnect_Company(oCompany)
                Next
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            sError = exCOM.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            sError = ex.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally
            Conexiones.Disconnect_SQLHANA(oDBSAP)
            refDI = Nothing
        End Try

    End Sub
#End Region

#Region "Enviar Mails Autorizaciones pendientes"
    Public Shared Sub EnviarMails(ByRef oLog As EXO_Log.EXO_Log)
#Region "Variables"
        Dim oDBSAP As HanaConnection = Nothing
        Dim sError As String = ""
        Dim sSQL As String = "" : Dim sSQLAUTPDTE As String = ""
        Dim odtDatos As System.Data.DataTable = Nothing : Dim odtUsuarios As System.Data.DataTable = Nothing
        Dim ESprimero As Boolean = True
        Dim sBBDD As String = "" : Dim sUsuario As String = ""
        Dim sSQLAct As String = ""
        Dim odtDatos2 As System.Data.DataTable = Nothing
#End Region
        Try
            Conexiones.Connect_SQLHANA(oDBSAP, "HANA", oLog)
            'campo val_pag
            sSQL = " SELECT ""EXO_USUARIO"",""EXO_VALPAG"" FROM ""SOL_AUTORIZ"".""EXO_USUARIOS"" " 'tambien puedes sacar el campo exo_valpag
            odtUsuarios = New System.Data.DataTable
            Conexiones.FillDtDB(oDBSAP, odtUsuarios, sSQL)
            If odtUsuarios.Rows.Count > 0 Then
                For Each ur As DataRow In odtUsuarios.Rows
                    sUsuario = ur.Item("EXO_USUARIO").ToString
                    sSQL = " SELECT ""EXO_BD"" FROM ""SOL_AUTORIZ"".""EXO_SOCIEDADES"" "
                    odtDatos = New System.Data.DataTable
                    Conexiones.FillDtDB(oDBSAP, odtDatos, sSQL)
                    If odtDatos.Rows.Count > 0 Then
                        oLog.escribeMensaje("Se va a proceder a recorrer las SOCIEDADES...", EXO_Log.EXO_Log.Tipo.advertencia)
                        odtDatos2 = New System.Data.DataTable

                        For Each dr As DataRow In odtDatos.Rows
                            sSQLAUTPDTE = "SELECT * FROM (" ': ESprimero = True
#Region "Creamos SQL para recorrer y envial mail"
                            'If ESprimero = False Then
                            '    sSQLAUTPDTE &= " UNION ALL "
                            'Else
                            '    ESprimero = False
                            'End If

                            sBBDD = dr.Item("EXO_BD").ToString
                            sSQLAUTPDTE &= " ("
                            sSQLAUTPDTE &= " Select '" + sBBDD + "' ""BD"", T11.""CompnyName"" ""BDName"",T0.""WddCode"",t0.""ProcesStat"" ""Status"", T8.""CardName""," +
                                " CASE WHEN T0.""IsDraft""='Y' THEN T0.""DraftEntry"" ELSE T7.""DocEntry"" END ""DocInterno"", " +
                                " CAST(T7.""NumAtCard"" as VARCHAR) ""DocNum"",T0.""ObjType"",  " +
                                " T0.""CreateDate"",T0.""CreateTime"", " +
                                " case  when  T7.""DocTotalFC""<> 0 then CAST((T7.""DocTotalFC"" - T7.""TotalExpFC"" - T7.""VatSumFC"") as decimal(10,2)) " +
                                    " else     CAST((T7.""DocTotal"" - T7.""TotalExpns"" - T7.""VatSum"") as decimal(10,2)) " +
                                    " end ""DocTotal""," +
                                " t6.""Name"" ""Departamento"",T0.""Remarks"",T0.""MaxReqr"", " +
                                " T10.""USER_CODE"" ""Aprobador"",T9.""Remarks"" ""ComAprobador"",T9.""Status"" ""StatusAprob"", " +
                                " T0.""IsDraft"" ""borrador"",T7.""DocCur"",T7.""DocRate"",T7.""Project"" ""Proyecto"" " +
                                 " from """ + sBBDD + """.""OWDD"" T0  " +
                                "   inner Join """ + sBBDD + """.""WDD1"" T9 On T0.""WddCode""=T9.""WddCode""  And T0.""CurrStep""=T9.""StepCode"" " +
                                " inner Join """ + sBBDD + """.""OUSR"" T1 on T0.""OwnerID""=T1.""USERID""  " +
                                " inner Join """ + sBBDD + """.""WTM2"" T2 ON T0.""WtmCode""=T2.""WtmCode"" And T9.""SortId""=T2.""SortId""  " +
                                " inner Join """ + sBBDD + """.""WST1"" T3 ON T2.""WstCode""=T3.""WstCode"" " +
                                " inner Join """ + sBBDD + """.""OWST"" T4 ON T2.""WstCode""=T4.""WstCode"" " +
                                " inner Join """ + sBBDD + """.""OUSR"" T5 ON T3.""UserID""=T5.""USERID"" " +
                                " Left Join """ + sBBDD + """.""OUDP"" T6 ON T6.""Code""=t1.""Department""  " +
                                "  inner join """ + sBBDD + """.""ODRF"" T7 ON T0.""DraftEntry""=T7.""DocEntry"" " +
                                " inner Join """ + sBBDD + """.""OCRD"" T8 ON T7.""CardCode""=T8.""CardCode"" " +
                                " Left join """ + sBBDD + """.""OUSR"" T10 On T10.""USERID""=T9.""UserID"" " +
                                " ,   """ + sBBDD + """.""OADM"" T11 " +
                                " where T10.""USER_CODE"" ='" + sUsuario + "' AND T0.""ObjType"" in (18,22) AND T0.""ProcesStat"" ='W' and t9.""U_EXO_AMailE""='N'"

                            'sSQLAUTPDTE &= " UNION ALL "
                            'sSQLAUTPDTE &= " Select '" + sBBDD + "' ""BD"", T11.""CompnyName"" ""BDName"",T0.""WddCode"",t0.""ProcesStat"" ""Status"", T8.""CardName""," +
                            '    " CASE WHEN T0.""IsDraft""='Y' THEN T0.""DraftEntry"" ELSE T7.""DocEntry"" END ""DocInterno"", " +
                            '    " CAST(T7.""DocNum"" as VARCHAR) ""DocNum"",T0.""ObjType"", " +
                            '    " T0.""CreateDate"",T0.""CreateTime"", " +
                            '    " case when T7.""DocTotalFC""<> 0 then CAST((T7.""DocTotalFC"" - T7.""VatSumFC"") as decimal(10,2)) else CAST((T7.""DocTotal"" -  T7.""VatSum"") as decimal(10,2)) end ""DocTotal"", " +
                            '    " t6.""Name"" ""Departamento"",T0.""Remarks"",T0.""MaxReqr"", " +
                            '    " T10.""USER_CODE"" ""Aprobador"",T9.""Remarks"" ""ComAprobador"",T9.""Status"" ""StatusAprob"", " +
                            '    " T0.""IsDraft"" ""borrador"",T7.""DocCurr"" ""DocCur"",T7.""DocRate"",'' ""Proyecto"" " +
                            '     " from """ + sBBDD + """.""OWDD"" T0  " +
                            '    "   inner Join """ + sBBDD + """.""WDD1"" T9 On T0.""WddCode""=T9.""WddCode""  And T0.""CurrStep""=T9.""StepCode"" " +
                            '    " inner Join """ + sBBDD + """.""OUSR"" T1 on T0.""OwnerID""=T1.""USERID""  " +
                            '    " inner Join """ + sBBDD + """.""WTM2"" T2 ON T0.""WtmCode""=T2.""WtmCode"" And T9.""SortId""=T2.""SortId""  " +
                            '    " inner Join """ + sBBDD + """.""WST1"" T3 ON T2.""WstCode""=T3.""WstCode"" " +
                            '    " inner Join """ + sBBDD + """.""OWST"" T4 ON T2.""WstCode""=T4.""WstCode"" " +
                            '    " inner Join """ + sBBDD + """.""OUSR"" T5 ON T3.""UserID""=T5.""USERID"" " +
                            '    " Left Join """ + sBBDD + """.""OUDP"" T6 ON T6.""Code""=t1.""Department""  " +
                            '    "  inner join """ + sBBDD + """.""OPDF"" T7 ON T0.""DraftEntry""=T7.""DocEntry"" " +
                            '    " inner Join """ + sBBDD + """.""OCRD"" T8 ON T7.""CardCode""=T8.""CardCode"" " +
                            '    " Left join """ + sBBDD + """.""OUSR"" T10 On T10.""USERID""=T9.""UserID"" " +
                            '    " ,   """ + sBBDD + """.""OADM"" T11 " +
                            '    " where T10.""USER_CODE"" ='" + sUsuario + "' AND T0.""ObjType"" in (46) AND T0.""ProcesStat"" ='W' and t9.""U_EXO_AMailE""='N'"
                            sSQLAUTPDTE &= ") "
                            'añadir las autorizaciones de pago por desarrollo
                            'que valpag=1 y enviadomail1 =n y validador1=n
                            'que valpag=2 y enviadomail2=n, validador1=Y, validador2=n y (status=O o que sum(doctotal)>3000)

                            'If ur.Item("EXO_VALPAG").ToString = "1" Then
                            '    sSQLAUTPDTE &= " UNION ALL "
                            '    sSQLAUTPDTE &= "( "
                            '    sSQLAUTPDTE &= " (Select '" + sBBDD + "' ""BD"", T11.""CompnyName"" ""BDName"",T0.""DocEntry"" ""WddCode"",'W' ""Status"",T1.""CardName"", " +
                            '        " T0.""DocEntry""  ""DocInterno"",  CAST(T0.""DocNum"" as VARCHAR) ""DocNum"",'46' ""ObjType"" " +
                            '        " , T0.""CreateDate"",T0.""CreateTime"",cast(sum(T2.""U_EXO_DOCTOTAL"") as decimal(10,2)) ""DocTotal"",  '' ""Departamento"",'' ""Remarks"",1 ""MaxReqr"", " +
                            '        " 'Eruiz' ""Aprobador"", T0.""U_EXO_COM1"" ""comAprobador"",CASE WHEN T0.""U_EXO_VAL1""='N' THEN 'W' ELSE T0.""U_EXO_VAL1"" END ""StatusAprobador"",  " +
                            '        " 'N' ""borrador"", Max(t3.""DocCur"")  ""DocCur"", max(T3.""DocRate"") ""DocRate"", '' ""Proyecto"" " +
                            '        " FROM """ + sBBDD + """.""@EXO_VALPAG"" T0 inner join """ + sBBDD + """.""OCRD"" T1 On T0.""U_EXO_CODPRO""=T1.""CardCode"" " +
                            '        " inner Join """ + sBBDD + """.""@EXO_VALPAG1"" T2 On T0.""DocEntry""=T2.""DocEntry"" " +
                            '        " INNER JOIN """ + sBBDD + """.""OPCH"" T3 On T2.""U_EXO_DOCE""=T3.""DocEntry"" " +
                            '        " ,""" + sBBDD + """.""OADM"" T11   " +
                            '        " WHERE  ""U_EXO_VAL1""='N' AND ""U_EXO_VAL2""='N'   and T0.""Status""='O'" +
                            '        " Group BY T11.""CompnyName"",T0.""DocEntry"",T1.""CardName"",T0.""DocNum"", T0.""CreateDate"",T0.""CreateTime"",T0.""U_EXO_COM1"",T0.""U_EXO_VAL1"", T0.""UpdateDate"",T0.""UpdateTime"") "
                            '    sSQLAUTPDTE &= ") "
                            'End If

                            'If ur.Item("EXO_VALPAG").ToString = "2" Then
                            '    sSQLAUTPDTE &= " UNION ALL "
                            '    sSQLAUTPDTE &= "( "
                            '    sSQLAUTPDTE &= "   (Select '" + sBBDD + "' ""BD"", T11.""CompnyName"" ""BDName"",T0.""DocEntry"" ""WddCode"",'W' ""Status"",T1.""CardName"", " +
                            '    " T0.""DocEntry""  ""DocInterno"",  CAST(T0.""DocNum"" as VARCHAR) ""DocNum"",'46' ""ObjType"" " +
                            '    " , T0.""CreateDate"",T0.""CreateTime"",cast(sum(T2.""U_EXO_DOCTOTAL"") as decimal(10,2)) ""DocTotal"",  '' ""Departamento"",'' ""Remarks"",1 ""MaxReqr"", " +
                            '    " 'Eruiz' ""Aprobador"", T0.""U_EXO_COM1"" ""comAprobador"",CASE WHEN T0.""U_EXO_VAL1""='N' THEN 'W' ELSE T0.""U_EXO_VAL1"" END ""StatusAprobador"",  " +
                            '    " 'N' ""borrador"", Max(t3.""DocCur"")  ""DocCur"", max(T3.""DocRate"") ""DocRate"", '' ""Proyecto"" " +
                            '    " FROM """ + sBBDD + """.""@EXO_VALPAG"" T0 inner join """ + sBBDD + """.""OCRD"" T1 On T0.""U_EXO_CODPRO""=T1.""CardCode"" " +
                            '    " inner Join """ + sBBDD + """.""@EXO_VALPAG1"" T2 On T0.""DocEntry""=T2.""DocEntry"" " +
                            '    " INNER JOIN """ + sBBDD + """.""OPCH"" T3 On T2.""U_EXO_DOCE""=T3.""DocEntry"" " +
                            '    " ,""" + sBBDD + """.""OADM"" T11   " +
                            '    " WHERE  ""U_EXO_VAL1""='Y' AND ""U_EXO_VAL2""='N'   and T0.""Status""='O'" +
                            '    " Group BY T11.""CompnyName"",T0.""DocEntry"",T1.""CardName"",T0.""DocNum"", T0.""CreateDate"",T0.""CreateTime"",T0.""U_EXO_COM1"",T0.""U_EXO_VAL1"", T0.""UpdateDate"",T0.""UpdateTime"" " +
                            '    " having sum(T2.""U_EXO_DOCTOTAL"")>3000 )"
                            '    sSQLAUTPDTE &= ") "
                            'End If
#End Region
                            sSQLAUTPDTE &= ") T "
                            sSQLAUTPDTE &= " group by  ""BD"",  ""BDName"",""WddCode"", ""Status"", ""CardName"", "
                            sSQLAUTPDTE &= " ""DocInterno"",  ""DocNum"",""ObjType"",  ""CreateDate"",""CreateTime"", ""DocTotal"", "
                            sSQLAUTPDTE &= " ""Departamento"",""Remarks"",""MaxReqr"",  ""Aprobador"",""ComAprobador"",""StatusAprob"",   ""borrador"",""DocCur"",""DocRate"",""Proyecto""    "
                            sSQLAUTPDTE &= " ORDER BY   t.""ObjType"", t.""BD"", t.""DocNum"" "

                            Conexiones.FillDtDB(oDBSAP, odtDatos2, sSQLAUTPDTE)
                        Next
                        sSQLAct = ""
                        Procesos.EnviarMails_Aut_Pdtes(oLog, odtDatos2, sUsuario, sSQLAct, oDBSAP)
#Region "Actualiza documentos"
                        'Actualizamos documentos
                        If sSQLAct <> "" Then
                            Dim sSQLArray() As String = Split(sSQLAct, ";")
                            oLog.escribeMensaje("Se va a proceder a actualizar los documentos..." & sSQLArray.Length - 1, EXO_Log.EXO_Log.Tipo.advertencia)
                            For i As Integer = 0 To (sSQLArray.Length - 1)
                                If sSQLArray(i) <> "" Then
                                    Try
                                        Conexiones.ExecuteSqlDB(oDBSAP, sSQLArray(i))
                                        oLog.escribeMensaje("OK - " & sSQLArray(i), EXO_Log.EXO_Log.Tipo.informacion)
                                    Catch ex As Exception
                                        oLog.escribeMensaje("ERROR - " & sSQLArray(i), EXO_Log.EXO_Log.Tipo.error)
                                    End Try
                                End If
                            Next
                        End If
#End Region
                    Else
                        oLog.escribeMensaje("No existen sociedades definidas", EXO_Log.EXO_Log.Tipo.error)
                        Exit Sub
                    End If
                Next

            Else
                oLog.escribeMensaje("No existen usuarios definidos", EXO_Log.EXO_Log.Tipo.error)
                Exit Sub
            End If



        Catch exCOM As System.Runtime.InteropServices.COMException
            sError = exCOM.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            sError = ex.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally
            Conexiones.Disconnect_SQLHANA(oDBSAP)
            odtDatos = Nothing : odtUsuarios = Nothing
        End Try
    End Sub
    Public Shared Sub EnviarMails_Aut_Pdtes(ByRef oLog As EXO_Log.EXO_Log, ByRef odtDatos As System.Data.DataTable, ByVal sUsuario As String, ByRef sSQLAct As String, ByRef oDBSAP As HanaConnection)
#Region "Variables"
        Dim sError As String = ""
        Dim sEmpresa As String = ""
        Dim sHora1 As String = "" : Dim SHora2 As String = "" : Dim EsUrgente As Boolean = False
        Dim sMailFROM As String = "" : Dim sDirmail As String = "" : Dim cuerpo As String = ""
        Dim sSMTP As String = "" : Dim sPuerto As String = "" : Dim sUsMail As String = "" : Dim sPSSMail As String = ""
        Dim sDoc As String = "" : Dim sTipo As String = "" : Dim sBBDD As String = "" : Dim sUrgente As String = "" : Dim stabla As String = ""
        Dim bExAut As Boolean = False
        Dim sCambiaBBDD As String = "" : Dim sCambiaTDoc As String = ""
        Dim sTextoEnviado As String = ""
#End Region
        Try

            If odtDatos.Rows.Count > 0 Then
                oLog.escribeMensaje("Existen autorizaciones pdtes de enviar del usuario " & sUsuario & ". ", EXO_Log.EXO_Log.Tipo.advertencia)
                sHora1 = Conexiones.Datos_Confi("NO_URGE", "HORA1") : SHora2 = Conexiones.Datos_Confi("NO_URGE", "HORA2")
                Dim shora As String = Now.Hour.ToString("00") & ":" & Now.Minute.ToString("00")
                sMailFROM = Conexiones.Datos_Confi("MAIL", "MAILFROM")
                sSMTP = Conexiones.Datos_Confi("MAIL", "SMTP") : sPuerto = Conexiones.Datos_Confi("MAIL", "PUERTO")
                sUsMail = Conexiones.Datos_Confi("MAIL", "USUARIO") : sPSSMail = Conexiones.Datos_Confi("MAIL", "PASS")
                sEmpresa = odtDatos.Rows(0).Item("BD").ToString
                oLog.escribeMensaje("Empresa: " & sEmpresa, EXO_Log.EXO_Log.Tipo.advertencia)
                Dim correo As New System.Net.Mail.MailMessage()
                'Comprobamos si estamos en tiempo de urgencia
                oLog.escribeMensaje("HORA1: " & shora & "=" & sHora1 & "--- HORA 2:" & shora & "=" & SHora2, EXO_Log.EXO_Log.Tipo.advertencia)
                If shora = sHora1 Or shora = SHora2 Then
                    EsUrgente = False
                    correo.Priority = System.Net.Mail.MailPriority.Normal 'Prioridad
                    oLog.escribeMensaje("Hora No Urgente...", EXO_Log.EXO_Log.Tipo.advertencia)
                    correo.Subject = "Autorizaciones Pendientes "
                    cuerpo = "Estas son las autorizaciones que tiene pdtes.:" & ChrW(10) ' & ChrW(13)
                Else
                    EsUrgente = True
                    oLog.escribeMensaje("Hora Urgente...", EXO_Log.EXO_Log.Tipo.advertencia)
                    correo.Priority = System.Net.Mail.MailPriority.High 'Prioridad
                    correo.Subject = "Autorizaciones Pendientes Urgentes"
                    cuerpo = "Estas son las autorizaciones urgentes que tiene pdtes.:" & ChrW(10) & ChrW(13)
                End If
                correo.From = New System.Net.Mail.MailAddress(sMailFROM, "Envío Aut. Autorizaciones Pdtes.")
                correo.To.Clear()
                sDirmail = Conexiones.GetValueDB(oDBSAP, sEmpresa & ".""OUSR""", """E_Mail""", """USER_CODE""='" & sUsuario & "' ")
                If sDirmail <> "" Then
                    correo.To.Add(sDirmail)

                    Dim strHeader As String = "<table><tbody>"
                    Dim strFooter As String = "</tbody></table>"
                    Dim sbContent As New StringBuilder()
                    sTextoEnviado = ""
                    sbContent.Append(String.Format("<td>{0}</td>", "https://autorizaciones.solariaenergia.com/"))
                    For Each dr As DataRow In odtDatos.Rows
#Region "Asignamos Varible sTipo y sTabla"
                        If dr("Borrador").ToString = "Y" Then
                            Select Case dr("ObjType").ToString
                                Case "13" : sTipo = "Factura de cliente" : stabla = "ODRF"
                                Case "18" : sTipo = "Factura de proveedor" : stabla = "ODRF"
                                Case "14" : sTipo = "Abono de cliente" : stabla = "ODRF"
                                Case "19" : sTipo = "Abono de proveedor" : stabla = "ODRF"
                                Case "46" : sTipo = "Pago Factura" : stabla = "OPDF"
                                Case "22" : sTipo = "Pedido de compra" : stabla = "ODRF"
                                Case "17" : sTipo = "Pedido de Venta" : stabla = "ODRF"
                                Case Else : sTipo = dr("ObjType").ToString : stabla = "ODRF"
                            End Select
                        Else
                            Select Case dr("ObjType").ToString
                                Case "13" : sTipo = "Factura de cliente" : stabla = "OINV"
                                Case "18" : sTipo = "Factura de proveedor" : stabla = "OPCH"
                                Case "14" : sTipo = "Abono de cliente" : stabla = "ORIN"
                                Case "19" : sTipo = "Abono de proveedor" : stabla = "ORPC"
                                Case "46" : sTipo = "Pago Factura" : stabla = "OVPM"
                                Case "22" : sTipo = "Pedido de compra" : stabla = "OPOR"
                                Case "17" : sTipo = "Pedido de Venta" : stabla = "ORDR"
                                Case Else : sTipo = dr("ObjType").ToString : stabla = ""
                            End Select
                        End If
#End Region
                        Dim sEnviado As String = Conexiones.GetValueDB(oDBSAP, dr.Item("BD").ToString & ".""" & stabla & """", """U_EXO_AMailE""", """DocEntry""='" & dr.Item("DocInterno").ToString & "' ")
                        If sEnviado <> "Y" Then
                            If stabla <> "" And sTipo <> "" Then
                                sUrgente = Conexiones.GetValueDB(oDBSAP, sEmpresa & ".""" & stabla & """", """U_EXO_AUrg""", """DocEntry""='" & dr.Item("DocInterno").ToString & "' ")
                            End If
                            If sUrgente = "" Then
                                sUrgente = "N"
                            End If
                            oLog.escribeMensaje("Tabla: " & stabla & " - DocEntry: " & dr.Item("DocInterno").ToString & " --- Urgente: " & sUrgente, EXO_Log.EXO_Log.Tipo.advertencia)
                            If EsUrgente = True Then
                                If sUrgente = "Y" Then
                                    If sCambiaTDoc <> sTipo Then
                                        sCambiaTDoc = sTipo
#Region "Asignamos cabecera de la tabla Tipo Documento"
                                        sbContent.Append("<tr>") : sbContent.Append(String.Format("<td>{0}</td>", "")) : sbContent.Append("</tr>")
                                        sbContent.Append("<tr>") : sbContent.Append(String.Format("<td>{0}</td>", "T. Documento:" & sTipo)) : sbContent.Append("</tr>")
                                        sbContent.Append("<tr>")
                                        sbContent.Append(String.Format("<td>{0}</td>", "Proveedor"))
                                        sbContent.Append(String.Format("<td><p align=""right"">{0}</p></td>", "Referencia"))
                                        sbContent.Append(String.Format("<td><p align=""right"">{0}</p></td>", "Proyecto"))
                                        sbContent.Append(String.Format("<td><p align=""right"">{0}</p></td>", "Importe"))
                                        sbContent.Append(String.Format("<td>{0}</td>", "Comentario"))
                                        sbContent.Append("</tr>")
                                        sbContent.Append("<tr>")
                                        sbContent.Append(String.Format("<td>{0}</td>", "________________________________________"))
                                        sbContent.Append(String.Format("<td>{0}</td>", "______________"))
                                        sbContent.Append(String.Format("<td>{0}</td>", "______________"))
                                        sbContent.Append(String.Format("<td>{0}</td>", "______________"))
                                        sbContent.Append(String.Format("<td>{0}</td>", "________________________________________"))
                                        sbContent.Append("</tr>")

                                        sTextoEnviado &= ChrW(10) '& ChrW(13)
                                        sTextoEnviado &= "Proveedor" & ChrW(9) & "Referencia" & ChrW(9) & "Proyecto" & ChrW(9) & "Importe" & ChrW(9) & "Comentario" & Chr(10) ' & ChrW(13)
                                        sTextoEnviado &= "#####################################################################################" & Chr(10) '& ChrW(13)
#End Region
                                    End If
                                    sbContent.Append("<tr>")
                                    sbContent.Append(String.Format("<td>{0}</td>", dr.Item("CardName").ToString))
                                    sbContent.Append(String.Format("<td><p align=""right"">{0}</p></td>", dr.Item("DocNum").ToString))
                                    sbContent.Append(String.Format("<td><p align=""right"">{0}</p></td>", dr.Item("Proyecto").ToString))
                                    sbContent.Append(String.Format("<td><p align=""right"">{0}</p></td>", CDbl(dr.Item("DocTotal").ToString).ToString("#,##0.00") & dr.Item("DocCur").ToString))
                                    sbContent.Append(String.Format("<td>{0}</td>", dr.Item("Remarks").ToString))
                                    sbContent.Append("</tr>")

                                    sTextoEnviado &= dr.Item("CardName").ToString & ChrW(9) & dr.Item("DocNum").ToString & ChrW(9) & dr.Item("Proyecto").ToString & ChrW(9) & CDbl(dr.Item("DocTotal").ToString).ToString("#,##0.00") & dr.Item("DocCur").ToString & ChrW(9) & dr.Item("Remarks").ToString & Chr(10) '& ChrW(13)

                                    sSQLAct &= "update T9 Set T9.""U_EXO_AMailE""='Y' "
                                    sSQLAct &= " from """ + dr.Item("BD").ToString + """.""WDD1"" T9  inner join """ + dr.Item("BD").ToString + """.""OUSR"" T10 On T10.""USERID""=T9.""UserID"" "
                                    sSQLAct &= " WHERE t10.""USER_CODE""='" + sUsuario + "' and t9.""WddCode""='" + dr.Item("WddCode").ToString + "'; "


                                    bExAut = True
                                End If
                            Else
                                If sCambiaTDoc <> sTipo Then
                                    sCambiaTDoc = sTipo
#Region "Asignamos cabecera de la tabla Tipo Documento"
                                    sbContent.Append("<tr>") : sbContent.Append(String.Format("<td>{0}</td>", "")) : sbContent.Append("</tr>")
                                    sbContent.Append("<tr>") : sbContent.Append(String.Format("<td>{0}</td>", "T. Documento:" & sTipo)) : sbContent.Append("</tr>")
                                    sbContent.Append("<tr>")

                                    sbContent.Append("<tr>") : sbContent.Append(String.Format("<td>{0}</td>", "")) : sbContent.Append("</tr>")
                                    sbContent.Append("<tr>") : sbContent.Append(String.Format("<td>{0}</td>", "T. Documento:" & sTipo)) : sbContent.Append("</tr>")
                                    sbContent.Append("<tr>")
                                    sbContent.Append(String.Format("<td><p align=""center"">{0}</p></td>", "Urgente"))
                                    sbContent.Append(String.Format("<td>{0}</td>", "Proveedor"))
                                    sbContent.Append(String.Format("<td><p align=""right"">{0}</p></td>", "Referencia"))
                                    sbContent.Append(String.Format("<td><p align=""right"">{0}</p></td>", "Proyecto"))
                                    sbContent.Append(String.Format("<td><p align=""right"">{0}</p></td>", "Importe"))
                                    sbContent.Append(String.Format("<td>{0}</td>", "Comentario"))
                                    sbContent.Append("</tr>")
                                    sbContent.Append("<tr>")
                                    sbContent.Append(String.Format("<td>{0}</td>", "__________"))
                                    sbContent.Append(String.Format("<td>{0}</td>", "________________________________________"))
                                    sbContent.Append(String.Format("<td>{0}</td>", "______________"))
                                    sbContent.Append(String.Format("<td>{0}</td>", "______________"))
                                    sbContent.Append(String.Format("<td>{0}</td>", "______________"))
                                    sbContent.Append(String.Format("<td>{0}</td>", "________________________________________"))
                                    sbContent.Append("</tr>")

                                    sTextoEnviado &= ChrW(10) '& ChrW(13)
                                    sTextoEnviado &= "Urgente" & ChrW(9) & "Proveedor" & ChrW(9) & "Referencia" & ChrW(9) & "Proyecto" & ChrW(9) & "Importe" & ChrW(9) & "Comentario" & Chr(10) ' & ChrW(13)
                                    sTextoEnviado &= "################################################################################################################################" & Chr(10) '& ChrW(13)                                  
#End Region
                                End If
                                sbContent.Append("<tr>")
                                sbContent.Append(String.Format("<td><p align=""center"">{0}</p></td>", sUrgente))
                                sbContent.Append(String.Format("<td>{0}</td>", dr.Item("CardName").ToString))
                                sbContent.Append(String.Format("<td><p align=""right"">{0}</p></td>", dr.Item("DocNum").ToString))
                                sbContent.Append(String.Format("<td><p align=""right"">{0}</p></td>", dr.Item("Proyecto").ToString))
                                sbContent.Append(String.Format("<td><p align=""right"">{0}</p></td>", CDbl(dr.Item("DocTotal").ToString).ToString("#,##0.00") & dr.Item("DocCur").ToString))
                                sbContent.Append(String.Format("<td>{0}</td>", dr.Item("Remarks").ToString))
                                sbContent.Append("</tr>")

                                sTextoEnviado &= sUrgente & ChrW(9) & dr.Item("CardName").ToString & ChrW(9) & dr.Item("DocNum").ToString & ChrW(9) & dr.Item("Proyecto").ToString & ChrW(9) & CDbl(dr.Item("DocTotal").ToString).ToString("#,##0.00") & dr.Item("DocCur").ToString & ChrW(9) & dr.Item("Remarks").ToString & Chr(10) '& ChrW(13)

                                sSQLAct &= "update T9 Set T9.""U_EXO_AMailE""='Y' "
                                sSQLAct &= " from """ + dr.Item("BD").ToString + """.""WDD1"" T9  inner join """ + dr.Item("BD").ToString + """.""OUSR"" T10 On T10.""USERID""=T9.""UserID"" "
                                sSQLAct &= " WHERE t10.""USER_CODE""='" + sUsuario + "' and t9.""WddCode""='" + dr.Item("WddCode").ToString + "'; "
                                bExAut = True
                            End If
                        End If
                    Next
                    sbContent.Append(String.Format("<td>{0}</td>", ""))
                    sbContent.Append(String.Format("<td>{0}</td>", ""))
                    Dim emailTemplate As String = strHeader & sbContent.ToString() & strFooter
                    cuerpo &= strHeader & sbContent.ToString() & strFooter
                    correo.Body = cuerpo
                    correo.IsBodyHtml = True
                    correo.Priority = System.Net.Mail.MailPriority.Normal

                    Dim smtp As New System.Net.Mail.SmtpClient
                    smtp.Host = sSMTP
                    smtp.Port = sPuerto
                    smtp.UseDefaultCredentials = True
                    smtp.Credentials = New System.Net.NetworkCredential(sUsMail, sPSSMail)
                    smtp.EnableSsl = True
                    Try
                        'oLog.escribeMensaje(sTextoEnviado, EXO_Log.EXO_Log.Tipo.informacion)
                        ' Si existen autorizaciones se envía
                        If bExAut = True Then
                            'If sDirmail = "mperiz@expertone.es" Then
                            smtp.Send(correo)
                            oLog.escribeMensaje(sEmpresa & "- Correo enviado: " & sDirmail, EXO_Log.EXO_Log.Tipo.informacion)
                            oLog.escribeMensaje(sTextoEnviado, EXO_Log.EXO_Log.Tipo.informacion)
                            'End If
                        Else
                            oLog.escribeMensaje(sEmpresa & "- Correo No enviado: " & sDirmail, EXO_Log.EXO_Log.Tipo.advertencia)
                        End If

                        correo.Dispose()
                    Catch ex As Exception
                        sError = ex.Message
                        oLog.escribeMensaje(sEmpresa & " - Error enviando correo: " & sDirmail & ". " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try
                Else
                    oLog.escribeMensaje("El Usuario " & sUsuario & " no tiene asignado mail para envío de información", EXO_Log.EXO_Log.Tipo.error)
                    Exit Sub
                End If
            Else
                oLog.escribeMensaje("No existen autorizaciones pdtes de enviar del usuario " & sUsuario & ". ", EXO_Log.EXO_Log.Tipo.advertencia)
                Exit Sub
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            sError = exCOM.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            sError = ex.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally
            oLog.escribeMensaje("SQL : " & sSQLAct, EXO_Log.EXO_Log.Tipo.advertencia)
        End Try
    End Sub
#End Region

#Region "enviar mails status aprobacion"

    Friend Shared Sub EnviarMails_statusaprob(oLog As EXO_Log.EXO_Log)

#Region "Variables"
        Dim oDBSAP As HanaConnection = Nothing
        Dim sError As String = ""
        Dim sSQL As String = "" : Dim sSQLAUTPDTE As String = ""
        Dim odtDatos As System.Data.DataTable = Nothing : Dim odtUsuarios As System.Data.DataTable = Nothing
        Dim ESprimero As Boolean = True
        Dim sBBDD As String = "" : Dim sUsuario As String = ""
        Dim sSQLAct As String = ""
        Dim odtDatos2 As System.Data.DataTable = Nothing
#End Region
        Try
            Conexiones.Connect_SQLHANA(oDBSAP, "HANA", oLog)
            'campo val_pag
            sSQL = " SELECT ""EXO_USUARIO"" FROM ""SOL_AUTORIZ"".""EXO_CREADORES"" " 'tambien puedes sacar el campo exo_valpag
            odtUsuarios = New System.Data.DataTable
            Conexiones.FillDtDB(oDBSAP, odtUsuarios, sSQL)
            If odtUsuarios.Rows.Count > 0 Then
                For Each ur As DataRow In odtUsuarios.Rows
                    sUsuario = ur.Item("EXO_USUARIO").ToString
                    sSQL = " SELECT ""EXO_BD"" FROM ""SOL_AUTORIZ"".""EXO_SOCIEDADES"" "
                    odtDatos = New System.Data.DataTable
                    Conexiones.FillDtDB(oDBSAP, odtDatos, sSQL)
                    If odtDatos.Rows.Count > 0 Then
                        oLog.escribeMensaje("Se va a proceder a recorrer las SOCIEDADES...", EXO_Log.EXO_Log.Tipo.advertencia)
                        odtDatos2 = New System.Data.DataTable

                        For Each dr As DataRow In odtDatos.Rows
                            sSQLAUTPDTE = "SELECT * FROM (" ': ESprimero = True

#Region "Creamos SQL para recorrer y envial mail"

                            sBBDD = dr.Item("EXO_BD").ToString


                            sSQLAUTPDTE &= " select trim(substring(T5.""AliasName"", 0, 50)) ""BD"",  " +
                              " CASE WHEN T0.""ObjType"" = 22 then 'Pedido' else 'Factura' end ""Tipo"", " +
                              " Case WHEN t0.""ProcesStat"" = 'Y' THEN 'Aprobado'  else 'Rechazado' end ""Estado"",  " +
                               " t1.""Remarks"",   Case  When  T4.""DocTotalFC""<> 0 Then CAST((T4.""DocTotalFC"" - T4.""TotalExpFC"" - T4.""VatSumFC"") As Decimal(10,2)) " +
                                    " Else     CAST((T4.""DocTotal"" - T4.""TotalExpns"" - T4.""VatSum"") As Decimal(10,2)) " +
                                    " End ""Total Documento""," +
                               "  t4.""NumAtCard"" ""Referencia"",  t4.""CardName"", t4.""Project"",T4.""DocCur"" " +
                              " from """ + sBBDD + """.""OWDD"" t0  " +
                                  " inner Join """ + sBBDD + """.""WDD1"" T1 On T0.""WddCode"" = T1.""WddCode"" And t0.""CurrStep"" = t1.""StepCode""  " +
                                  " inner join """ + sBBDD + """.""OWST"" T2 On T1.""StepCode"" = t2.""WstCode""  " +
                                "  inner Join """ + sBBDD + """.""OUSR"" T3 On T0.""UserSign"" = t3.""USERID""  " +
                                  " inner join """ + sBBDD + """.""ODRF"" T4 On T0.""DraftEntry"" = t4.""DocEntry"" " +
                                "  inner Join """ + sBBDD + """.""OUDP"" T6 On T3.""Department"" = T6.""Code""  , " +
                                "             """ + sBBDD + """.""OADM"" T5 " +
                              " WHERE((t0.""Status"" = 'N' ) or(t0.""Status"" = 'Y' AND t0.""ProcesStat"" = 'Y'))   " +
                              " And T1.""UpdateDate""= CURRENT_DATE and T3.""USER_CODE"" ='" + sUsuario + "' "

#End Region
                            sSQLAUTPDTE &= ") T ORDER BY T.""Tipo"""

                            Conexiones.FillDtDB(oDBSAP, odtDatos2, sSQLAUTPDTE)
                        Next

                        Procesos.EnviarMails_statusaprobSMTP(oLog, odtDatos2, sUsuario, oDBSAP)

                    Else
                        oLog.escribeMensaje("No existen sociedades definidas", EXO_Log.EXO_Log.Tipo.error)
                        Exit Sub
                    End If
                Next
            Else
                oLog.escribeMensaje("No existen usuarios definidos", EXO_Log.EXO_Log.Tipo.error)
                Exit Sub
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            sError = exCOM.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            sError = ex.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally
            Conexiones.Disconnect_SQLHANA(oDBSAP)
            odtDatos = Nothing : odtUsuarios = Nothing
        End Try
    End Sub


    Private Shared Sub EnviarMails_statusaprobSMTP(oLog As EXO_Log.EXO_Log, odtDatos As DataTable, sUsuario As String, oDBSAP As HanaConnection)
#Region "Variables"
        Dim sError As String = ""
        Dim sEmpresa As String = ""
        Dim sHora1 As String = "" : Dim SHora2 As String = "" : Dim EsUrgente As Boolean = False
        Dim sMailFROM As String = "" : Dim sDirmail As String = "" : Dim cuerpo As String = ""
        Dim sSMTP As String = "" : Dim sPuerto As String = "" : Dim sUsMail As String = "" : Dim sPSSMail As String = ""
        Dim sDoc As String = "" : Dim sTipo As String = "" : Dim sBBDD As String = "" : Dim sUrgente As String = "" : Dim stabla As String = ""
        Dim bExAut As Boolean = False
        Dim sCambiaBBDD As String = "" : Dim sCambiaTDoc As String = ""
        Dim sTextoEnviado As String = ""
#End Region
        Try

            If odtDatos.Rows.Count > 0 Then
                oLog.escribeMensaje("Existen status aprobados pendientes de enviar " & sUsuario & ". ", EXO_Log.EXO_Log.Tipo.advertencia)

                sMailFROM = Conexiones.Datos_Confi("MAIL", "MAILFROM")
                sSMTP = Conexiones.Datos_Confi("MAIL", "SMTP") : sPuerto = Conexiones.Datos_Confi("MAIL", "PUERTO")
                sUsMail = Conexiones.Datos_Confi("MAIL", "USUARIO") : sPSSMail = Conexiones.Datos_Confi("MAIL", "PASS")
                sEmpresa = odtDatos.Rows(0).Item("BD").ToString
                oLog.escribeMensaje("Empresa: " & sEmpresa, EXO_Log.EXO_Log.Tipo.advertencia)
                Dim correo As New System.Net.Mail.MailMessage()
                'Comprobamos si estamos en tiempo de urgencia

                correo.From = New System.Net.Mail.MailAddress(sMailFROM, "Envío Aut. Autorizaciones Pdtes.")
                correo.To.Clear()
                sDirmail = Conexiones.GetValueDB(oDBSAP, """SOL_AUTORIZ"".""EXO_CREADORES""", """EXO_MAIL""", """EXO_USUARIO""='" & sUsuario & "' ")

                If sDirmail <> "" Then
                    correo.To.Add(sDirmail)

                    Dim strHeader As String = "<table><tbody>"
                    Dim strFooter As String = "</tbody></table>"
                    Dim sbContent As New StringBuilder()
                    sTextoEnviado = ""
                    sbContent.Append(String.Format("<td>{0}</td>", "https://autorizaciones.solariaenergia.com/"))

                    For Each dr As DataRow In odtDatos.Rows

#Region "Asignamos Varible sTipo y sTabla"

                        Select Case dr("Tipo").ToString
                            Case "Pedido" : sTipo = "Pedio de compra" : stabla = "ODRF"
                            Case "Factura" : sTipo = "Factura de proveedor" : stabla = "ODRF"
                        End Select
#End Region

                        oLog.escribeMensaje("Tabla: " & sTipo & " - DocEntry: " & dr.Item("Referencia").ToString, EXO_Log.EXO_Log.Tipo.advertencia)

                        If sCambiaTDoc <> sTipo Then
                            sCambiaTDoc = sTipo

#Region "Asignamos cabecera de la tabla Tipo Documento"
                            sbContent.Append("<tr>")

                            sbContent.Append("<tr>") : sbContent.Append(String.Format("<td>{0}</td>", "")) : sbContent.Append("</tr>")
                            sbContent.Append("<tr>") : sbContent.Append(String.Format("<td>{0}</td>", "T. Documento:" & sTipo)) : sbContent.Append("</tr>")
                            sbContent.Append("<tr>")

                            sbContent.Append("<tr>")

                            sbContent.Append(String.Format("<td>{0}</td>", "Empresa"))
                            sbContent.Append(String.Format("<td><p align=""right"">{0}</p></td>", "Estado"))
                            sbContent.Append(String.Format("<td><p align=""right"">{0}</p></td>", "Comentario"))
                            sbContent.Append(String.Format("<td><p align=""right"">{0}</p></td>", "Importe"))
                            sbContent.Append(String.Format("<td>{0}</td>", "Referencia"))
                            sbContent.Append(String.Format("<td>{0}</td>", "Proveedor"))
                            sbContent.Append(String.Format("<td>{0}</td>", "Proyecto"))
                            sbContent.Append("</tr>")
                            sbContent.Append("<tr>")
                            sbContent.Append(String.Format("<td>{0}</td>", "__________"))
                            sbContent.Append(String.Format("<td>{0}</td>", "______________"))
                            sbContent.Append(String.Format("<td>{0}</td>", "________________________________________"))
                            sbContent.Append(String.Format("<td>{0}</td>", "______________"))
                            sbContent.Append(String.Format("<td>{0}</td>", "______________"))
                            sbContent.Append(String.Format("<td>{0}</td>", "_________________"))
                            sbContent.Append(String.Format("<td>{0}</td>", "_________________"))
                            sbContent.Append("</tr>")

                            sTextoEnviado &= ChrW(10) '& ChrW(13)
                            sTextoEnviado &= "Empresa" & ChrW(9) & "Estado" & ChrW(9) & "Comentario" & ChrW(9) & "Importe" & ChrW(9) & "Referencia" & ChrW(9) & "Proveedor" & "Proyecto" & Chr(10) ' & ChrW(13)
                            sTextoEnviado &= "################################################################################################################################" & Chr(10) '& ChrW(13)                                  
#End Region
                        End If
                        sbContent.Append("<tr>")

                        sbContent.Append(String.Format("<td>{0}</td>", dr.Item("BD").ToString))
                        sbContent.Append(String.Format("<td><p align=""left"">{0}</p></td>", dr.Item("Estado").ToString))
                        sbContent.Append(String.Format("<td><p align=""left"">{0}</p></td>", dr.Item("Remarks").ToString))
                        sbContent.Append(String.Format("<td><p align=""right"">{0}</p></td>", CDbl(dr.Item("Total Documento").ToString).ToString("#,##0.00") & dr.Item("DocCur").ToString))
                        sbContent.Append(String.Format("<td><p align=""left"">{0}</p></td>", dr.Item("Referencia").ToString))
                        sbContent.Append(String.Format("<td><p align=""left"">{0}</p></td>", dr.Item("CardName").ToString))
                        sbContent.Append(String.Format("<td><p align=""left"">{0}</p></td>", dr.Item("Project").ToString))
                        sbContent.Append("</tr>")

                        sTextoEnviado &= sUrgente & ChrW(9) & dr.Item("BD").ToString & ChrW(9) & dr.Item("Estado").ToString & ChrW(9) & dr.Item("Remarks").ToString & ChrW(9) & CDbl(dr.Item("Total Documento").ToString).ToString("#,##0.00") & dr.Item("DocCur").ToString & ChrW(9) & dr.Item("Referencia").ToString & Chr(10) '& ChrW(13)

                        bExAut = True
                    Next

                    sbContent.Append(String.Format("<td>{0}</td>", ""))
                    sbContent.Append(String.Format("<td>{0}</td>", ""))
                    Dim emailTemplate As String = strHeader & sbContent.ToString() & strFooter
                    cuerpo &= strHeader & sbContent.ToString() & strFooter
                    correo.Body = cuerpo
                    correo.IsBodyHtml = True
                    correo.Priority = System.Net.Mail.MailPriority.Normal

                    Dim smtp As New System.Net.Mail.SmtpClient
                    smtp.Host = sSMTP
                    smtp.Port = sPuerto
                    smtp.UseDefaultCredentials = True
                    smtp.Credentials = New System.Net.NetworkCredential(sUsMail, sPSSMail)
                    smtp.EnableSsl = True
                    Try
                        'oLog.escribeMensaje(sTextoEnviado, EXO_Log.EXO_Log.Tipo.informacion)
                        ' Si existen autorizaciones se envía
                        If bExAut = True Then
                            'If sDirmail = "mperiz@expertone.es" Then
                            smtp.Send(correo)
                            oLog.escribeMensaje(sEmpresa & "- Correo enviado: " & sDirmail, EXO_Log.EXO_Log.Tipo.informacion)
                            oLog.escribeMensaje(sTextoEnviado, EXO_Log.EXO_Log.Tipo.informacion)
                            'End If
                        Else
                            oLog.escribeMensaje(sEmpresa & "- Correo No enviado: " & sDirmail, EXO_Log.EXO_Log.Tipo.advertencia)
                        End If

                        correo.Dispose()
                    Catch ex As Exception
                        sError = ex.Message
                        oLog.escribeMensaje(sEmpresa & " - Error enviando correo: " & sDirmail & ". " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try
                Else
                    oLog.escribeMensaje("El Usuario " & sUsuario & " no tiene asignado mail para envío de información", EXO_Log.EXO_Log.Tipo.error)
                    Exit Sub
                End If
            Else
                oLog.escribeMensaje("No existen autorizaciones pdtes de enviar del usuario " & sUsuario & ". ", EXO_Log.EXO_Log.Tipo.advertencia)
                Exit Sub
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            sError = exCOM.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            sError = ex.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally

        End Try
    End Sub

#End Region

End Class
