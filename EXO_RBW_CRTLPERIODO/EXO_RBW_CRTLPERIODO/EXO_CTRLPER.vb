Imports System.Xml
Imports SAPbouiCOM

Public Class EXO_CTRLPER
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

        cargamenu()
        If actualizar Then
            cargaCampos()
        End If
    End Sub
    Private Sub cargamenu()
        Dim Path As String = ""
        Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_MENU.xml")
        objGlobal.SBOApp.LoadBatchActions(menuXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults
    End Sub
    Private Sub cargaCampos()
        If objGlobal.refDi.comunes.esAdministrador Then
            Dim oXML As String = ""
            Dim udoObj As EXO_Generales.EXO_UDO = Nothing

            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_CTRLPER.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UDO_EXO_CTRLPER", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If
    End Sub
    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function
    Public Overrides Function menus() As System.Xml.XmlDocument
        Return Nothing
    End Function
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Try
            If infoEvento.BeforeAction = True Then

            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnACTCP"
                        'Cargamos UDO
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_CTRLPER")
                End Select
            End If

            Return MyBase.SBOApp_MenuEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_CTRLPER"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_CTRLPER"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                    If EventHandler_VALIDATE_Before(objGlobal, infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                            End Select

                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_CTRLPER"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    If EventHandler_FORM_VISIBLE(objGlobal, infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_CTRLPER"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                            End Select
                    End Select
                End If
            End If

            Return MyBase.SBOApp_ItemEvent(infoEvento)
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function
    Private Function EventHandler_FORM_VISIBLE(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oConds As SAPbouiCOM.Conditions = Nothing
        Dim oCond As SAPbouiCOM.Condition = Nothing

        EventHandler_FORM_VISIBLE = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True Then
                oForm.Freeze(True)
                CargaCombos(oForm)
                sSQL = "SELECT ""PrcCode"" FROM ""OPRC"" WHERE ""DimCode""=1 and ""Active""='Y' Order by ""PrcCode"" "
                oRs.DoQuery(sSQL)
                oConds = New SAPbouiCOM.Conditions
                'oCond = oConds.Add
                'oCond.Alias = "PrcCode"
                'oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                'oCond.CondVal = "0"
                'oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                For i = 0 To oRs.RecordCount - 1
                    oCond = oConds.Add
                    oCond.Alias = "PrcCode"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = oRs.Fields.Item("PrcCode").Value.ToString
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oRs.MoveNext()
                Next
                If oConds.Count > 0 Then oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_NONE
                oForm.ChooseFromLists.Item("CFLC").SetConditions(oConds)
            End If

            EventHandler_FORM_VISIBLE = True


        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function

    Private Function EventHandler_VALIDATE_Before(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sGrupo As String = ""
        EventHandler_VALIDATE_Before = False
        Dim sTable As String = ""
        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If pVal.ItemUID = "0_U_G" Then 'And pVal.ColUID = "C_0_1" Then
                If CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.ToString.Trim <> "" Then
                    sGrupo = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.ToString
                    'Comprobamos que en toda la matrix no haya mas de 1 codigo de grupo seleccionado
                    If MatrixToNet(oForm, sGrupo) = False Then
                        CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Active = True
                        Exit Function
                    End If
                End If
            End If


            EventHandler_VALIDATE_Before = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
#Region "Métodos auxiliares"
    Private Function CargaCombos(ByRef oForm As SAPbouiCOM.Form) As Boolean

        CargaCombos = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            oForm.Freeze(True)

            'Año
            sSQL = " Select ""Indicator"",""Indicator"" FROM ""OPID"" WHERE ""Indicator""<>'Valor de p' "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
                CType(oForm.Items.Item("0_U_E").Specific, SAPbouiCOM.ComboBox).ExpandType = BoExpandType.et_ValueOnly

            End If

            CargaCombos = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Private Function MatrixToNet(ByRef oForm As SAPbouiCOM.Form, ByVal sGrupo As String) As Boolean
        Dim sXML As String = ""
        Dim oMatrixXML As New Xml.XmlDocument
        Dim oXmlListRow As Xml.XmlNodeList = Nothing
        Dim oXmlListColumn As Xml.XmlNodeList = Nothing
        Dim oXmlNodeField As Xml.XmlNode = Nothing
        Dim sGrupoleido As String = "" : Dim iGrupoTotal As Integer = 0
        Dim oArrCampos As Boolean = False
        Dim sMatrixUID As String = ""

        MatrixToNet = False

        Try
            sXML = CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All)
            oMatrixXML.LoadXml(sXML)

            sMatrixUID = oMatrixXML.SelectSingleNode("//Matrix/UniqueID").InnerText
            oXmlListRow = oMatrixXML.SelectNodes("//Matrix/Rows/Row")
            iGrupoTotal = 0

            For Each oXmlNodeRow As Xml.XmlNode In oXmlListRow
                oXmlListColumn = oXmlNodeRow.SelectNodes("Columns/Column")

                'Inicializamos para de dejar a False

                oArrCampos = False

                'Inicializamos los datos del registro
                sGrupoleido = ""

                For Each oXmlNodeColumn As Xml.XmlNode In oXmlListColumn
                    oXmlNodeField = oXmlNodeColumn.SelectSingleNode("ID")

                    If oXmlNodeField.InnerXml = "C_0_1" Then 'CodigoGrupo
                        oXmlNodeField = oXmlNodeColumn.SelectSingleNode("Value")

                        sGrupoleido = oXmlNodeField.InnerText

                        oArrCampos = True
                        If sGrupo = sGrupoleido Then
                            iGrupoTotal += 1
                        End If
                    End If

                    If oArrCampos = True And iGrupoTotal >= 2 Then
                        Exit For
                    End If
                Next

                'Hemos recorrido el registro, y comprobamos el almacén
                If iGrupoTotal >= 2 Then
                    objGlobal.SBOApp.StatusBar.SetText("No es posible seleccionar el CeCo " & sGrupo & " varias veces", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                    Exit Function
                End If
            Next

            MatrixToNet = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

    Public Overrides Function SBOApp_FormDataEvent(infoEvento As BusinessObjectInfo) As Boolean
        Dim oform As SAPbouiCOM.Form = Nothing
        Try
            oform = CType(objGlobal.SBOApp.Forms.Item(infoEvento.FormUID), SAPbouiCOM.Form)
            If oform.TypeEx = "UDO_FT_EXO_CTRLPER" Then
                If infoEvento.BeforeAction = True Then
                    Select Case infoEvento.EventType
                        Case BoEventTypes.et_FORM_DATA_ADD
                        Case BoEventTypes.et_FORM_DATA_DELETE
                        Case BoEventTypes.et_FORM_DATA_UPDATE
                        Case BoEventTypes.et_FORM_DATA_LOAD

                    End Select
                Else
                    If infoEvento.ActionSuccess = True Then
                        Select Case infoEvento.EventType
                            Case BoEventTypes.et_FORM_DATA_ADD
                                If FORM_DATA_ADDUPDATE(infoEvento) = False Then
                                    GC.Collect()
                                    Return False
                                End If
                            Case BoEventTypes.et_FORM_DATA_DELETE
                            Case BoEventTypes.et_FORM_DATA_UPDATE
                                If FORM_DATA_ADDUPDATE(infoEvento) = False Then
                                    GC.Collect()
                                    Return False
                                End If
                            Case BoEventTypes.et_FORM_DATA_LOAD
                        End Select
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            objGlobal.SBOApp.MessageBox(ex.Message)
            Return False

        Finally
            If objGlobal.SBOApp.ClientType = BoClientType.ct_Desktop Then
                EXO_CleanCOM.CLiberaCOM.Form(oform)
            End If

        End Try
    End Function
    Private Function FORM_DATA_ADDUPDATE(ByRef pVal As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
        Dim sControl As String = ""
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsActivo As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

        Dim vItem As SAPbobsCOM.Items = Nothing
        Dim sCeCo As String = "" : Dim sAnno As String = ""
        FORM_DATA_ADDUPDATE = False
        Try
            sControl = oForm.DataSources.DBDataSources.Item("@EXO_CTRLPER").GetValue("Code", 0).ToUpper

            sSQL = "SELECT * FROM ""@EXO_CTRLPERL"" "
            sSQL &= " WHERE ""Code""=" & sControl & ";"
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                For i = 0 To oRs.RecordCount - 1
                    sCeCo = oRs.Fields.Item("U_EXO_CECO").Value.ToString
                    sAnno = oRs.Fields.Item("Code").Value.ToString
                    sSQL = " SELECT ""ItemCode"" FROM ""ITM6"" WHERE ""OcrCode""='" & sCeCo & "' and (ifnull(""ValidTo"",'')='' or year(""ValidTo"")='" & sAnno & "' ) "
                    oRsActivo.DoQuery(sSQL)
                    For a = 0 To oRsActivo.RecordCount - 1
                        vItem = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems), SAPbobsCOM.Items)

                        If vItem.GetByKey(oRsActivo.Fields.Item("ItemCode").Value.ToString) = False Then
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - No se encuentra el activo fijo """ & oRsActivo.Fields.Item("ItemCode").Value.ToString & """ ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            objGlobal.SBOApp.MessageBox("No se encuentra el activo fijo """ & oRsActivo.Fields.Item("ItemCode").Value.ToString & """ ")
                            Exit Function
                        Else
                            For p = 0 To vItem.PeriodControls.Count - 1
                                vItem.PeriodControls.SetCurrentLine(p)
                                If vItem.PeriodControls.FiscalYear = sAnno Then
                                    Dim dFactor As Double = 1
                                    Select Case vItem.PeriodControls.SubPeriod.ToString.Trim
                                        Case "1" : dFactor = oRs.Fields.Item("U_EXO_ENE").Value
                                        Case "2" : dFactor = oRs.Fields.Item("U_EXO_FEB").Value
                                        Case "3" : dFactor = oRs.Fields.Item("U_EXO_MAR").Value
                                        Case "4" : dFactor = oRs.Fields.Item("U_EXO_ABR").Value
                                        Case "5" : dFactor = oRs.Fields.Item("U_EXO_MAY").Value
                                        Case "6" : dFactor = oRs.Fields.Item("U_EXO_JUN").Value
                                        Case "7" : dFactor = oRs.Fields.Item("U_EXO_JUL").Value
                                        Case "8" : dFactor = oRs.Fields.Item("U_EXO_AGO").Value
                                        Case "9" : dFactor = oRs.Fields.Item("U_EXO_SEP").Value
                                        Case "10" : dFactor = oRs.Fields.Item("U_EXO_OCT").Value
                                        Case "11" : dFactor = oRs.Fields.Item("U_EXO_NOV").Value
                                        Case "12" : dFactor = oRs.Fields.Item("U_EXO_DIC").Value
                                    End Select
                                    vItem.PeriodControls.Factor = dFactor
                                    If dFactor = 0 Then
                                        vItem.PeriodControls.DepreciationStatus = SAPbobsCOM.BoYesNoEnum.tNO
                                    Else
                                        vItem.PeriodControls.DepreciationStatus = SAPbobsCOM.BoYesNoEnum.tYES
                                    End If
                                End If
                            Next
                            If vItem.Update() <> 0 Then
                                Throw New Exception(objGlobal.compañia.GetLastErrorCode & " / No se puede actualizar los períodos de control del Activo Fijo -" & oRsActivo.Fields.Item("ItemCode").Value.ToString & " - " & objGlobal.compañia.GetLastErrorDescription)
                            Else
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Se ha actualizado los períodos de control del Activo fijo - " & oRsActivo.Fields.Item("ItemCode").Value.ToString & ". Por favor, revise los datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If
                        End If
                        oRsActivo.MoveNext()
                    Next


                    oRs.MoveNext()
                Next
            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Ha ocurrido un error inesperado y no se encuentra los registros de actualización de períodos de conrtrol.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objGlobal.SBOApp.MessageBox("Ha ocurrido un error inesperado y no se encuentra los registros de actualización de períodos de conrtrol.")
            End If
            FORM_DATA_ADDUPDATE = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsActivo, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(vItem, Object))
        End Try
    End Function
End Class
