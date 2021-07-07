Public Class EXO_GLOBALES
#Region "Enumeraciones auxiliares"

    Public Enum FuenteInformacion
        Visual = 1
        Otros = 2
    End Enum

#End Region
#Region "Métodos auxiliares"

    Public Shared Function DblNumberToText(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByVal cValor As Double, ByVal oDestino As FuenteInformacion) As String
        Dim sNumberDouble As String = "0"

        DblNumberToText = "0"

        Try
            If cValor.ToString <> "" Then
                If oObjGlobal.refDi.OADM.separadorMillarB1 = "." AndAlso oObjGlobal.refDi.OADM.separadorDecimalB1 = "," Then 'Decimales ES
                    sNumberDouble = cValor.ToString
                Else 'Decimales USA
                    sNumberDouble = cValor.ToString.Replace(",", ".")
                End If
            End If

            If oDestino = FuenteInformacion.Visual Then
                If oObjGlobal.refDi.OADM.separadorDecimalSO = "," Then
                    DblNumberToText = sNumberDouble
                Else
                    DblNumberToText = sNumberDouble.Replace(".", ",")
                End If
            Else
                DblNumberToText = sNumberDouble.Replace(",", ".")
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function DblTextToNumber(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByVal sValor As String) As Double
        Dim cValor As Double = 0
        Dim sValorAux As String = "0"

        DblTextToNumber = 0

        Try
            sValorAux = sValor

            If oObjGlobal.refDi.OADM.separadorDecimalSO = "," Then
                If sValorAux <> "" Then
                    If Left(sValorAux, 1) = "." Then sValorAux = "0" & sValorAux

                    If oObjGlobal.refDi.OADM.separadorMillarB1 = "." AndAlso oObjGlobal.refDi.OADM.separadorDecimalB1 = "," Then 'Decimales ES
                        If sValorAux.IndexOf(".") > 0 AndAlso sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", ""))
                        ElseIf sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", ","))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    Else 'Decimales USA
                        If sValorAux.IndexOf(".") > 0 AndAlso sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", "").Replace(".", ","))
                        ElseIf sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", ","))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    End If
                End If
            Else
                If sValorAux <> "" Then
                    If Left(sValorAux, 1) = "," Then sValorAux = "0" & sValorAux

                    If oObjGlobal.refDi.OADM.separadorMillarB1 = "." AndAlso oObjGlobal.refDi.OADM.separadorDecimalB1 = "," Then 'Decimales ES
                        If sValorAux.IndexOf(",") > 0 AndAlso sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", "").Replace(",", "."))
                        ElseIf sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", "."))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    Else 'Decimales USA
                        If sValorAux.IndexOf(",") > 0 AndAlso sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", ""))
                        ElseIf sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", "."))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    End If
                End If
            End If

            DblTextToNumber = cValor

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function DateTextToDate(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByVal sFechaYYYYMMDD As String) As Date
        DateTextToDate = Nothing

        Try
            DateTextToDate = CDate(Right(sFechaYYYYMMDD, 2) & "/" & Mid(sFechaYYYYMMDD, 5, 2) & "/" & Left(sFechaYYYYMMDD, 4))

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function DateToText(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByVal dFecha As Date, Optional ByVal bFormatoYYYYMMDD As Boolean = False) As String
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        DateToText = dFecha.ToShortDateString

        Try
            oRs = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            If bFormatoYYYYMMDD = True Then
                DateToText = dFecha.Year & Right("0" & dFecha.Month.ToString, 2) & Right("0" & dFecha.Day.ToString, 2)
            Else
                oRs.DoQuery("SELECT ""DateFormat"" FROM ""OADM"" WHERE ""Code"" = 1")

                If oRs.RecordCount > 0 Then
                    If oRs.Fields.Item("DateFormat").Value.ToString = "0" OrElse oRs.Fields.Item("DateFormat").Value.ToString = "1" Then 'Modo ES
                        DateToText = dFecha.ToShortDateString
                    Else 'Modo USA
                        DateToText = Right("0" & dFecha.Month.ToString, 2) & "/" & Right("0" & dFecha.Day.ToString, 2) & "/" & dFecha.Year.ToString
                    End If
                End If
            End If


        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function

#End Region
End Class
