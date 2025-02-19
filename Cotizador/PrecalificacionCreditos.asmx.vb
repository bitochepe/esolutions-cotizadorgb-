Imports System.ComponentModel
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Xml
Imports System.IO
Imports System
Imports System.Linq
Imports System.Collections.Generic
Imports System.Text.Json
Imports Newtonsoft.Json

<System.Web.Script.Services.ScriptService()>
<System.Web.Services.WebService(Namespace:="http://Cotizador.com/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class PrecalificacionCreditos
    Inherits System.Web.Services.WebService

    <WebMethod()>
    Public Function ObtenerDatos(dcmMontoSolicitado As Decimal, strDestinoCredito As String, strTipoGarantia As String, strActividadEconomica As String,
                                 dcmIngresoConstancia As Decimal, strActividadEconomica2 As String, strIngresoConstancia2 As String,
                                 strAuxilioPostumo As String, strMontepio As String, strSeguros As String, strDescuentoConstancia As String,
                                 strTipoDeuda1 As String, strSaldoDeuda1 As String, strCuotaDeuda1 As String,
                                 strTipoDeuda2 As String, strSaldoDeuda2 As String, strCuotaDeuda2 As String,
                                 strTipoDeuda3 As String, strSaldoDeuda3 As String, strCuotaDeuda3 As String,
                                 strTipoDeuda4 As String, strSaldoDeuda4 As String, strCuotaDeuda4 As String,
                                 strTipoDeuda5 As String, strSaldoDeuda5 As String, strCuotaDeuda5 As String,
                                 strTipoDeuda6 As String, strSaldoDeuda6 As String, strCuotaDeuda6 As String,
                                 strMes1 As String, strMes2 As String, strMes3 As String, strTerreno As String, strConstrucciones As String,
                                 strScorePredictivo As String, strClasificacionSIB As String, strConteoCCR As String) As XmlDocument

        Dim xmlRespuesta As New XmlDocument
        Dim xDoc As XDocument
        Dim dcmTasaInteres As Decimal
        Dim intPlazoMeses As Integer
        Dim dcmBonificacionAE As Decimal
        Dim dcmIgssAE As Decimal
        Dim dcmIsrAE As Decimal
        Dim dcmBonificacionAE2 As Decimal
        Dim dcmIgssAE2 As Decimal
        Dim dcmIsrAE2 As Decimal
        Dim dcmRDG As Decimal
        Dim dcmValorEGarantiaH As Decimal
        Dim dcmOtrasDeducciones As Decimal
        Dim dcmTotalS As Decimal
        Dim dcmTotalSalario1 As Decimal
        Dim dcmTotalSalario2 As Decimal
        Dim dcmTotalSalarios As Decimal
        Dim dcmCuota As Decimal
        Dim dcmEndeudamientoDirecto As Decimal
        Dim dcmTotalCuotasDirectas As Decimal
        Dim intRCI1 As Integer
        Dim intRCI2 As Integer
        Dim dcmValorGarantia As Decimal
        Dim dcmLIPFHA1 As Decimal
        Dim dcmLIPFHA2 As Decimal
        Dim dcmLIPDIRECTO1 As Decimal
        Dim dcmLIPDIRECTO2 As Decimal
        Dim strValorDiferente As String = ""
        Dim strRCIDiferente As String = ""
        Dim strGuiaRegistro As String
        Dim strParametros As Dictionary(Of String, Object)
        Dim strParametrosDetalle As String
        Dim strRutaArchivo As String
        Dim strFechaHora As String
        Dim strDetalle As String = ""


        Try

            strParametros = New Dictionary(Of String, Object) From {
                {"Monto solicitado", dcmMontoSolicitado},
                {"Destino del Crédito", strDestinoCredito}, {"Tipo de Garantía", strTipoGarantia}, {"Actividad Económica 1", strActividadEconomica},
                {"Ingresos según Constancia", dcmIngresoConstancia}, {"Actividad Económica 2", strActividadEconomica2}, {"Ingresos según Constancia 2", strIngresoConstancia2},
                {"Auxilio Postumo", strAuxilioPostumo}, {"Montepio", strMontepio}, {"Seguros", strSeguros}, {"Otros Descuentos de Cosntancia", strDescuentoConstancia},
                {"Tipo Deuda1", strTipoDeuda1}, {"Capital Original1", strSaldoDeuda1}, {"Cuota Deuda1", strCuotaDeuda1},
                {"Tipo Deuda2", strTipoDeuda2}, {"Capital Original2", strSaldoDeuda2}, {"Cuota Deuda2", strCuotaDeuda2},
                {"Tipo Deuda3", strTipoDeuda3}, {"Capital Original3", strSaldoDeuda3}, {"Cuota Deuda3", strCuotaDeuda3},
                {"Tipo Deuda4", strTipoDeuda4}, {"Capital Original4", strSaldoDeuda4}, {"Cuota Deuda4", strCuotaDeuda4},
                {"Tipo Deuda5", strTipoDeuda5}, {"Capital Original5", strSaldoDeuda5}, {"Cuota Deuda5", strCuotaDeuda5},
                {"Tipo Deuda6", strTipoDeuda6}, {"Capital Original6", strSaldoDeuda6}, {"Cuota Deuda6", strCuotaDeuda6},
                {"Mes 1", strMes1}, {"Mes 2", strMes2}, {"Mes 3", strMes3}, {"Terreno", strTerreno}, {"Construcciones", strConstrucciones},
                {"Score predictivo", strScorePredictivo}, {"Clasificación SIB", strClasificacionSIB}, {"Conteo de CCR", strConteoCCR}
            }

            'strParametrosDetalle = String.Join(", ", strParametros.Select(Function(kvp) $"{kvp.Key} = {kvp.Value}"))
            strGuiaRegistro = Guid.NewGuid().ToString()
            strFechaHora = DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")
            strRutaArchivo = Server.MapPath("~/Archivos/Logs/" + DateTime.Now.ToString("dd-MM-yyyy HH.mm.ss") + ".log")
            strParametrosDetalle = JsonConvert.SerializeObject(strParametros)
            'strParametrosDetalle = Json.JsonSerializer.Serialize(strParametros, New JsonSerializerOptions With {
            '    .WriteIndented = True
            '})
            EscribirRegistro("ENTRADA", strGuiaRegistro, strParametrosDetalle, strRutaArchivo)

            If Not String.IsNullOrEmpty(strDestinoCredito) Then
                dcmTasaInteres = ObtenerTasaXDestinoCredito(strDestinoCredito)
            End If
            If Not String.IsNullOrEmpty(strTipoGarantia) Then
                intPlazoMeses = ObtenerPlazoMesesXTipoGarantia(strTipoGarantia)

                'If strTipoGarantia = "Usada" Then
                '    dcmRDG = 75
                'ElseIf strTipoGarantia = "Nueva" Then
                '    dcmRDG = 80
                'End If
            End If

            If Not String.IsNullOrEmpty(strActividadEconomica) Then
                If Not String.IsNullOrEmpty(dcmIngresoConstancia) Then
                    dcmIsrAE = (dcmIngresoConstancia * ObtenerIgss_Isr("Isr")) / 100

                    If Not String.IsNullOrEmpty(strDescuentoConstancia) Then
                        dcmOtrasDeducciones += Convert.ToDecimal(strDescuentoConstancia)
                    End If
                    If Not String.IsNullOrEmpty(strAuxilioPostumo) Then
                        dcmOtrasDeducciones += Convert.ToDecimal(strAuxilioPostumo)
                    End If
                    If Not String.IsNullOrEmpty(strMontepio) Then
                        dcmOtrasDeducciones += Convert.ToDecimal(strMontepio)
                    End If
                    If Not String.IsNullOrEmpty(strSeguros) Then
                        dcmOtrasDeducciones += Convert.ToDecimal(strSeguros)
                    End If

                    If String.Equals(strActividadEconomica, "Relacion de dependencia", StringComparison.OrdinalIgnoreCase) Then
                        dcmBonificacionAE = 250
                        dcmIgssAE = (dcmIngresoConstancia * ObtenerIgss_Isr("Igss")) / 100

                        dcmTotalS = dcmIngresoConstancia + dcmBonificacionAE - dcmIgssAE - dcmIsrAE - dcmOtrasDeducciones
                        dcmTotalSalario1 = Math.Round((dcmTotalS * 14) / 12, 2)
                    Else
                        dcmBonificacionAE = 0
                        dcmIgssAE = 0

                        dcmTotalS = dcmIngresoConstancia - dcmIsrAE - dcmOtrasDeducciones
                        dcmTotalSalario1 = Math.Round(dcmTotalS * 0.6, 2)
                    End If
                End If
            End If
            If Not String.IsNullOrEmpty(strActividadEconomica2) Then
                If Not String.IsNullOrEmpty(strIngresoConstancia2) Then
                    dcmIsrAE2 = (Convert.ToDecimal(strIngresoConstancia2) * ObtenerIgss_Isr("Isr")) / 100

                    If String.Equals(strActividadEconomica2, "Relacion de dependencia", StringComparison.OrdinalIgnoreCase) Then
                        dcmBonificacionAE2 = 250
                        dcmIgssAE2 = (Convert.ToDecimal(strIngresoConstancia2) * ObtenerIgss_Isr("Igss")) / 100

                        dcmTotalS = Convert.ToDecimal(strIngresoConstancia2) + dcmBonificacionAE2 - dcmIgssAE2 - dcmIsrAE2
                        dcmTotalSalario2 = Math.Round((dcmTotalS * 14) / 12, 2)
                    Else
                        dcmBonificacionAE2 = 0
                        dcmIgssAE2 = 0

                        dcmTotalS = Convert.ToDecimal(strIngresoConstancia2) - dcmIsrAE2
                        dcmTotalSalario2 = Math.Round(dcmTotalS * 0.6, 2)
                    End If
                End If
            End If

            dcmEndeudamientoDirecto += SumaEndeudamientoDirecto(strTipoDeuda1, strSaldoDeuda1, strCuotaDeuda1)
            dcmEndeudamientoDirecto += SumaEndeudamientoDirecto(strTipoDeuda2, strSaldoDeuda2, strCuotaDeuda2)
            dcmEndeudamientoDirecto += SumaEndeudamientoDirecto(strTipoDeuda3, strSaldoDeuda3, strCuotaDeuda3)
            dcmEndeudamientoDirecto += SumaEndeudamientoDirecto(strTipoDeuda4, strSaldoDeuda4, strCuotaDeuda4)
            dcmEndeudamientoDirecto += SumaEndeudamientoDirecto(strTipoDeuda5, strSaldoDeuda5, strCuotaDeuda5)
            dcmEndeudamientoDirecto += SumaEndeudamientoDirecto(strTipoDeuda6, strSaldoDeuda6, strCuotaDeuda6)

            dcmTotalSalarios = dcmTotalSalario1 + dcmTotalSalario2

            If String.Equals(strTipoGarantia, "LIP FHA", StringComparison.OrdinalIgnoreCase) Then
                Dim tasaConjunta = dcmTasaInteres
                Dim seguroHipoteca = 1.26
                Dim tasaReferencia = 5.98 'Preguntar de donde se obtiene
                Dim tasaBanco = tasaConjunta - seguroHipoteca
                Dim tasa1al48 = (tasaBanco / 100) - (40 / 100) * (tasaReferencia / 100)
                Dim dcmValorMatricular As Decimal
                Dim dcmValorConstruccion As Decimal

                Dim cuotaSinSeguro = CalcularPago((tasaConjunta / 100) / 12, intPlazoMeses, dcmMontoSolicitado)
                Dim interesCliente = Math.Round(dcmMontoSolicitado * tasa1al48 / 12, 2) 'Columna E11
                Dim interesSubsidio = Math.Round(dcmMontoSolicitado * tasaBanco / 100 / 12, 2) - interesCliente 'Columna F11
                Dim S_H = Math.Round(dcmMontoSolicitado * seguroHipoteca / 100 / 12, 2) 'Columna G11
                Dim IUSI = 0.009 * dcmValorMatricular / 12
                Dim seguro = 0.28 / 100 * dcmValorConstruccion / 12
                Dim capital = cuotaSinSeguro - interesCliente - S_H - interesSubsidio
                Dim pagoTotal = capital + interesCliente + interesSubsidio + S_H + IUSI + seguro
                Dim pagoCliente = pagoTotal - interesSubsidio

                dcmLIPFHA1 = pagoCliente + dcmEndeudamientoDirecto
                dcmLIPFHA2 = pagoTotal + dcmEndeudamientoDirecto
                intRCI1 = (dcmLIPFHA1 / dcmTotalSalarios) * 100
                intRCI2 = (dcmLIPFHA2 / dcmTotalSalarios) * 100

                If intRCI1 > 50 Then
                    strDetalle += "ALTO RCI INICIAL"
                End If
                If intRCI2 > 65 Then
                    strDetalle += " | ALTO RCI DESPUES DEL SUBSIDIO"
                End If

                strValorDiferente = dcmLIPFHA1 & " | " & dcmLIPFHA2
                strRCIDiferente = Math.Round(intRCI1, MidpointRounding.AwayFromZero) & "% | " & Math.Round(intRCI2, MidpointRounding.AwayFromZero) & "%"

            ElseIf String.Equals(strTipoGarantia, "LIP DIRECTO", StringComparison.OrdinalIgnoreCase) Then
                Dim tasaBanco = dcmTasaInteres
                Dim tasaReferencia = 5.89 'Preguntar de donde se obtiene
                Dim tasa1al48 = (tasaBanco / 100) - (40 / 100) * (tasaReferencia / 100)
                'Dim dcmConVida As Decimal
                Dim dcmValorConstruccion As Decimal

                Dim cuotaSinSeguro = CalcularPago((tasaBanco / 100) / 12, intPlazoMeses, dcmMontoSolicitado)
                Dim interesCliente = Math.Round(dcmMontoSolicitado * tasa1al48 / 12, 2)
                Dim interesSubsidio = Math.Round(dcmMontoSolicitado * tasaBanco / 100 / 12, 2) - interesCliente
                Dim seguro = 0.32 / 100 * dcmValorConstruccion / 12
                Dim vida = 0.0051072 * dcmMontoSolicitado / 12
                Dim capital = cuotaSinSeguro - interesCliente - interesSubsidio
                Dim pagoTotal = capital + interesCliente + interesSubsidio + vida + seguro
                Dim pagoCliente = pagoTotal - interesSubsidio

                dcmLIPDIRECTO1 = pagoCliente + dcmEndeudamientoDirecto
                dcmLIPDIRECTO2 = pagoTotal + dcmEndeudamientoDirecto
                intRCI1 = (dcmLIPDIRECTO1 / dcmTotalSalarios) * 100
                intRCI2 = (dcmLIPDIRECTO2 / dcmTotalSalarios) * 100

                If intRCI1 > 50 Then
                    strDetalle += "ALTO RCI INICIAL"
                End If
                If intRCI2 > 65 Then
                    strDetalle += " | ALTO RCI DESPUES DEL SUBSIDIO"
                End If

                strValorDiferente = dcmLIPDIRECTO1 & " | " & dcmLIPDIRECTO2
                strRCIDiferente = Math.Round(intRCI1, MidpointRounding.AwayFromZero) & "% | " & Math.Round(intRCI2, MidpointRounding.AwayFromZero) & "%"

            Else
                dcmCuota = Math.Round(CalcularPago(dcmTasaInteres / 100 / 12, intPlazoMeses, dcmMontoSolicitado), 2)
                dcmTotalCuotasDirectas = Math.Round(dcmCuota + dcmEndeudamientoDirecto + PolizadeSeguro(dcmMontoSolicitado, strConstrucciones, dcmTasaInteres), 2)

                intRCI1 = (dcmTotalCuotasDirectas / dcmTotalSalarios) * 100

                If intRCI1 > 50 Then
                    strDetalle += "Alto RCI, Precalificación no viable"
                Else
                    strDetalle += "RCI Favorable"
                End If

                intRCI1 = Math.Round(intRCI1, MidpointRounding.AwayFromZero)

            End If

            If intRCI1 > 50 Or intRCI2 > 50 Then
                If dcmEndeudamientoDirecto <> 0 Then
                    strDetalle += " | Se recomienda consolidar todas las deudas para bajar el RCI"
                End If
            End If

            Dim dcmIngresosVerificados As Decimal

            If Not String.IsNullOrEmpty(strMes1) Then
                dcmIngresosVerificados += Convert.ToDecimal(strMes1)
            End If
            If Not String.IsNullOrEmpty(strMes2) Then
                dcmIngresosVerificados += Convert.ToDecimal(strMes2)
            End If
            If Not String.IsNullOrEmpty(strMes3) Then
                dcmIngresosVerificados += Convert.ToDecimal(strMes3)
            End If

            Dim dcmPromedio = dcmIngresosVerificados / 3
            Dim dcmValidacionIngresos = Math.Round((dcmPromedio / dcmTotalSalarios) * 100, MidpointRounding.AwayFromZero)

            If dcmValidacionIngresos < 75 Then
                strDetalle += " | Los ingresos no se comprueban en estados de cuenta"
            End If

            If Not String.IsNullOrEmpty(strConstrucciones) Then
                'dcmValorGarantia = dcmTerreno + dcmConstrucciones
                dcmValorGarantia = Convert.ToDecimal(strConstrucciones)
            End If

            If dcmValorGarantia <> 0 Then
                dcmRDG = Math.Round(dcmMontoSolicitado / dcmValorGarantia * 100, MidpointRounding.AwayFromZero)
            End If

            If String.Equals(strDestinoCredito, "Compra de Terreno", StringComparison.OrdinalIgnoreCase) Then
                dcmValorEGarantiaH = Math.Round(dcmMontoSolicitado * 100 / 60, 2)
            Else
                dcmValorEGarantiaH = Math.Round(dcmMontoSolicitado * 100 / 80, 2)
            End If

            Dim dcmHipoteca As Decimal = 0

            If Not String.IsNullOrEmpty(strTerreno) Then
                dcmValorGarantia += Convert.ToDecimal(strTerreno)
            End If
            If Not String.IsNullOrEmpty(strConstrucciones) Then
                dcmValorGarantia += Convert.ToDecimal(strConstrucciones)
            End If
            If dcmValorGarantia <> 0 Then
                dcmHipoteca = (dcmMontoSolicitado / dcmValorGarantia) * 100
            End If
            If dcmHipoteca < 60 Then
                strDetalle += " | Faltante de garantia"
            End If

            Context.Response.ContentType = "application/xml"
            xmlRespuesta.LoadXml($"<?xml version=""1.0""?><Root>
                                  <tasaInteres>{dcmTasaInteres}%</tasaInteres>
                                  <plazoMeses>{intPlazoMeses}</plazoMeses>
                                  <noCuota>{dcmCuota}</noCuota>
                                  <rci>{intRCI1}%</rci>
                                  <bonificacionAE>{dcmBonificacionAE}</bonificacionAE>
                                  <igssAE>{Math.Round(dcmIgssAE, 2)}</igssAE>
                                  <isrAE>{Math.Round(dcmIsrAE, 2)}</isrAE>
                                  <bonificacionAE2>{dcmBonificacionAE2}</bonificacionAE2>
                                  <igssAE2>{Math.Round(dcmIgssAE2, 2)}</igssAE2>
                                  <isrAE2>{Math.Round(dcmIsrAE2, 2)}</isrAE2>
                                  <cuota>{dcmCuota}</cuota>
                                  <totalCuentasDirectas>{dcmTotalCuotasDirectas}</totalCuentasDirectas>
                                  <valorGarantiaHipotecaria>{dcmValorEGarantiaH}</valorGarantiaHipotecaria>
                                  <rdg>{dcmRDG}%</rdg>
                                  <valorDiferente>{strValorDiferente}</valorDiferente>
                                  <rciDiferente>{strRCIDiferente}</rciDiferente>
                                  <trfLip></trfLip>
                                  <detalle>{strDetalle}</detalle>
                              </Root>")


            xDoc = XDocument.Parse(xmlRespuesta.OuterXml)
            strParametrosDetalle = JsonConvert.SerializeXNode(xDoc, Xml.Formatting.Indented)
            strParametrosDetalle = JsonConvert.SerializeObject(xDoc)
            EscribirRegistro("SALIDA", strGuiaRegistro, strParametrosDetalle, strRutaArchivo)

            Return xmlRespuesta

        Catch ex As Exception
            xmlRespuesta.LoadXml($"<respuesta>
                                  <error>{ex.Message}</error>
                              </respuesta>")

            Return xmlRespuesta
        End Try
    End Function

    Function ObtenerTasaXDestinoCredito(ByVal strDestinoCredito As String) As Decimal
        Dim strRutaArchivo As String = Server.MapPath("~/Archivos/TasaInteres_DestinoCredito.xml")
        Dim documento As XDocument = XDocument.Load(strRutaArchivo)

        Dim strValor As String = (From entrada In documento.Descendants("Entrada")
                                  Where String.Equals(entrada.Element("Texto").Value, strDestinoCredito, StringComparison.OrdinalIgnoreCase)
                                  Select entrada.Element("Valor").Value).FirstOrDefault()

        If String.IsNullOrEmpty(strValor) Then
            Throw New ArgumentException("El texto proporcionado no está en el archivo XML.")
        End If

        Return Convert.ToDecimal(strValor)
    End Function

    Function ObtenerPlazoMesesXTipoGarantia(ByVal strTipoGarantia As String) As Decimal
        Dim strRutaArchivo As String = Server.MapPath("~/Archivos/PlazoMeses_TipoGarantia.xml")
        Dim documento As XDocument = XDocument.Load(strRutaArchivo)

        Dim strValor As String = (From entrada In documento.Descendants("Entrada")
                                  Where String.Equals(entrada.Element("Texto").Value, strTipoGarantia, StringComparison.OrdinalIgnoreCase)
                                  Select entrada.Element("Valor").Value).FirstOrDefault()

        If String.IsNullOrEmpty(strValor) Then
            Throw New ArgumentException("El texto proporcionado no está en el archivo XML.")
        End If

        Return Convert.ToDecimal(strValor)
    End Function

    Function ObtenerIgss_Isr(ByVal strImpuesto As String) As Decimal
        Dim strRutaArchivo As String = Server.MapPath("~/Archivos/Igss_Isr.xml")
        Dim documento As XDocument = XDocument.Load(strRutaArchivo)

        Dim strValor As String = (From entrada In documento.Descendants("Entrada")
                                  Where String.Equals(entrada.Element("Texto").Value, strImpuesto, StringComparison.OrdinalIgnoreCase)
                                  Select entrada.Element("Valor").Value).FirstOrDefault()

        If String.IsNullOrEmpty(strValor) Then
            Throw New ArgumentException("El texto proporcionado no está en el archivo XML.")
        End If

        Return Convert.ToDecimal(strValor)
    End Function

    Function CalcularPago(dcmTasaMensual As Decimal, dcmTotalPeriodos As Integer, dcmValorPresente As Decimal) As Decimal
        Dim pago As Decimal

        If dcmTasaMensual = 0 Then
            Return dcmValorPresente / dcmTotalPeriodos
        End If

        pago = (dcmTasaMensual * dcmValorPresente) / (1 - Math.Pow(1 + dcmTasaMensual, -dcmTotalPeriodos))

        Return pago
    End Function

    Function SumaEndeudamientoDirecto(strTipoDeuda As String, strSaldoDeuda As String, strCuotaDeuda As String)
        Dim dcmEndeudamientoDirecto As Decimal
        dcmEndeudamientoDirecto = 0

        If Not String.IsNullOrEmpty(strTipoDeuda) Then
            If Not String.IsNullOrEmpty(strCuotaDeuda) Then
                dcmEndeudamientoDirecto = Convert.ToDecimal(strCuotaDeuda)
            ElseIf Not String.IsNullOrEmpty(strSaldoDeuda) Then
                dcmEndeudamientoDirecto = CalcularCuotaDeudaXTipoDeuda(strTipoDeuda, Convert.ToDecimal(strSaldoDeuda))
            End If
        End If
        Return dcmEndeudamientoDirecto
    End Function

    Function CalcularCuotaDeudaXTipoDeuda(strTipoDeuda As String, dcmSaldoDeuda As Decimal)
        Dim dcmCuota As Decimal
        Dim intPlazo As Integer
        Dim dcmTasaPromedio As Decimal

        ObtenerPlazoTasaPromedioXTipo(strTipoDeuda, intPlazo, dcmTasaPromedio)

        Select Case strTipoDeuda
            Case "Tarjeta de Credito Interna"
                dcmCuota = (dcmSaldoDeuda * 0.33 / intPlazo) + (dcmSaldoDeuda * 0.033) * (dcmTasaPromedio / 12)
            Case "Tarjeta de Credito Mayor a 9 m"
                dcmCuota = CalcularPago(dcmTasaPromedio / 12, intPlazo, dcmSaldoDeuda)
            Case "Tarjeta de Credito Menor a 9m"
                dcmCuota = CalcularPago(dcmTasaPromedio / 12, intPlazo, dcmSaldoDeuda * 0.33)
            Case "Prestamo Hipotecario"
                dcmCuota = (CalcularPago(dcmTasaPromedio / 12, intPlazo, dcmSaldoDeuda)) / 2
            Case "Prestamo Fiduciario"
                dcmCuota = CalcularPago(dcmTasaPromedio / 12, intPlazo, dcmSaldoDeuda)
            Case "Prestamo Fiduciario Indirecto Bantrab"
                dcmCuota = CalcularPago(dcmTasaPromedio / 12, intPlazo, dcmSaldoDeuda * 0.2)
            Case "Factorje Fiduciario"
                dcmCuota = CalcularPago(dcmTasaPromedio / 12, intPlazo, dcmSaldoDeuda)
            Case "Prestamo Bienes Inmuebles"
                dcmCuota = (CalcularPago(dcmTasaPromedio / 12, intPlazo, dcmSaldoDeuda)) / 2
            Case "Prestamo Bienes Inmuebles - Prendas"
                dcmCuota = CalcularPago(dcmTasaPromedio / 12, intPlazo, dcmSaldoDeuda)
            Case "Prestamo Fiduciario - Prendas"
                dcmCuota = CalcularPago(dcmTasaPromedio / 12, intPlazo, dcmSaldoDeuda)
            Case "Fiduciaria Otras Garantias"
                dcmCuota = CalcularPago(dcmTasaPromedio / 12, intPlazo, dcmSaldoDeuda)
        End Select

        Return dcmCuota
    End Function

    Function PolizadeSeguro(dcmMontoSolicitado As Decimal, strConstrucciones As String, dcmTasaInteres As Decimal)
        Dim strRutaArchivo As String = Server.MapPath("~/Archivos/PolizaSeguro.xml")
        Dim documento As XDocument = XDocument.Load(strRutaArchivo)
        Dim dcmDanios As Decimal
        Dim dcmVidaTotal As Decimal
        Dim dcmVidaMes As Decimal
        Dim dcmPrimaNeta As Decimal
        Dim dcmGastosEmision As Decimal
        Dim dcmDecreto1422 As Decimal
        Dim dcmIVA As Decimal
        Dim dcmFraccionamiento As Decimal
        Dim dcmPrimaTotal As Decimal
        Dim dcmPrimaMes As Decimal
        Dim dcmIntereses As Decimal
        Dim dcmGastos As Decimal

        Dim strValor = (From entrada In documento.Descendants("Entrada")
                        Where String.Equals(entrada.Element("Texto").Value, "PolizaSeguro", StringComparison.OrdinalIgnoreCase)
                        Select New With {
                            .Vida = entrada.Element("Vida").Value,
                            .PrimaNeta = entrada.Element("PrimaNeta").Value,
                            .GastosEmision = entrada.Element("GastosEmision").Value,
                            .Decreto_1422 = entrada.Element("Decreto_1422").Value,
                            .IVA = entrada.Element("IVA").Value,
                            .Fraccionamiento = entrada.Element("Fraccionamiento").Value
                        }).FirstOrDefault()

        dcmVidaTotal = dcmMontoSolicitado * Convert.ToDecimal(strValor.Vida)
        dcmVidaMes = dcmVidaTotal / 12
        dcmDanios = dcmMontoSolicitado * Convert.ToDecimal(strValor.PrimaNeta) / 100

        If Not String.IsNullOrEmpty(strConstrucciones) Then
            dcmPrimaNeta = Convert.ToDecimal(strConstrucciones) * Convert.ToDecimal(strValor.PrimaNeta) / 100
        Else
            dcmPrimaNeta = dcmMontoSolicitado * Convert.ToDecimal(strValor.PrimaNeta) / 100
        End If

        dcmGastosEmision = dcmPrimaNeta * Convert.ToDecimal(strValor.GastosEmision) / 100
        dcmDecreto1422 = dcmPrimaNeta * Convert.ToDecimal(strValor.Decreto_1422) / 100
        dcmFraccionamiento = dcmPrimaNeta * Convert.ToDecimal(strValor.Fraccionamiento) / 100
        dcmIVA = (dcmPrimaNeta + dcmGastosEmision + dcmFraccionamiento) * Convert.ToDecimal(strValor.IVA) / 100

        dcmPrimaTotal = dcmPrimaNeta + dcmGastosEmision + dcmDecreto1422 + dcmIVA + dcmFraccionamiento
        dcmPrimaMes = dcmPrimaTotal / 12
        dcmVidaMes = dcmVidaTotal / 12
        dcmIntereses = dcmMontoSolicitado * dcmTasaInteres / 100 / 360 * 30

        dcmGastos = dcmIntereses + dcmPrimaTotal + dcmVidaTotal

        Return dcmVidaMes + dcmPrimaMes
    End Function

    Sub ObtenerPlazoTasaPromedioXTipo(strTipoDeuda As String, ByRef intPlazo As Integer, ByRef dcmTasaPromedio As Decimal)
        Dim strRutaArchivo As String = Server.MapPath("~/Archivos/PlazoTasaPromedio_TipoDeuda.xml")
        Dim documento As XDocument = XDocument.Load(strRutaArchivo)

        Dim strValor = (From entrada In documento.Descendants("Entrada")
                        Where String.Equals(entrada.Element("Texto").Value, strTipoDeuda, StringComparison.OrdinalIgnoreCase)
                        Select New With {
                              .intPlazo = entrada.Element("Plazo").Value,
                              .dcmTasaPromedio = entrada.Element("Tasa").Value
                        }).FirstOrDefault()

        If String.IsNullOrEmpty(strValor.intPlazo) Then
            intPlazo = 0
            dcmTasaPromedio = 0
        Else
            intPlazo = strValor.intPlazo
            dcmTasaPromedio = strValor.dcmTasaPromedio
        End If
    End Sub

    Public Sub EscribirRegistro(strTitulo As String, strGuiaRegistro As String, strMensaje As String, strRutaArchivo As String)

        Try
            Using writer As New StreamWriter(strRutaArchivo, True)
                writer.WriteLine(strTitulo)
            End Using

            Using writer As StreamWriter = New StreamWriter(strRutaArchivo, True)
                writer.WriteLine($"{DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss")} - GUID: {strGuiaRegistro} - PARAMETROS: {strMensaje}")
            End Using
        Catch ex As Exception
            Console.WriteLine($"Error al escribir en el log: {ex.Message}")
        End Try

    End Sub



End Class