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

' Clase para retornar ambos valores de seguro
Public Class ResultadoPolizaSeguro
    Public Property SeguroVida As Decimal
    Public Property SeguroDanios As Decimal
End Class

' Clase auxiliar para validaciones de datos
Public Class ValidadorDatos
    ''' <summary>
    ''' Valida si una cadena puede convertirse a Decimal de forma segura
    ''' </summary>
    ''' <param name="valor">Valor a validar</param>
    ''' <param name="nombreCampo">Nombre del campo para mensajes de error</param>
    ''' <param name="esOpcional">Indica si el campo es opcional (puede estar vacío)</param>
    ''' <returns>El valor convertido o 0 si es opcional y está vacío</returns>
    Public Shared Function ValidarDecimal(valor As String, nombreCampo As String, Optional esOpcional As Boolean = True) As Decimal
        ' Si es opcional y está vacío, retornar 0
        If esOpcional AndAlso String.IsNullOrWhiteSpace(valor) Then
            Return 0
        End If
        
        ' Si no es opcional y está vacío, lanzar excepción
        If Not esOpcional AndAlso String.IsNullOrWhiteSpace(valor) Then
            Throw New ArgumentException($"El campo '{nombreCampo}' es obligatorio y no puede estar vacío.")
        End If
        
        ' Intentar convertir a decimal
        Dim resultado As Decimal
        If Decimal.TryParse(valor, resultado) Then
            ' Validar que no sea negativo
            If resultado < 0 Then
                Throw New ArgumentException($"El campo '{nombreCampo}' no puede ser negativo. Valor recibido: {valor}")
            End If
            Return resultado
        Else
            Throw New ArgumentException($"El campo '{nombreCampo}' debe ser un número válido. Valor recibido: '{valor}'")
        End If
    End Function
    
    ''' <summary>
    ''' Valida si una cadena puede convertirse a Integer de forma segura
    ''' </summary>
    ''' <param name="valor">Valor a validar</param>
    ''' <param name="nombreCampo">Nombre del campo para mensajes de error</param>
    ''' <param name="esOpcional">Indica si el campo es opcional (puede estar vacío)</param>
    ''' <returns>El valor convertido o 0 si es opcional y está vacío</returns>
    Public Shared Function ValidarInteger(valor As String, nombreCampo As String, Optional esOpcional As Boolean = True) As Integer
        ' Si es opcional y está vacío, retornar 0
        If esOpcional AndAlso String.IsNullOrWhiteSpace(valor) Then
            Return 0
        End If
        
        ' Si no es opcional y está vacío, lanzar excepción
        If Not esOpcional AndAlso String.IsNullOrWhiteSpace(valor) Then
            Throw New ArgumentException($"El campo '{nombreCampo}' es obligatorio y no puede estar vacío.")
        End If
        
        ' Intentar convertir a integer
        Dim resultado As Integer
        If Integer.TryParse(valor, resultado) Then
            ' Validar que no sea negativo
            If resultado < 0 Then
                Throw New ArgumentException($"El campo '{nombreCampo}' no puede ser negativo. Valor recibido: {valor}")
            End If
            Return resultado
        Else
            Throw New ArgumentException($"El campo '{nombreCampo}' debe ser un número entero válido. Valor recibido: '{valor}'")
        End If
    End Function
    
    ''' <summary>
    ''' Valida campos obligatorios de texto
    ''' </summary>
    ''' <param name="valor">Valor a validar</param>
    ''' <param name="nombreCampo">Nombre del campo para mensajes de error</param>
    ''' <param name="esOpcional">Indica si el campo es opcional</param>
    ''' <returns>El valor validado</returns>
    Public Shared Function ValidarTexto(valor As String, nombreCampo As String, Optional esOpcional As Boolean = True) As String
        ' Si es opcional, permitir valores vacíos
        If esOpcional Then
            Return If(valor, "")
        End If
        
        ' Si no es opcional y está vacío, lanzar excepción
        If String.IsNullOrWhiteSpace(valor) Then
            Throw New ArgumentException($"El campo '{nombreCampo}' es obligatorio y no puede estar vacío.")
        End If
        
        Return valor.Trim()
    End Function
    
    ''' <summary>
    ''' Valida que un valor decimal esté dentro de un rango específico
    ''' </summary>
    ''' <param name="valor">Valor a validar</param>
    ''' <param name="nombreCampo">Nombre del campo</param>
    ''' <param name="valorMinimo">Valor mínimo permitido</param>
    ''' <param name="valorMaximo">Valor máximo permitido</param>
    Public Shared Sub ValidarRango(valor As Decimal, nombreCampo As String, valorMinimo As Decimal, valorMaximo As Decimal)
        If valor < valorMinimo OrElse valor > valorMaximo Then
            Throw New ArgumentException($"El campo '{nombreCampo}' debe estar entre {valorMinimo} y {valorMaximo}. Valor recibido: {valor}")
        End If
    End Sub
End Class

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
                                 strTipoDeuda1 As String, strSaldoDeuda1 As String, strLimiteTarjeta1 As String, strCuotaDeuda1 As String,
                                 strTipoDeuda2 As String, strSaldoDeuda2 As String, strLimiteTarjeta2 As String, strCuotaDeuda2 As String,
                                 strTipoDeuda3 As String, strSaldoDeuda3 As String, strLimiteTarjeta3 As String, strCuotaDeuda3 As String,
                                 strTipoDeuda4 As String, strSaldoDeuda4 As String, strLimiteTarjeta4 As String, strCuotaDeuda4 As String,
                                 strTipoDeuda5 As String, strSaldoDeuda5 As String, strLimiteTarjeta5 As String, strCuotaDeuda5 As String,
                                 strTipoDeuda6 As String, strSaldoDeuda6 As String, strLimiteTarjeta6 As String, strCuotaDeuda6 As String,
                                 strMes1 As String, strMes2 As String, strMes3 As String, strTerreno As String, strConstrucciones As String,
                                 strScorePredictivo As String, strClasificacionSIB As String, strConteoCCR As String, strPlazoMeses As String,
                                 strBonificacionActividadEconomica As String, strIgssActividadEconomica As String, strIsrActividadEconomica As String,
                                 strBonificacionActividadEconomica2 As String, strIgssActividadEconomica2 As String, strIsrActividadEconomica2 As String,
                                 strComisionesActividadEconomica As String, strComisionesActividadEconomica2 As String, strTasaInteres As String,
                                 strTipoCuota As String, strTasaReferenciaLIP As String) As XmlDocument

        Dim xmlRespuesta As New XmlDocument
        Dim xDoc As XDocument
        Dim dcmTasaInteres As Decimal
        Dim intPlazoMeses As Integer
        Dim dcmBonificacionActividadEconomica As Decimal
        Dim dcmIgssActividadEconomica As Decimal
        Dim dcmIsrActividadEconomica As Decimal
        Dim dcmBonificacionActividadEconomica2 As Decimal
        Dim dcmIgssActividadEconomica2 As Decimal
        Dim dcmIsrActividadEconomica2 As Decimal
        Dim dcmRDG As Decimal
        Dim dcmValorEGarantiaHipotecaria As Decimal
        Dim dcmOtrasDeducciones As Decimal
        Dim dcmTotalS As Decimal
        Dim dcmTotalSalario1 As Decimal
        Dim dcmTotalSalario2 As Decimal
        Dim dcmTotalSalarios As Decimal
        Dim dcmCuota As Decimal
        Dim dcmTotalCuotasDirectas As Decimal
        Dim dcmVidaTotal As Decimal
        Dim dcmVidaMes As Decimal
        Dim dcmDanios As Decimal
        Dim dcmEndeudamientoDirecto As Decimal
        Dim dcmEndeudamientoInterno As Decimal
        Dim dcmEndeudamientoExterno As Decimal
        Dim intRCI1 As Integer
        Dim intRCI2 As Integer
        Dim intRCI1al48 As Integer = 0
        Dim intRCI49al84 As Integer = 0
        Dim intRCI85mas As Integer = 0
        Dim intRCI1al48Total As Integer = 0
        Dim intRCI49al84Total As Integer = 0
        Dim intRCI85masTotal As Integer = 0
        Dim dcmLIPFHA1al48 As Decimal = 0
        Dim dcmLIPFHA1al48Total As Decimal = 0
        Dim dcmLIPFHA49al84 As Decimal = 0
        Dim dcmLIPFHA49al84Total As Decimal = 0
        Dim dcmLIPFHA85mas As Decimal = 0
        Dim dcmLIPFHA85masTotal As Decimal = 0
        Dim dcmLIPDIRECTO1al48 As Decimal = 0
        Dim dcmLIPDIRECTO1al48Total As Decimal = 0
        Dim dcmLIPDIRECTO49al84 As Decimal = 0
        Dim dcmLIPDIRECTO49al84Total As Decimal = 0
        Dim dcmLIPDIRECTO85mas As Decimal = 0
        Dim dcmLIPDIRECTO85masTotal As Decimal = 0
        Dim dcmValorGarantia As Decimal
        Dim dcmLIPFHA1 As Decimal
        Dim dcmLIPFHA2 As Decimal
        Dim dcmLIPDIRECTO1 As Decimal
        Dim dcmLIPDIRECTO2 As Decimal
        Dim strValorDiferente As String
        Dim strRCIDiferente As String
        Dim strGuiaRegistro As String
        Dim strParametros As Dictionary(Of String, Object)
        Dim strParametrosDetalle As String
        Dim strRutaArchivo As String
        Dim strFechaHora As String
        Dim strDetalle As String
        Dim strPlazoAdvertencia As String = ""
        Dim dcmBonificacionAE As Decimal
        Dim dcmIgssAE As Decimal
        Dim dcmIsrAE As Decimal
        Dim dcmBonificacionAE2 As Decimal
        Dim dcmIgssAE2 As Decimal
        Dim dcmIsrAE2 As Decimal
        Dim dcmComisionesActividadEconomica As Decimal
        Dim dcmComisionesActividadEconomica2 As Decimal

        Try
            ' ===========================================
            ' VALIDACIONES DE DATOS DE ENTRADA
            ' ===========================================

            ' Validar campos obligatorios principales
            If dcmMontoSolicitado <= 0 Then
                Throw New ArgumentException("El monto solicitado debe ser mayor a 0.")
            End If

            strDestinoCredito = ValidadorDatos.ValidarTexto(strDestinoCredito, "Destino del Crédito", False)
            strTipoGarantia = ValidadorDatos.ValidarTexto(strTipoGarantia, "Tipo de Garantía", False)
            strActividadEconomica = ValidadorDatos.ValidarTexto(strActividadEconomica, "Actividad Económica 1", True)

            ' Validar campos numéricos opcionales
            Dim dcmIngresoConstanciaValidado As Decimal = dcmIngresoConstancia
            If dcmIngresoConstanciaValidado < 0 Then
                Throw New ArgumentException("Los ingresos según constancia no pueden ser negativos.")
            End If

            strActividadEconomica2 = ValidadorDatos.ValidarTexto(strActividadEconomica2, "Actividad Económica 2", True)

            ' Validar todos los campos de deudas (6 deudas posibles)
            Dim deudasValidadas As New List(Of Dictionary(Of String, Object))
            For i As Integer = 1 To 6
                Dim tipoDeuda As String = ""
                Dim saldoDeuda As Decimal = 0
                Dim limiteTarjeta As Decimal = 0
                Dim cuotaDeuda As Decimal = Nothing

                Select Case i
                    Case 1
                        tipoDeuda = ValidadorDatos.ValidarTexto(strTipoDeuda1, $"Tipo Deuda {i}", True)
                        saldoDeuda = ValidadorDatos.ValidarDecimal(strSaldoDeuda1, $"Saldo Deuda {i}", True)
                        limiteTarjeta = ValidadorDatos.ValidarDecimal(strLimiteTarjeta1, $"Límite Tarjeta {i}", True)
                        cuotaDeuda = ValidadorDatos.ValidarDecimal(strCuotaDeuda1, $"Cuota Deuda {i}", True)
                    Case 2
                        tipoDeuda = ValidadorDatos.ValidarTexto(strTipoDeuda2, $"Tipo Deuda {i}", True)
                        saldoDeuda = ValidadorDatos.ValidarDecimal(strSaldoDeuda2, $"Saldo Deuda {i}", True)
                        limiteTarjeta = ValidadorDatos.ValidarDecimal(strLimiteTarjeta2, $"Límite Tarjeta {i}", True)
                        cuotaDeuda = ValidadorDatos.ValidarDecimal(strCuotaDeuda2, $"Cuota Deuda {i}", True)
                    Case 3
                        tipoDeuda = ValidadorDatos.ValidarTexto(strTipoDeuda3, $"Tipo Deuda {i}", True)
                        saldoDeuda = ValidadorDatos.ValidarDecimal(strSaldoDeuda3, $"Saldo Deuda {i}", True)
                        limiteTarjeta = ValidadorDatos.ValidarDecimal(strLimiteTarjeta3, $"Límite Tarjeta {i}", True)
                        cuotaDeuda = ValidadorDatos.ValidarDecimal(strCuotaDeuda3, $"Cuota Deuda {i}", True)
                    Case 4
                        tipoDeuda = ValidadorDatos.ValidarTexto(strTipoDeuda4, $"Tipo Deuda {i}", True)
                        saldoDeuda = ValidadorDatos.ValidarDecimal(strSaldoDeuda4, $"Saldo Deuda {i}", True)
                        limiteTarjeta = ValidadorDatos.ValidarDecimal(strLimiteTarjeta4, $"Límite Tarjeta {i}", True)
                        cuotaDeuda = ValidadorDatos.ValidarDecimal(strCuotaDeuda4, $"Cuota Deuda {i}", True)
                    Case 5
                        tipoDeuda = ValidadorDatos.ValidarTexto(strTipoDeuda5, $"Tipo Deuda {i}", True)
                        saldoDeuda = ValidadorDatos.ValidarDecimal(strSaldoDeuda5, $"Saldo Deuda {i}", True)
                        limiteTarjeta = ValidadorDatos.ValidarDecimal(strLimiteTarjeta5, $"Límite Tarjeta {i}", True)
                        cuotaDeuda = ValidadorDatos.ValidarDecimal(strCuotaDeuda5, $"Cuota Deuda {i}", True)
                    Case 6
                        tipoDeuda = ValidadorDatos.ValidarTexto(strTipoDeuda6, $"Tipo Deuda {i}", True)
                        saldoDeuda = ValidadorDatos.ValidarDecimal(strSaldoDeuda6, $"Saldo Deuda {i}", True)
                        limiteTarjeta = ValidadorDatos.ValidarDecimal(strLimiteTarjeta6, $"Límite Tarjeta {i}", True)
                        cuotaDeuda = ValidadorDatos.ValidarDecimal(strCuotaDeuda6, $"Cuota Deuda {i}", True)
                End Select

                ' Si hay tipo de deuda, validar que tenga datos relacionados
                If Not String.IsNullOrEmpty(tipoDeuda) Then
                    If saldoDeuda = 0 AndAlso cuotaDeuda = 0 Then
                        Throw New ArgumentException($"Para la deuda {i} de tipo '{tipoDeuda}', debe proporcionar al menos el saldo de la deuda o la cuota mensual.")
                    End If
                End If

                deudasValidadas.Add(New Dictionary(Of String, Object) From {
                    {"Tipo", tipoDeuda},
                    {"Saldo", saldoDeuda},
                    {"Limite", limiteTarjeta},
                    {"Cuota", cuotaDeuda}
                })
            Next

            ' Validar campos de ingresos verificados
            Dim mes1 As Decimal = ValidadorDatos.ValidarDecimal(strMes1, "Mes 1", True)
            Dim mes2 As Decimal = ValidadorDatos.ValidarDecimal(strMes2, "Mes 2", True)
            Dim mes3 As Decimal = ValidadorDatos.ValidarDecimal(strMes3, "Mes 3", True)

            ' Validar campos de garantía
            Dim terreno As Decimal = ValidadorDatos.ValidarDecimal(strTerreno, "Terreno", True)
            Dim construcciones As Decimal = ValidadorDatos.ValidarDecimal(strConstrucciones, "Construcciones", True)

            ' Validar campos de scoring
            strScorePredictivo = ValidadorDatos.ValidarTexto(strScorePredictivo, "Score Predictivo", True)
            strClasificacionSIB = ValidadorDatos.ValidarTexto(strClasificacionSIB, "Clasificación SIB", True)
            Dim conteoCCR As Integer = ValidadorDatos.ValidarInteger(strConteoCCR, "Conteo CCR", True)

            ' Validar plazo de meses
            Dim plazoMesesValidado As Integer = ValidadorDatos.ValidarInteger(strPlazoMeses, "Plazo de Meses", True)

            ' Validar campos de bonificaciones y descuentos
            Dim bonificacionAE As Decimal = ValidadorDatos.ValidarDecimal(strBonificacionActividadEconomica, "Bonificación Actividad Económica", True)
            Dim igssAE As Decimal = ValidadorDatos.ValidarDecimal(strIgssActividadEconomica, "IGSS Actividad Económica", True)
            Dim isrAE As Decimal = ValidadorDatos.ValidarDecimal(strIsrActividadEconomica, "ISR Actividad Económica", True)

            Dim bonificacionAE2 As Decimal = ValidadorDatos.ValidarDecimal(strBonificacionActividadEconomica2, "Bonificación Actividad Económica 2", True)
            Dim igssAE2 As Decimal = ValidadorDatos.ValidarDecimal(strIgssActividadEconomica2, "IGSS Actividad Económica 2", True)
            Dim isrAE2 As Decimal = ValidadorDatos.ValidarDecimal(strIsrActividadEconomica2, "ISR Actividad Económica 2", True)

            Dim comisionesAE As Decimal = ValidadorDatos.ValidarDecimal(strComisionesActividadEconomica, "Comisiones Actividad Económica", True)
            Dim comisionesAE2 As Decimal = ValidadorDatos.ValidarDecimal(strComisionesActividadEconomica2, "Comisiones Actividad Económica 2", True)

            ' Validar tasa de interés
            Dim tasaInteresValidada As Decimal = ValidadorDatos.ValidarDecimal(strTasaInteres, "Tasa de Interés", True)

            ' Validar tipo de cuota
            strTipoCuota = ValidadorDatos.ValidarTexto(strTipoCuota, "Tipo de Cuota", True)
            If Not String.IsNullOrEmpty(strTipoCuota) AndAlso strTipoCuota.ToLower() <> "saldos" AndAlso strTipoCuota.ToLower() <> "nivelada" Then
                Throw New ArgumentException("El tipo de cuota debe ser 'saldos' o 'nivelada'.")
            End If

            ' Validar tasa de referencia LIP (solo para LIP FHA y LIP DIRECTO)
            Dim tasaReferenciaLIP As Decimal = ValidadorDatos.ValidarDecimal(strTasaReferenciaLIP, "Tasa de Referencia LIP", True)

            ' Validar otros descuentos
            Dim auxilioPostumo As Decimal = ValidadorDatos.ValidarDecimal(strAuxilioPostumo, "Auxilio Póstumo", True)
            Dim montepio As Decimal = ValidadorDatos.ValidarDecimal(strMontepio, "Montepio", True)
            Dim segurosDescuento As Decimal = ValidadorDatos.ValidarDecimal(strSeguros, "Seguros", True)
            Dim descuentoConstancia As Decimal = ValidadorDatos.ValidarDecimal(strDescuentoConstancia, "Descuento Constancia", True)

            ' Validar ingresos constancia 2
            Dim ingresoConstancia2Validado As Decimal = ValidadorDatos.ValidarDecimal(strIngresoConstancia2, "Ingresos Constancia 2", True)
            If ingresoConstancia2Validado < 0 Then
                Throw New ArgumentException("Los ingresos según constancia 2 no pueden ser negativos.")
            End If

            ' ===========================================
            ' FIN DE VALIDACIONES - CONTINUAR CON LÓGICA ORIGINAL
            ' ===========================================

            strParametros = New Dictionary(Of String, Object) From {
                {"Monto solicitado", dcmMontoSolicitado},
                {"Destino del Crédito", strDestinoCredito}, {"Tipo de Garantía", strTipoGarantia}, {"Actividad Económica 1", strActividadEconomica},
                {"Ingresos según Constancia", dcmIngresoConstanciaValidado}, {"Actividad Económica 2", strActividadEconomica2}, {"Ingresos según Constancia 2", ingresoConstancia2Validado},
                {"Auxilio Postumo", auxilioPostumo}, {"Montepio", montepio}, {"Seguros", segurosDescuento}, {"Otros Descuentos de Constancia", descuentoConstancia},
                {"Tipo Deuda1", deudasValidadas(0)("Tipo")}, {"Capital Original1", deudasValidadas(0)("Saldo")}, {"Limite Tarjeta1", deudasValidadas(0)("Limite")}, {"Cuota Deuda1", deudasValidadas(0)("Cuota")},
                {"Tipo Deuda2", deudasValidadas(1)("Tipo")}, {"Capital Original2", deudasValidadas(1)("Saldo")}, {"Limite Tarjeta2", deudasValidadas(1)("Limite")}, {"Cuota Deuda2", deudasValidadas(1)("Cuota")},
                {"Tipo Deuda3", deudasValidadas(2)("Tipo")}, {"Capital Original3", deudasValidadas(2)("Saldo")}, {"Limite Tarjeta3", deudasValidadas(2)("Limite")}, {"Cuota Deuda3", deudasValidadas(2)("Cuota")},
                {"Tipo Deuda4", deudasValidadas(3)("Tipo")}, {"Capital Original4", deudasValidadas(3)("Saldo")}, {"Limite Tarjeta4", deudasValidadas(3)("Limite")}, {"Cuota Deuda4", deudasValidadas(3)("Cuota")},
                {"Tipo Deuda5", deudasValidadas(4)("Tipo")}, {"Capital Original5", deudasValidadas(4)("Saldo")}, {"Limite Tarjeta5", deudasValidadas(4)("Limite")}, {"Cuota Deuda5", deudasValidadas(4)("Cuota")},
                {"Tipo Deuda6", deudasValidadas(5)("Tipo")}, {"Capital Original6", deudasValidadas(5)("Saldo")}, {"Limite Tarjeta6", deudasValidadas(5)("Limite")}, {"Cuota Deuda6", deudasValidadas(5)("Cuota")},
                {"Mes 1", mes1}, {"Mes 2", mes2}, {"Mes 3", mes3}, {"Terreno", terreno}, {"Construcciones", construcciones},
                {"Score predictivo", strScorePredictivo}, {"Clasificación SIB", strClasificacionSIB}, {"Conteo de CCR", conteoCCR},
                {"No de cuotas", plazoMesesValidado}, {"Bonificación AE", bonificacionAE}, {"IGSS AE", igssAE},
                {"ISR AE", isrAE}, {"Bonificación AE 2", bonificacionAE2}, {"IGSS AE 2", igssAE2},
                {"ISR AE 2", isrAE2}, {"Comisiones AE", comisionesAE}, {"Comisiones AE 2", comisionesAE2},
                {"Tasa de Referencia LIP", tasaReferenciaLIP}
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

            ' Usar valores validados para tasa de interés
            If Not String.IsNullOrEmpty(strDestinoCredito) Then
                If tasaInteresValidada > 0 Then
                    dcmTasaInteres = tasaInteresValidada
                Else
                    dcmTasaInteres = ObtenerTasaXDestinoCredito(strDestinoCredito)
                End If
            End If

            ' Usar valores validados para plazo de meses
            If Not String.IsNullOrEmpty(strTipoGarantia) Then
                If plazoMesesValidado > 0 Then
                    intPlazoMeses = plazoMesesValidado
                Else
                    intPlazoMeses = ObtenerPlazoMesesXTipoGarantia(strTipoGarantia, "")
                End If

                ' Guardamos el plazo máximo para poder mostrarlo en el mensaje
                Dim intPlazoMaximo As Integer = ObtenerPlazoMesesXTipoGarantia(strTipoGarantia, "")
                If intPlazoMeses > intPlazoMaximo Then
                    strPlazoAdvertencia = "El plazo ingresado excede al plazo maximo para este tipo de prestamo: " & intPlazoMaximo
                End If
            End If

            ' Usar valores validados para bonificaciones y descuentos
            dcmBonificacionAE = bonificacionAE
            dcmIgssAE = igssAE
            dcmIsrAE = isrAE
            dcmBonificacionAE2 = bonificacionAE2
            dcmIgssAE2 = igssAE2
            dcmIsrAE2 = isrAE2

            ' Calculo de los ingresos de la actividad economica usando valores validados
            If Not String.IsNullOrEmpty(strActividadEconomica) And dcmIngresoConstanciaValidado > 0 Then

                ' Usar valores validados para ISR
                If isrAE > 0 Then
                    dcmIsrActividadEconomica = isrAE
                Else
                    dcmIsrActividadEconomica = (dcmIngresoConstanciaValidado * ObtenerIgss_Isr("Isr")) / 100
                End If

                ' Usar valores validados para otras deducciones
                dcmOtrasDeducciones = descuentoConstancia + auxilioPostumo + montepio + segurosDescuento

                If String.Equals(strActividadEconomica, "Relacion de dependencia", StringComparison.OrdinalIgnoreCase) Then

                    ' Usar valores validados para bonificación
                    If bonificacionAE > 0 Then
                        dcmBonificacionActividadEconomica = bonificacionAE
                    Else
                        dcmBonificacionActividadEconomica = 250
                    End If

                    ' Usar valores validados para IGSS
                    If igssAE > 0 Then
                        dcmIgssActividadEconomica = igssAE
                    Else
                        dcmIgssActividadEconomica = (dcmIngresoConstanciaValidado * ObtenerIgss_Isr("Igss")) / 100
                    End If

                    ' Usar valores validados para comisiones
                    If comisionesAE > 0 Then
                        dcmComisionesActividadEconomica = comisionesAE * 0.5
                    Else
                        dcmComisionesActividadEconomica = 0
                    End If

                    dcmTotalS = dcmIngresoConstanciaValidado + dcmBonificacionActividadEconomica - dcmIgssActividadEconomica - dcmIsrActividadEconomica - dcmOtrasDeducciones + dcmComisionesActividadEconomica
                    dcmTotalSalario1 = Math.Round((dcmTotalS * 14) / 12, 2)
                Else
                    dcmBonificacionActividadEconomica = 0
                    dcmIgssActividadEconomica = 0

                    ' Usar valores validados para comisiones
                    If comisionesAE > 0 Then
                        dcmComisionesActividadEconomica = comisionesAE * 0.5
                    Else
                        dcmComisionesActividadEconomica = 0
                    End If

                    dcmTotalS = dcmIngresoConstanciaValidado - dcmIsrActividadEconomica + dcmBonificacionActividadEconomica - dcmIgssActividadEconomica - dcmOtrasDeducciones + dcmComisionesActividadEconomica
                    dcmTotalSalario1 = Math.Round(dcmTotalS * 0.6, 2)
                End If
            End If

            ' Calculo de los ingresos de la segunda actividad economica usando valores validados
            If Not String.IsNullOrEmpty(strActividadEconomica2) And ingresoConstancia2Validado > 0 Then

                ' Usar valores validados para ISR
                If isrAE2 > 0 Then
                    dcmIsrActividadEconomica2 = isrAE2
                Else
                    dcmIsrActividadEconomica2 = (ingresoConstancia2Validado * ObtenerIgss_Isr("Isr")) / 100
                End If

                If String.Equals(strActividadEconomica2, "Relacion de dependencia", StringComparison.OrdinalIgnoreCase) Then
                    ' Usar valores validados para bonificación
                    If bonificacionAE2 > 0 Then
                        dcmBonificacionActividadEconomica2 = bonificacionAE2
                    Else
                        dcmBonificacionActividadEconomica2 = 250
                    End If

                    ' Usar valores validados para IGSS
                    If igssAE2 > 0 Then
                        dcmIgssActividadEconomica2 = igssAE2
                    Else
                        dcmIgssActividadEconomica2 = (ingresoConstancia2Validado * ObtenerIgss_Isr("Igss")) / 100
                    End If

                    ' Usar valores validados para comisiones
                    If comisionesAE2 > 0 Then
                        dcmComisionesActividadEconomica2 = comisionesAE2 * 0.5
                    Else
                        dcmComisionesActividadEconomica2 = 0
                    End If

                    ' Calculo del total de salarios
                    dcmTotalS = ingresoConstancia2Validado + dcmBonificacionActividadEconomica2 - dcmIgssActividadEconomica2 - dcmIsrActividadEconomica2 + dcmComisionesActividadEconomica2
                    dcmTotalSalario2 = Math.Round((dcmTotalS * 14) / 12, 2)
                Else
                    dcmBonificacionActividadEconomica2 = 0
                    dcmIgssActividadEconomica2 = 0

                    ' Usar valores validados para comisiones
                    If comisionesAE2 > 0 Then
                        dcmComisionesActividadEconomica2 = comisionesAE2 * 0.5
                    Else
                        dcmComisionesActividadEconomica2 = 0
                    End If

                    ' Calculo del total de salarios
                    dcmTotalS = ingresoConstancia2Validado + dcmBonificacionActividadEconomica2 - dcmIgssActividadEconomica2 - dcmIsrActividadEconomica2 + dcmComisionesActividadEconomica2
                    dcmTotalSalario2 = Math.Round(dcmTotalS * 0.6, 2)
                End If
            End If

            ' Cálculo de endeudamiento usando valores validados
            dcmEndeudamientoDirecto = 0
            dcmEndeudamientoInterno = 0
            dcmEndeudamientoExterno = 0

            ' Procesar cada deuda validada
            For i As Integer = 0 To 5
                Dim deuda As Dictionary(Of String, Object) = deudasValidadas(i)
                Dim tipoDeuda As String = deuda("Tipo").ToString()
                Dim saldoDeuda As Decimal = Convert.ToDecimal(deuda("Saldo"))
                Dim limiteTarjeta As Decimal = Convert.ToDecimal(deuda("Limite"))
                Dim cuotaDeuda As Decimal = Convert.ToDecimal(deuda("Cuota"))

                ' Solo procesar si hay tipo de deuda
                If Not String.IsNullOrEmpty(tipoDeuda) Then
                    Dim endeudamiento As Decimal = SumaEndeudamientoDirecto(tipoDeuda, saldoDeuda.ToString(), limiteTarjeta.ToString(), cuotaDeuda.ToString())

                    If tipoDeuda = "Prestamo Fiduciario Indirecto Bantrab" Or tipoDeuda = "Tarjeta de Credito Interna" Then
                        dcmEndeudamientoInterno += endeudamiento
                    Else
                        dcmEndeudamientoExterno += endeudamiento
                    End If
                    dcmEndeudamientoDirecto += endeudamiento
                End If
            Next

            dcmTotalSalarios = dcmTotalSalario1 + dcmTotalSalario2

            ' Inicializar la variable seguros antes del bloque If/Else para LIP FHA y LIP DIRECTO
            Dim seguros As New ResultadoPolizaSeguro()

            ' Cálculo de cuota sin seguros según tipo de cuota
            Dim cuotaSinSeguro As Decimal
            If strTipoCuota = "saldos" Then
                cuotaSinSeguro = Math.Round((dcmMontoSolicitado / intPlazoMeses) + (dcmMontoSolicitado * (dcmTasaInteres / 100) / 360 * 30), 2)
            Else
                cuotaSinSeguro = Math.Round(CalcularPago(dcmTasaInteres / 100 / 12, intPlazoMeses, dcmMontoSolicitado), 2)
            End If

            If String.Equals(strDestinoCredito, "LIP FHA", StringComparison.OrdinalIgnoreCase) Then
                ' Solo realizar el cálculo LIP FHA si se proporciona la tasa de referencia
                If tasaReferenciaLIP > 0 Then
                    Dim tasaConjunta = dcmTasaInteres
                    Dim seguroHipoteca = 1.26
                    Dim tasaBanco = tasaConjunta - seguroHipoteca
                    Dim dcmValorMatricular As Decimal
                    Dim dcmValorConstruccion As Decimal

                    cuotaSinSeguro = Math.Round(CalcularPago((tasaConjunta / 100) / 12, intPlazoMeses, dcmMontoSolicitado), 2)
                    Dim S_H = Math.Round(dcmMontoSolicitado * seguroHipoteca / 100 / 12, 2) 'Columna G11
                    Dim IUSI = 0.009 * dcmValorMatricular / 12
                    Dim seguroHipotecaValor = 0.28 / 100 * dcmValorConstruccion / 12

                    ' Variables para los diferentes RCI según plazo ya están declaradas arriba

                    ' Cálculo para cuotas 1-48 (40% subsidio)
                    If intPlazoMeses >= 1 Then
                        Dim plazo1al48 = Math.Min(intPlazoMeses, 48)
                        Dim tasa1al48 = (tasaBanco / 100) - (40 / 100) * (tasaReferenciaLIP / 100)
                        Dim interesCliente1al48 = Math.Round(dcmMontoSolicitado * tasa1al48 / 12, 2)
                        Dim interesSubsidio1al48 = Math.Round(dcmMontoSolicitado * tasaBanco / 100 / 12, 2) - interesCliente1al48
                        Dim capital1al48 = cuotaSinSeguro - interesCliente1al48 - S_H - interesSubsidio1al48
                        Dim pagoTotal1al48 = capital1al48 + interesCliente1al48 + interesSubsidio1al48 + S_H + IUSI + seguroHipotecaValor
                        Dim pagoCliente1al48 = pagoTotal1al48 - interesSubsidio1al48

                        dcmLIPFHA1al48 = pagoCliente1al48
                        dcmLIPFHA1al48Total = pagoTotal1al48
                        intRCI1al48 = ((dcmLIPFHA1al48 + dcmEndeudamientoDirecto) / dcmTotalSalarios) * 100
                        intRCI1al48Total = ((dcmLIPFHA1al48Total + dcmEndeudamientoDirecto) / dcmTotalSalarios) * 100
                    End If

                    ' Cálculo para cuotas 49-84 (30% subsidio)
                    If intPlazoMeses > 48 Then
                        Dim plazo49al84 = Math.Min(intPlazoMeses - 48, 36)
                        Dim tasa49al84 = (tasaBanco / 100) - (30 / 100) * (tasaReferenciaLIP / 100)
                        Dim interesCliente49al84 = Math.Round(dcmMontoSolicitado * tasa49al84 / 12, 2)
                        Dim interesSubsidio49al84 = Math.Round(dcmMontoSolicitado * tasaBanco / 100 / 12, 2) - interesCliente49al84
                        Dim capital49al84 = cuotaSinSeguro - interesCliente49al84 - S_H - interesSubsidio49al84
                        Dim pagoTotal49al84 = capital49al84 + interesCliente49al84 + interesSubsidio49al84 + S_H + IUSI + seguroHipotecaValor
                        Dim pagoCliente49al84 = pagoTotal49al84 - interesSubsidio49al84

                        dcmLIPFHA49al84 = pagoCliente49al84
                        dcmLIPFHA49al84Total = pagoTotal49al84
                        intRCI49al84 = ((dcmLIPFHA49al84 + dcmEndeudamientoDirecto) / dcmTotalSalarios) * 100
                        intRCI49al84Total = ((dcmLIPFHA49al84Total + dcmEndeudamientoDirecto) / dcmTotalSalarios) * 100
                    End If

                    ' Cálculo para cuotas 85+ (0% subsidio)
                    If intPlazoMeses > 84 Then
                        Dim tasa85mas = (tasaBanco / 100) - (0 / 100) * (tasaReferenciaLIP / 100)
                        Dim interesCliente85mas = Math.Round(dcmMontoSolicitado * tasa85mas / 12, 2)
                        Dim interesSubsidio85mas = Math.Round(dcmMontoSolicitado * tasaBanco / 100 / 12, 2) - interesCliente85mas
                        Dim capital85mas = cuotaSinSeguro - interesCliente85mas - S_H - interesSubsidio85mas
                        Dim pagoTotal85mas = capital85mas + interesCliente85mas + interesSubsidio85mas + S_H + IUSI + seguroHipotecaValor
                        Dim pagoCliente85mas = pagoTotal85mas - interesSubsidio85mas

                        dcmLIPFHA85mas = pagoCliente85mas
                        dcmLIPFHA85masTotal = pagoTotal85mas
                        intRCI85mas = ((dcmLIPFHA85mas + dcmEndeudamientoDirecto) / dcmTotalSalarios) * 100
                        intRCI85masTotal = ((dcmLIPFHA85masTotal + dcmEndeudamientoDirecto) / dcmTotalSalarios) * 100
                    End If

                    ' Usar el primer RCI calculado para compatibilidad
                    If intPlazoMeses <= 48 Then
                        intRCI1 = intRCI1al48
                        intRCI2 = intRCI1al48Total
                    ElseIf intPlazoMeses <= 84 Then
                        intRCI1 = intRCI49al84
                        intRCI2 = intRCI49al84Total
                    Else
                        intRCI1 = intRCI85mas
                        intRCI2 = intRCI85masTotal
                    End If

                    strDetalle = ""
                    If intRCI1 > 50 Then
                        strDetalle += "ALTO RCI INICIAL"
                    End If
                    If intRCI2 > 65 Then
                        strDetalle += " | ALTO RCI DESPUES DEL SUBSIDIO"
                    End If

                    strValorDiferente = dcmLIPFHA1 & " | " & dcmLIPFHA2
                    strRCIDiferente = Math.Round(intRCI1, MidpointRounding.AwayFromZero) & "% | " & Math.Round(intRCI2, MidpointRounding.AwayFromZero) & "%"
                Else
                    ' Si no se proporciona tasa de referencia, no hacer cálculo
                    strDetalle = "Para LIP FHA se requiere proporcionar la Tasa de Referencia LIP"
                    ' Inicializar valores en 0 para evitar errores
                    dcmCuota = 0
                    dcmTotalCuotasDirectas = 0
                    intRCI1 = 0
                    intRCI2 = 0
                End If

            ElseIf String.Equals(strDestinoCredito, "LIP DIRECTO", StringComparison.OrdinalIgnoreCase) Then
                ' Solo realizar el cálculo LIP DIRECTO si se proporciona la tasa de referencia
                If tasaReferenciaLIP > 0 Then
                    Dim tasaBanco = dcmTasaInteres
                    Dim dcmValorConstruccion = construcciones

                    cuotaSinSeguro = Math.Round(CalcularPago((tasaBanco / 100) / 12, intPlazoMeses, dcmMontoSolicitado), 2)
                    seguros = PolizadeSeguro(dcmMontoSolicitado, construcciones.ToString(), dcmTasaInteres, strTipoGarantia)

                    Dim seguroDirectoValor = seguros.SeguroDanios
                    Dim vida = seguros.SeguroVida

                    ' Variables para los diferentes RCI según plazo ya están declaradas arriba

                    ' Cálculo para cuotas 1-48 (40% subsidio)
                    If intPlazoMeses >= 1 Then
                        Dim plazo1al48 = Math.Min(intPlazoMeses, 48)
                        Dim tasa1al48 = (tasaBanco / 100) - (40 / 100) * (tasaReferenciaLIP / 100)
                        Dim interesCliente1al48 = Math.Round(dcmMontoSolicitado * tasa1al48 / 12, 2)
                        Dim interesSubsidio1al48 = Math.Round(dcmMontoSolicitado * tasaBanco / 100 / 12, 2) - interesCliente1al48
                        Dim capital1al48 = cuotaSinSeguro - interesCliente1al48 - interesSubsidio1al48
                        Dim pagoTotal1al48 = capital1al48 + interesCliente1al48 + interesSubsidio1al48 + vida + seguroDirectoValor
                        Dim pagoCliente1al48 = pagoTotal1al48 - interesSubsidio1al48

                        dcmLIPDIRECTO1al48 = pagoCliente1al48
                        dcmLIPDIRECTO1al48Total = pagoTotal1al48
                        intRCI1al48 = ((dcmLIPDIRECTO1al48 + dcmEndeudamientoDirecto) / dcmTotalSalarios) * 100
                        intRCI1al48Total = ((dcmLIPDIRECTO1al48Total + dcmEndeudamientoDirecto) / dcmTotalSalarios) * 100
                    End If

                    ' Cálculo para cuotas 49-84 (30% subsidio)
                    If intPlazoMeses > 48 Then
                        Dim plazo49al84 = Math.Min(intPlazoMeses - 48, 36)
                        Dim tasa49al84 = (tasaBanco / 100) - (30 / 100) * (tasaReferenciaLIP / 100)
                        Dim interesCliente49al84 = Math.Round(dcmMontoSolicitado * tasa49al84 / 12, 2)
                        Dim interesSubsidio49al84 = Math.Round(dcmMontoSolicitado * tasaBanco / 100 / 12, 2) - interesCliente49al84
                        Dim capital49al84 = cuotaSinSeguro - interesCliente49al84 - interesSubsidio49al84
                        Dim pagoTotal49al84 = capital49al84 + interesCliente49al84 + interesSubsidio49al84 + vida + seguroDirectoValor
                        Dim pagoCliente49al84 = pagoTotal49al84 - interesSubsidio49al84

                        dcmLIPDIRECTO49al84 = pagoCliente49al84
                        dcmLIPDIRECTO49al84Total = pagoTotal49al84
                        intRCI49al84 = ((dcmLIPDIRECTO49al84 + dcmEndeudamientoDirecto) / dcmTotalSalarios) * 100
                        intRCI49al84Total = ((dcmLIPDIRECTO49al84Total + dcmEndeudamientoDirecto) / dcmTotalSalarios) * 100
                    End If

                    ' Cálculo para cuotas 85+ (0% subsidio)
                    If intPlazoMeses > 84 Then
                        Dim tasa85mas = (tasaBanco / 100) - (0 / 100) * (tasaReferenciaLIP / 100)
                        Dim interesCliente85mas = Math.Round(dcmMontoSolicitado * tasa85mas / 12, 2)
                        Dim interesSubsidio85mas = Math.Round(dcmMontoSolicitado * tasaBanco / 100 / 12, 2) - interesCliente85mas
                        Dim capital85mas = cuotaSinSeguro - interesCliente85mas - interesSubsidio85mas
                        Dim pagoTotal85mas = capital85mas + interesCliente85mas + interesSubsidio85mas + vida + seguroDirectoValor
                        Dim pagoCliente85mas = pagoTotal85mas - interesSubsidio85mas

                        dcmLIPDIRECTO85mas = pagoCliente85mas
                        dcmLIPDIRECTO85masTotal = pagoTotal85mas
                        intRCI85mas = ((dcmLIPDIRECTO85mas + dcmEndeudamientoDirecto) / dcmTotalSalarios) * 100
                        intRCI85masTotal = ((dcmLIPDIRECTO85masTotal + dcmEndeudamientoDirecto) / dcmTotalSalarios) * 100
                    End If

                    ' Usar el primer RCI calculado para compatibilidad
                    If intPlazoMeses <= 48 Then
                        intRCI1 = intRCI1al48
                        intRCI2 = intRCI1al48Total
                    ElseIf intPlazoMeses <= 84 Then
                        intRCI1 = intRCI49al84
                        intRCI2 = intRCI49al84Total
                    Else
                        intRCI1 = intRCI85mas
                        intRCI2 = intRCI85masTotal
                    End If

                    strDetalle = ""
                    If intRCI1 > 50 Then
                        strDetalle += "ALTO RCI INICIAL"
                    End If
                    If intRCI2 > 65 Then
                        strDetalle += " | ALTO RCI DESPUES DEL SUBSIDIO"
                    End If

                    strValorDiferente = dcmLIPDIRECTO1 & " | " & dcmLIPDIRECTO2
                    strRCIDiferente = Math.Round(intRCI1, MidpointRounding.AwayFromZero) & "% | " & Math.Round(intRCI2, MidpointRounding.AwayFromZero) & "%"
                Else
                    ' Si no se proporciona tasa de referencia, no hacer cálculo
                    strDetalle = "Para LIP DIRECTO se requiere proporcionar la Tasa de Referencia LIP"
                    ' Inicializar valores en 0 para evitar errores
                    dcmCuota = 0
                    dcmTotalCuotasDirectas = 0
                    intRCI1 = 0
                    intRCI2 = 0
                End If

            Else
                seguros = PolizadeSeguro(dcmMontoSolicitado, construcciones.ToString(), dcmTasaInteres, strTipoGarantia)
                dcmCuota = Math.Round(cuotaSinSeguro, 2)
                dcmTotalCuotasDirectas = Math.Round(dcmCuota + dcmEndeudamientoDirecto + seguros.SeguroVida + seguros.SeguroDanios, 2)

                intRCI1 = (dcmTotalCuotasDirectas / dcmTotalSalarios) * 100
                strDetalle = ""

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

            ' Validación de ingresos verificados usando valores validados
            Dim dcmIngresosVerificados As Decimal = mes1 + mes2 + mes3

            Dim dcmPromedio = dcmIngresosVerificados / 3
            Dim dcmValidacionIngresos = Math.Round((dcmPromedio / dcmTotalSalarios) * 100, MidpointRounding.AwayFromZero)

            If dcmValidacionIngresos < 75 Then
                strDetalle += " | Los ingresos no se comprueban en estados de cuenta"
            End If

            ' Validación de garantía usando valores validados
            dcmValorGarantia = construcciones

            If dcmValorGarantia <> 0 Then
                dcmRDG = Math.Round(dcmMontoSolicitado / dcmValorGarantia * 100, MidpointRounding.AwayFromZero)
            End If

            ' Cálculo del valor de garantía hipotecaria según tipo
            If String.Equals(strTipoGarantia, "Terreno", StringComparison.OrdinalIgnoreCase) Then
                dcmValorEGarantiaHipotecaria = Math.Round(dcmMontoSolicitado * 100 / 60, 2)
            ElseIf String.Equals(strTipoGarantia, "Usada", StringComparison.OrdinalIgnoreCase) Then
                dcmValorEGarantiaHipotecaria = Math.Round(dcmMontoSolicitado * 100 / 75, 2)
            Else
                dcmValorEGarantiaHipotecaria = Math.Round(dcmMontoSolicitado * 100 / 80, 2)
            End If

            ' Validación de hipoteca usando valores validados
            Dim dcmHipoteca As Decimal = 0
            dcmValorGarantia = terreno + construcciones
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
                                   <bonificacionActividadEconomica>{dcmBonificacionActividadEconomica}</bonificacionActividadEconomica>
                                   <igssActividadEconomica>{Math.Round(dcmIgssActividadEconomica, 2)}</igssActividadEconomica>
                                   <isrActividadEconomica>{Math.Round(dcmIsrActividadEconomica, 2)}</isrActividadEconomica>
                                   <bonificacionActividadEconomica2>{dcmBonificacionActividadEconomica2}</bonificacionActividadEconomica2>
                                   <igssActividadEconomica2>{Math.Round(dcmIgssActividadEconomica2, 2)}</igssActividadEconomica2>
                                   <isrActividadEconomica2>{Math.Round(dcmIsrActividadEconomica2, 2)}</isrActividadEconomica2>
                                   <cuota>{dcmCuota}</cuota>
                                   <totalCuentasDirectas>{dcmTotalCuotasDirectas}</totalCuentasDirectas>
                                   <valorGarantiaHipotecaria>{dcmValorEGarantiaHipotecaria}</valorGarantiaHipotecaria>
                                   <rdg>{dcmRDG}%</rdg>
                                   <valorDiferente>{strValorDiferente}</valorDiferente>
                                   <rciDiferente>{strRCIDiferente}</rciDiferente>
                                   <rci1al48>{Math.Round(intRCI1al48, MidpointRounding.AwayFromZero)}%</rci1al48>
                                   <rci1al48Total>{Math.Round(intRCI1al48Total, MidpointRounding.AwayFromZero)}%</rci1al48Total>
                                   <rci49al84>{Math.Round(intRCI49al84, MidpointRounding.AwayFromZero)}%</rci49al84>
                                   <rci49al84Total>{Math.Round(intRCI49al84Total, MidpointRounding.AwayFromZero)}%</rci49al84Total>
                                   <rci85mas>{Math.Round(intRCI85mas, MidpointRounding.AwayFromZero)}%</rci85mas>
                                   <rci85masTotal>{Math.Round(intRCI85masTotal, MidpointRounding.AwayFromZero)}%</rci85masTotal>
                                   <cuota1al48Cliente>{If(intPlazoMeses >= 1, Math.Round(If(String.Equals(strDestinoCredito, "LIP FHA", StringComparison.OrdinalIgnoreCase), dcmLIPFHA1al48, dcmLIPDIRECTO1al48), 2), "")}</cuota1al48Cliente>
                                   <cuota1al48Total>{If(intPlazoMeses >= 1, Math.Round(If(String.Equals(strDestinoCredito, "LIP FHA", StringComparison.OrdinalIgnoreCase), dcmLIPFHA1al48Total, dcmLIPDIRECTO1al48Total), 2), "")}</cuota1al48Total>
                                   <cuota49al84Cliente>{If(intPlazoMeses > 48, Math.Round(If(String.Equals(strDestinoCredito, "LIP FHA", StringComparison.OrdinalIgnoreCase), dcmLIPFHA49al84, dcmLIPDIRECTO49al84), 2), "")}</cuota49al84Cliente>
                                   <cuota49al84Total>{If(intPlazoMeses > 48, Math.Round(If(String.Equals(strDestinoCredito, "LIP FHA", StringComparison.OrdinalIgnoreCase), dcmLIPFHA49al84Total, dcmLIPDIRECTO49al84Total), 2), "")}</cuota49al84Total>
                                   <cuota85masCliente>{If(intPlazoMeses > 84, Math.Round(If(String.Equals(strDestinoCredito, "LIP FHA", StringComparison.OrdinalIgnoreCase), dcmLIPFHA85mas, dcmLIPDIRECTO85mas), 2), "")}</cuota85masCliente>
                                   <cuota85masTotal>{If(intPlazoMeses > 84, Math.Round(If(String.Equals(strDestinoCredito, "LIP FHA", StringComparison.OrdinalIgnoreCase), dcmLIPFHA85masTotal, dcmLIPDIRECTO85masTotal), 2), "")}</cuota85masTotal>
                                   <trfLip></trfLip>
                                   <detalle>{strDetalle}</detalle>
                                   <plazoAdvertencia>{strPlazoAdvertencia}</plazoAdvertencia>
                                   <endeudamientoInterno>{Math.Round(dcmEndeudamientoInterno, 2)}</endeudamientoInterno>
                                   <endeudamientoExterno>{Math.Round(dcmEndeudamientoExterno, 2)}</endeudamientoExterno>
                                   <endeudamiento1>{If(String.IsNullOrEmpty(deudasValidadas(0)("Tipo").ToString()), "", Math.Round(SumaEndeudamientoDirecto(deudasValidadas(0)("Tipo").ToString(), deudasValidadas(0)("Saldo").ToString(), deudasValidadas(0)("Limite").ToString(), deudasValidadas(0)("Cuota").ToString()), 2).ToString())}</endeudamiento1>
                                   <endeudamiento2>{If(String.IsNullOrEmpty(deudasValidadas(1)("Tipo").ToString()), "", Math.Round(SumaEndeudamientoDirecto(deudasValidadas(1)("Tipo").ToString(), deudasValidadas(1)("Saldo").ToString(), deudasValidadas(1)("Limite").ToString(), deudasValidadas(1)("Cuota").ToString()), 2).ToString())}</endeudamiento2>
                                   <endeudamiento3>{If(String.IsNullOrEmpty(deudasValidadas(2)("Tipo").ToString()), "", Math.Round(SumaEndeudamientoDirecto(deudasValidadas(2)("Tipo").ToString(), deudasValidadas(2)("Saldo").ToString(), deudasValidadas(2)("Limite").ToString(), deudasValidadas(2)("Cuota").ToString()), 2).ToString())}</endeudamiento3>
                                   <endeudamiento4>{If(String.IsNullOrEmpty(deudasValidadas(3)("Tipo").ToString()), "", Math.Round(SumaEndeudamientoDirecto(deudasValidadas(3)("Tipo").ToString(), deudasValidadas(3)("Saldo").ToString(), deudasValidadas(3)("Limite").ToString(), deudasValidadas(3)("Cuota").ToString()), 2).ToString())}</endeudamiento4>
                                   <endeudamiento5>{If(String.IsNullOrEmpty(deudasValidadas(4)("Tipo").ToString()), "", Math.Round(SumaEndeudamientoDirecto(deudasValidadas(4)("Tipo").ToString(), deudasValidadas(4)("Saldo").ToString(), deudasValidadas(4)("Limite").ToString(), deudasValidadas(4)("Cuota").ToString()), 2).ToString())}</endeudamiento5>
                                   <endeudamiento6>{If(String.IsNullOrEmpty(deudasValidadas(5)("Tipo").ToString()), "", Math.Round(SumaEndeudamientoDirecto(deudasValidadas(5)("Tipo").ToString(), deudasValidadas(5)("Saldo").ToString(), deudasValidadas(5)("Limite").ToString(), deudasValidadas(5)("Cuota").ToString()), 2).ToString())}</endeudamiento6>
                                   <seguroVida>{Math.Round(seguros.SeguroVida, 2)}</seguroVida>
                                   <seguroDanios>{Math.Round(seguros.SeguroDanios, 2)}</seguroDanios>
                               </Root>")


            xDoc = XDocument.Parse(xmlRespuesta.OuterXml)
            strParametrosDetalle = JsonConvert.SerializeXNode(xDoc, Xml.Formatting.Indented)
            strParametrosDetalle = JsonConvert.SerializeObject(xDoc)
            EscribirRegistro("SALIDA", strGuiaRegistro, strParametrosDetalle, strRutaArchivo)

            Return xmlRespuesta

        Catch ex As Exception
            xmlRespuesta.LoadXml($"<respuesta>
                                  <error>{ex.Message}</error>
                                  <complete>{ex}</complete>
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

    Function ObtenerPlazoMesesXTipoGarantia(ByVal strTipoGarantia As String, ByVal strPlazoMeses As String) As Decimal 'Obtiene el plazo de cuotas para el tipo de garantia
        Dim strRutaArchivo As String = Server.MapPath("~/Archivos/PlazoMeses_TipoGarantia.xml")
        Dim documento As XDocument = XDocument.Load(strRutaArchivo)

        Dim strValor As String = (From entrada In documento.Descendants("Entrada")
                                  Where String.Equals(entrada.Element("Texto").Value, strTipoGarantia, StringComparison.OrdinalIgnoreCase)
                                  Select entrada.Element("Valor").Value).FirstOrDefault()

        If String.IsNullOrEmpty(strValor) Then
            Throw New ArgumentException("El texto proporcionado no está en el archivo XML.")
        End If

        Dim dcmXmlPlazo As Decimal = Convert.ToDecimal(strValor)

        If String.IsNullOrEmpty(strPlazoMeses) Then
            Return dcmXmlPlazo
        End If

        Dim dcmPlazoMeses As Decimal = Convert.ToDecimal(strPlazoMeses)
        Return dcmPlazoMeses

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
            Return Math.Round(dcmValorPresente / dcmTotalPeriodos, 2)
        End If

        pago = (dcmTasaMensual * dcmValorPresente) / (1 - Math.Pow(1 + dcmTasaMensual, -dcmTotalPeriodos))

        Return Math.Round(pago, 2)
    End Function

    Function SumaEndeudamientoDirecto(strTipoDeuda As String, strSaldoDeuda As String, strLimiteTarjeta As String, strCuotaDeuda As String)
        Dim dcmEndeudamientoDirecto As Decimal
        dcmEndeudamientoDirecto = 0

        If String.IsNullOrEmpty(strLimiteTarjeta) Then
            strLimiteTarjeta = "0"
        End If

        If Not String.IsNullOrEmpty(strTipoDeuda) Then
            If Not String.IsNullOrEmpty(strCuotaDeuda) And strCuotaDeuda <> "0" Then
                dcmEndeudamientoDirecto = Convert.ToDecimal(strCuotaDeuda)
            ElseIf Not String.IsNullOrEmpty(strSaldoDeuda) Then
                dcmEndeudamientoDirecto = CalcularCuotaDeudaXTipoDeuda(strTipoDeuda, Convert.ToDecimal(strSaldoDeuda), Convert.ToDecimal(strLimiteTarjeta))
            End If
        End If
        Return dcmEndeudamientoDirecto
    End Function

    Function CalcularCuotaDeudaXTipoDeuda(strTipoDeuda As String, dcmSaldoDeuda As Decimal, dcmLimiteTarjeta As Decimal)
        Dim dcmCuota As Decimal
        Dim intPlazo As Integer
        Dim dcmTasaPromedio As Decimal

        ObtenerPlazoTasaPromedioXTipo(strTipoDeuda, intPlazo, dcmTasaPromedio)

        'Saldo deuda es el valor inciial del prestamo, tarjeta interta, etc. El unico que cambia es tarjeta de credito mejor a 9m
        'En esa tarjeta el saldoDeuda representa el monto que debe y el limiteTarjeta representa el limite

        Select Case strTipoDeuda
            Case "Tarjeta de Credito Interna"
                dcmCuota = (dcmLimiteTarjeta * 0.33 / intPlazo) + (dcmLimiteTarjeta * 0.33) * (dcmTasaPromedio / 12)
            Case "Tarjeta de Credito Mayor a 9 m"
                dcmCuota = CalcularPago(dcmTasaPromedio / 12, intPlazo, dcmSaldoDeuda)
            Case "Tarjeta de Credito Menor a 9m"
                If dcmSaldoDeuda <= 0 Then
                    dcmCuota = CalcularPago(dcmTasaPromedio / 12, intPlazo, dcmLimiteTarjeta * 0.33)
                Else
                    Dim formula1 = CalcularPago(dcmTasaPromedio / 12, intPlazo, dcmLimiteTarjeta * 0.33)
                    Dim formula2 = CalcularPago(dcmTasaPromedio / 12, intPlazo, dcmSaldoDeuda)
                    If formula1 > formula2 Then
                        dcmCuota = formula1
                    Else
                        dcmCuota = formula2
                    End If
                End If
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

    Function PolizadeSeguro(dcmMontoSolicitado As Decimal, strConstrucciones As String, dcmTasaInteres As Decimal, strTipoGarantia As String) As ResultadoPolizaSeguro
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
            If String.Equals(strTipoGarantia, "Terreno", StringComparison.OrdinalIgnoreCase) Then
                dcmPrimaNeta = 0
            Else
                dcmPrimaNeta = dcmMontoSolicitado * Convert.ToDecimal(strValor.PrimaNeta) / 100
            End If
        End If

        dcmGastosEmision = dcmPrimaNeta * Convert.ToDecimal(strValor.GastosEmision) / 100
        dcmDecreto1422 = dcmPrimaNeta * Convert.ToDecimal(strValor.Decreto_1422) / 100
        dcmFraccionamiento = dcmPrimaNeta * Convert.ToDecimal(strValor.Fraccionamiento) / 100
        dcmIVA = (dcmPrimaNeta + dcmGastosEmision + dcmFraccionamiento) * Convert.ToDecimal(strValor.IVA) / 100

        dcmPrimaTotal = dcmPrimaNeta + dcmGastosEmision + dcmDecreto1422 + dcmIVA + dcmFraccionamiento
        dcmPrimaMes = dcmPrimaTotal / 12
        dcmIntereses = dcmMontoSolicitado * dcmTasaInteres / 100 / 360 * 30

        dcmGastos = dcmIntereses + dcmPrimaTotal + dcmVidaTotal

        Dim resultado As New ResultadoPolizaSeguro()
        resultado.SeguroVida = dcmVidaMes
        resultado.SeguroDanios = dcmPrimaMes
        Return resultado
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

    <WebMethod()>
    Public Function SumarNumeros(dcmNumero1 As Decimal, dcmNumero2 As Decimal) As XmlDocument
        Dim xmlRespuesta As New XmlDocument

        Try
            ' Calculamos la suma de los dos números
            Dim dcmResultado As Decimal = dcmNumero1 + dcmNumero2

            ' Creamos la respuesta XML con el resultado
            Context.Response.ContentType = "application/xml"
            xmlRespuesta.LoadXml($"<?xml version=""1.0""?><Root>
                                   <numero1>{dcmNumero1}</numero1>
                                   <numero2>{dcmNumero2}</numero2>
                                   <resultado>{dcmResultado}</resultado>
                                   <mensaje>Suma realizada exitosamente</mensaje>
                               </Root>")

            Return xmlRespuesta

        Catch ex As Exception
            xmlRespuesta.LoadXml($"<respuesta>
                                  <error>{ex.Message}</error>
                                  <complete>{ex}</complete>
                              </respuesta>")

            Return xmlRespuesta
        End Try
    End Function

End Class