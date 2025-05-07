'
Option Explicit On
'
Imports System.Data.SqlClient
Imports System.Text
Imports System.IO
Imports SAPbobsCOM.BoObjectTypes

Imports System.Collections.Generic
Imports System.Collections.ObjectModel

Imports System.Reflection
Imports System.Xml


Imports System.Windows.Forms




Public Class SEI_Facturas
    Private Form As SEI_SRV_VOXEL

#Region "Contructor"
    '
    Public Sub New(ByRef o_Form As SEI_SRV_VOXEL)
        Form = o_Form
        'if creating controls via code, use initialize
        Initialize()
    End Sub

    Private Sub Initialize()

    End Sub
    '
    Public Sub GENERAR_FACTURES_TLY()
        Dim ls As String
        Dim oSqlcomand As SqlCommand
        Dim oDataReader As SqlClient.SqlDataReader = Nothing
        Dim oDataReader2 As SqlClient.SqlDataReader = Nothing
        Dim oRcsFactLin As SqlClient.SqlDataReader = Nothing
        Dim oXml As XmlDocument = Nothing
        Dim oItem As System.Xml.XmlNodeList = Nothing
        Dim sPath As String
        Dim sFichero As String
        Dim numSerie As String = "0"
        Dim iFila As Integer
        Dim go_conn3 As SqlConnection = Nothing
        Dim HashFEnviats As Hashtable = New Hashtable
        '
        ' dades de CBG
        Dim xCIF As String = ""
        Dim xAdreca As String = ""
        Dim xCompany As String = ""
        Dim xAddress As String = ""
        Dim xCity As String = ""
        Dim xProvince As String = ""
        Dim xCountry As String = ""
        Dim xZipCode As String = ""

        ls = ""
        ls = ls & " select top 1 OADM.TaxIdNum,OADM.CompnyName,OADM.CompnyAddr, adm1.City, adm1.State, adm1.Country, adm1.ZipCode, OCST.Name as nomprov from OADM "
        ls = ls & " left join ADM1 on oadm.currPeriod = adm1.currPeriod  "
        ls = ls & "   left join OCST   on OCST.code = adm1.State "
        oSqlcomand = New SqlCommand(ls, go_conn)
        oDataReader = oSqlcomand.ExecuteReader()
        While oDataReader.Read()
            xCIF = oDataReader("TaxIdNum").ToString
            xCompany = oDataReader("CompnyName").ToString
            xAddress = oDataReader("CompnyAddr").ToString
            xCity = oDataReader("City").ToString
            xProvince = oDataReader("nomprov").ToString
            If oDataReader("Country").ToString = "ES" Then
                xCountry = "ESP"
            Else
                xCountry = oDataReader("Country").ToString
            End If
            xZipCode = oDataReader("ZipCode").ToString
        End While
        ''''''' fi agafar dades cbg 
        go_conn.Close()
        SEI_SRV_VOXEL.ConectarSQLNative(go_conn)

        'Consutla Capçalera 
        ls = ""
        ls = ls & " SELECT"
        ls = ls & " T0.CardCode,  T0.DocNum,     T0.DocEntry,  T1.U_SEIEdiC,"
        ls = ls & " T0.DocDate,   T0.DocDueDate, T0.U_SEI_EDIR,T0.U_Dpto,"
        ls = ls & " T0.U_SEI_EDIF,T0.U_SEI_EDIE, T0.U_SEI_EDI2,T0.U_SEI_EDIL,"
        ls = ls & " T0.DocDate,   T0.CardName,   T1.Address,    T1.City      ,T1.ZipCode,"
        ls = ls & " T1.LicTradNum,T0.Doccur,     T0.GroupNum,  T2.U_SEI_EDIC,"
        ls = ls & " (T0.DocTotal- T0.VatSumSy + T0.DiscSumSy) as BASEIMP,"
        ls = ls & " T0.VatSumSy as TOTIMP,"
        ls = ls & " T0.DocTotal as TOTAL,"
        ls = ls & " T0.Discprcnt as PORCEN1,"   ' Porcentaje Cabecera
        ls = ls & " T0.DiscSumSy as IMPDES1,"   ' Importe Porcentaje Cabecera
        ls = ls & " T0.Comments,"                ' Observaciones
        ls = ls & " T0.U_SEI_EDI3,  T1.MailAddres, T1.MailCity, T1.MailZipCod  "              ' PO Quien emite EDI es el "DESTINATARIO" de la factura
        ls = ls & " FROM OINV T0"
        ls = ls & " INNER JOIN OCRD T1"
        ls = ls & " ON T0.CardCode= T1.CardCode "
        ls = ls & " LEFT OUTER JOIN OCTG T2"
        ls = ls & " ON T0.GroupNum= T2.GroupNum "
        ls = ls & " WHERE T1.QryGroup41 = 'Y' "         ' Cliente con Flag Facturas VOXEL
        ls = ls & " AND ISNULL(T0.U_SEIFiVox,'')=''"    ' Factura no exportada a Voxel   
        '
        Try
            oSqlcomand = New SqlCommand(ls, go_conn)
            oDataReader = oSqlcomand.ExecuteReader()
            While oDataReader.Read()
                iFila = 0
                oXml = ObtenerXML("FacturaVoxel.xml")
                oItem = oXml.SelectNodes("//GeneralData")
                oItem.Item(0).Attributes("Ref").InnerText = oDataReader("DocNum").ToString
                oItem.Item(0).Attributes("Type").InnerText = "FacturaComercial"
                oItem.Item(0).Attributes("Date").InnerText = Convert.ToDateTime(oDataReader("Docdate")).ToShortDateString
                '''' String.Format("{dd/mm/YYYY}", Now)
                '' '' ''Dim strFechaActual As String = Date.Now.ToShortDateString

                oItem = oXml.SelectNodes("//Supplier")
                oItem.Item(0).Attributes("CIF").InnerText = xCIF
                oItem.Item(0).Attributes("Company").InnerText = xCompany
                oItem.Item(0).Attributes("Address").InnerText = xAddress
                oItem.Item(0).Attributes("City").InnerText = xCity
                oItem.Item(0).Attributes("PC").InnerText = xZipCode
                oItem.Item(0).Attributes("Province").InnerText = xProvince
                oItem.Item(0).Attributes("Country").InnerText = xCountry

                oItem = oXml.SelectNodes("//Client")
                oItem.Item(0).Attributes("SupplierClientID").InnerText = oDataReader("CardCode").ToString
                oItem.Item(0).Attributes("CIF").InnerText = oDataReader("LicTradNum").ToString
                oItem.Item(0).Attributes("Company").InnerText = oDataReader("CardName").ToString
                oItem.Item(0).Attributes("Address").InnerText = oDataReader("Address").ToString
                oItem.Item(0).Attributes("City").InnerText = oDataReader("City").ToString
                oItem.Item(0).Attributes("PC").InnerText = oDataReader("ZipCode").ToString

                ''' aquí poso les dades de facturació 
                ls = ""
                ls = ls & " SELECT state,CRD1.Country,OCST.Name from CRD1  "
                ls = ls & "   left join OCST   on OCST.code = state "
                ls = ls & "where cardcode =  '" & oDataReader("CardCode").ToString & "' and adrestype = 'B'"
                SEI_SRV_VOXEL.ConectarSQLNative(go_conn3)
                oSqlcomand = New SqlCommand(ls, go_conn3)
                oDataReader2 = oSqlcomand.ExecuteReader()
                While oDataReader2.Read()
                    oItem.Item(0).Attributes("Province").InnerText = oDataReader2("Name").ToString
                    If oDataReader2("Country").ToString = "ES" Then
                        oItem.Item(0).Attributes("Country").InnerText = "ESP"
                    Else
                        oItem.Item(0).Attributes("Country").InnerText = oDataReader2("Country").ToString
                    End If
                End While
                go_conn3.Close()

                ''' aquí poso les dades de enviament 
                ls = ""
                ls = ls & " SELECT state,CRD1.Country,OCST.Name from CRD1  "
                ls = ls & "   left join OCST   on OCST.code = state "
                ls = ls & "where cardcode =  '" & oDataReader("CardCode").ToString & "' and adrestype = 'S'"
                SEI_SRV_VOXEL.ConectarSQLNative(go_conn3)
                oSqlcomand = New SqlCommand(ls, go_conn3)
                oDataReader2 = oSqlcomand.ExecuteReader()
                While oDataReader2.Read()
                    oItem = oXml.SelectNodes("//Customers/Customer")
                    oItem.Item(0).Attributes("SupplierClientID").InnerText = oDataReader("CardCode").ToString
                    oItem.Item(0).Attributes("Customer").InnerText = oDataReader("CardName").ToString
                    oItem.Item(0).Attributes("Address").InnerText = oDataReader("MailAddres").ToString
                    oItem.Item(0).Attributes("PC").InnerText = oDataReader("MailZipCod").ToString
                    oItem.Item(0).Attributes("City").InnerText = oDataReader("MailCity").ToString
                    oItem.Item(0).Attributes("Province").InnerText = oDataReader2("Name").ToString
                    If oDataReader2("Country").ToString = "ES" Then
                        oItem.Item(0).Attributes("Country").InnerText = "ESP"
                    Else
                        oItem.Item(0).Attributes("Country").InnerText = oDataReader2("Country").ToString
                    End If
                End While
                go_conn3.Close()

                '''' aqui miro les referències 
                ls = ""
                ls = ls & "  select top 1 baseref FROM  OINV T0 "
                ls = ls & "  INNER JOIN  INV1 T1  ON T0.DocEntry=T1.DocEntry  "
                ls = ls & "   where(T0.DocEntry = " & oDataReader("DocEntry").ToString & ")"
                ls = ls & "  group by baseref "

                SEI_SRV_VOXEL.ConectarSQLNative(go_conn3)
                oSqlcomand = New SqlCommand(ls, go_conn3)
                oDataReader2 = oSqlcomand.ExecuteReader()
                Dim xLinia As Integer = 0
                Dim oDocumentLines As Xml.XmlNode
                Dim oFirstRow As Xml.XmlNode
                Dim oNewRow As Xml.XmlNode
                While oDataReader2.Read()
                    If xLinia > 0 Then
                        oDocumentLines = oXml.SelectSingleNode("//References")
                        oFirstRow = oDocumentLines.FirstChild
                        oNewRow = oFirstRow.CloneNode(True)
                        oDocumentLines.AppendChild(oNewRow)
                    End If
                    oItem = oXml.SelectNodes("//References/Reference")
                    If IsNothing(oDataReader2("baseref").ToString) Then
                        oItem.Item(xLinia).Attributes("DNRef").InnerText = oDataReader("DocNum").ToString
                    Else
                        oItem.Item(xLinia).Attributes("DNRef").InnerText = oDataReader2("baseref").ToString
                    End If
                    xLinia = xLinia + 1
                End While
                go_conn3.Close()
                Me.Form.lblmsg.Text = "Factura: " & oDataReader("CardCode").ToString
                '-----------------------------------------------------------------------------------------------------
                '-----------------------------------------------------------------------------------------------------
                '############# LINFAC.TXT Detalle de la Factura (Sumatorio de Lineas necesario para la cabecera) #####
                '-----------------------------------------------------------------------------------------------------
                '-----------------------------------------------------------------------------------------------------
                oRcsFactLin = ObtenerRcsFactLin(oDataReader("DocEntry").ToString)
                '
                While oRcsFactLin.Read()
                    XML_Linea(oXml, oRcsFactLin, iFila)
                    iFila = iFila + 1
                End While
                '
                oItem = oXml.SelectNodes("//TotalSummary")
                oItem.Item(0).Attributes("SubTotal").InnerText = oDataReader("BASEIMP").ToString
                oItem.Item(0).Attributes("Tax").InnerText = oDataReader("TOTIMP").ToString
                oItem.Item(0).Attributes("Total").InnerText = oDataReader("TOTAL").ToString
                '
                sPath = Application.StartupPath() & "\"
                sPath = IniGet(Application.StartupPath & "\S_SEI_CBG_VOXEL.ini", "Parametros", "C") '''' "c:\PROVES VOXEL\"
                sFichero = sPath & "Factura_" & oDataReader("DocNum").ToString & "_" & "000" & ".xml"
                ''''sFichero = sPath & "Prova_" & numSerie & ".xml"
                oXml.Save(sFichero)


                ' '' ''ls = ls & " update  OINV  set  U_SEIFivox = '" & sFichero & "'where docentry  = " & oDataReader("DocEntry").ToString
                ' '' ''SEI_SRV_VOXEL.ConectarSQLNative(go_conn3)
                ' '' ''oSqlcomand = Nothing
                ' '' ''oSqlcomand = New SqlCommand(ls, go_conn3)
                ' '' ''oSqlcomand.ExecuteNonQuery()
                ' '' ''go_conn3.Close()

                HashFEnviats(oDataReader("DocEntry").ToString) = sFichero

                oDataReader2 = Nothing
                oSqlcomand = Nothing
            End While
 
            Dim oEnumerador As IDictionaryEnumerator
            oEnumerador = HashFEnviats.GetEnumerator
            While oEnumerador.MoveNext
                ls = ls & " update  OINV  set  U_SEIFivox = '" & oEnumerador.Value & "'where docentry  = " & oEnumerador.Key
                SEI_SRV_VOXEL.ConectarSQLNative(go_conn3)
                oSqlcomand = Nothing
                oSqlcomand = New SqlCommand(ls, go_conn3)
                oSqlcomand.ExecuteNonQuery()
                go_conn3.Close()
            End While
            HashFEnviats.Clear()
        Catch ex As Exception
            Dim oEnumerador As IDictionaryEnumerator
            oEnumerador = HashFEnviats.GetEnumerator
            While oEnumerador.MoveNext
                ls = ls & " update  OINV  set  U_SEIFivox = '" & oEnumerador.Value & "'where docentry  = " & oEnumerador.Key
                SEI_SRV_VOXEL.ConectarSQLNative(go_conn3)
                oSqlcomand = Nothing
                oSqlcomand = New SqlCommand(ls, go_conn3)
                oSqlcomand.ExecuteNonQuery()
                go_conn3.Close()
            End While
            HashFEnviats.Clear()
            Me.Form.lblmsg.Text = ex.Message
        Finally
            If Not IsNothing(oDataReader) Then
                oDataReader.Close()
            End If
        End Try

    End Sub
    '
    Private Sub UPDATE_SBO_CLIENTE(ByRef oDataReader As SqlClient.SqlDataReader)
        '
        Dim oCliente As SAPbobsCOM.BusinessPartners = Nothing
        Try
            '
            oCliente = oCompany.GetBusinessObject(oBusinessPartners)
            '
            If oCliente.GetByKey(oDataReader("Code").ToString) Then
                oCliente.EmailAddress = oDataReader("Email").ToString
                oCliente.Fax = oDataReader("Fax").ToString
                oCliente.Phone1 = oDataReader("Phone").ToString
                oCliente.Phone2 = oDataReader("Phone2").ToString
                oCliente.Cellular = oDataReader("Movil").ToString
                '
                If oCliente.Update <> 0 Then
                    Throw New Exception(RecuperarErrorSap())
                End If
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            SEI_Globals.LiberarObjCOM(oCliente)
        End Try
        '
    End Sub
    '

    Private Function ObtenerRcsFactLin(ByVal lDocEntry As String) As SqlClient.SqlDataReader
        '
        Dim ls As String
        Dim oRcs As SqlClient.SqlDataReader = Nothing
        Dim oSqlcomand As SqlCommand
        Dim coonLocal As SqlConnection
        '
        ls = ""
        ls = ls & " SELECT  T1.DocEntry, T1.LineNum,T2.CodeBars,T1.ItemCode ,"
        ls = ls & " T1.Dscription,T1.Quantity,T1.PriceBefDi,T1.Price,T1.VatGroup,T3.Rate,T1.VatSum,T1.LineTotal,"
        ls = ls & "  T1.DiscPrcnt  , T2.SalUnitMsr  "
        ls = ls & " FROM  OINV T0 INNER JOIN  INV1 T1"
        ls = ls & " ON T0.DocEntry=T1.DocEntry"
        ls = ls & " INNER JOIN OITM T2"
        ls = ls & " ON T1.ItemCode=T2.ItemCode"
        ls = ls & " LEFT OUTER JOIN OVTG T3 "
        ls = ls & " ON T1.VatGroup = T3.Code "
        ls = ls & " Where T0.DocEntry = " & lDocEntry
        ls = ls & " ORDER BY T1.DocEntry,T1.LineNum"
        '
        Try
            '
            SEI_SRV_VOXEL.ConectarSQLNative(coonLocal)
            oSqlcomand = New SqlCommand(ls, coonLocal)
            oRcs = oSqlcomand.ExecuteReader()
            '
            ObtenerRcsFactLin = oRcs
            '
            oRcs = Nothing
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            oRcs = Nothing
            oSqlcomand = Nothing
        End Try
        '
    End Function



    Public Shared Function GetEmbeddedResource(ByVal p_objTypeForNameSpace As Type, ByVal p_strScriptFileName As String) As String
        Dim s As StringBuilder = New StringBuilder
        Dim ass As [Assembly] = [Assembly].GetAssembly(p_objTypeForNameSpace)
        Dim sr As StreamReader
        '
        sr = New StreamReader(ass.GetManifestResourceStream(p_objTypeForNameSpace, p_strScriptFileName))
        s.Append(sr.ReadToEnd())
        '
        Return s.ToString()
        '
    End Function


    Private Function ObtenerXML(ByVal sFileName As String) As XmlDocument
        Dim oXMLDocument As XmlDocument = New XmlDocument
        oXMLDocument.LoadXml(GetEmbeddedResource(Me.GetType, sFileName))
        'SetFormPosition(oXMLDocument)
        Return oXMLDocument

    End Function


    Private Sub XML_Linea(ByRef oXML As Xml.XmlDocument, _
                          ByRef oRcs As SqlClient.SqlDataReader, _
                          ByRef iFila As Integer)
        '
        Dim oItem As Xml.XmlNodeList
        Dim oDocumentLines As Xml.XmlNode
        Dim oFirstRow As Xml.XmlNode
        Dim oNewRow As Xml.XmlNode
        '
        If iFila > 0 Then
            ' hauré de fer les referències abans 
            'Lineas Documento (Pedido de  Ventas)
            oDocumentLines = oXML.SelectSingleNode("//ProductList")
            'get the first row 
            oFirstRow = oDocumentLines.FirstChild
            'copy the first row the th new one -> for getting the same structure
            oNewRow = oFirstRow.CloneNode(True)
            'add the new row to the DocumentLines object
            oDocumentLines.AppendChild(oNewRow)
        End If
        '
        'Items
        oItem = oXML.SelectNodes("//ProductList/Product")
        oItem.Item(iFila).Attributes("SupplierSKU").InnerText = oRcs("ItemCode").ToString   ' Código de artículo interno del proveedor
        '''' oItem.Item(iFila).Attributes("CustomerSKU").InnerText = oRcs("ItemCode").ToString  ' Código de artículo interno del cliente
        oItem.Item(iFila).Attributes("Item").InnerText = String.Format("{0:0}", oRcs("Dscription").ToString).Replace(",", ".")
        oItem.Item(iFila).Attributes("Qty").InnerText = String.Format("{0:0.000000}", oRcs("Quantity").ToString).Replace(",", ".")
        oItem.Item(iFila).Attributes("MU").InnerText = oRcs("SalUnitMsr").ToString
        oItem.Item(iFila).Attributes("Total").InnerText = String.Format("{0:0.00}", oRcs("PriceBefDi") * oRcs("Quantity")).Replace(",", ".")
        oItem.Item(iFila).Attributes("UP").InnerText = oRcs("PriceBefDi").ToString.Replace(",", ".")
        '''   If Convert.ToDouble(oRcs("DiscPrcnt").ToString) <> 0 Then
        oItem = oXML.SelectNodes("//ProductList/Product/Discounts/Discount")
        oItem.Item(iFila).Attributes("Amount").InnerText = String.Format("{0:0.00}", (oRcs("Quantity") * (oRcs("PriceBefDi") * oRcs("DiscPrcnt") / 100))).Replace(",", ".")
        ''' End If
        ''' If Convert.ToDouble(oRcs("VatSum").ToString) <> 0 Then
        oItem = oXML.SelectNodes("//ProductList/Product/Taxes/Tax")
        oItem.Item(iFila).Attributes("Amount").InnerText = oRcs("VatSum").ToString.Replace(",", ".")
        oItem.Item(iFila).Attributes("Rate").InnerText = oRcs("rate").ToString.Replace(",", ".")
        ''' End If
    End Sub
#End Region

End Class
