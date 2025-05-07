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

Public Class SEI_CreaComandes
    Private Form As SEI_SRV_VOXEL


    Public Sub New(ByRef o_Form As SEI_SRV_VOXEL)
        Form = o_Form
        'if creating controls via code, use initialize
        Initialize()
    End Sub

    Private Sub Initialize()

    End Sub

    Public Sub LLEGIR_COMANDES_TLY()
        '''
        Dim ls As String
        Dim oSqlcomand As SqlCommand
        Dim oDataReader As SqlClient.SqlDataReader = Nothing
        Dim oDataReader2 As SqlClient.SqlDataReader = Nothing
        Dim oRcsFactLin As SqlClient.SqlDataReader = Nothing
        Dim oXml As XmlDocument = Nothing
        Dim oItem As System.Xml.XmlNodeList = Nothing
        Dim sPath, sPathD As String
        Dim sFichero As String
        Dim numSerie As String = "0"
        Dim iFila As Integer
        Dim go_conn3 As SqlConnection = Nothing
        Dim HashFEnviats As Hashtable = New Hashtable
        Dim hashliniaComentari As Hashtable = New Hashtable
        '''
        ' dades de CBG
        Dim xCIF As String = ""
        Dim xAdreca As String = ""
        Dim xCompany As String = ""
        Dim xAddress As String = ""
        Dim xCity As String = ""
        Dim xProvince As String = ""
        Dim xCountry As String = ""
        Dim xZipCode As String = ""
        Dim NomFitxerLog As String = ""
        ''' 
        sPath = IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "L")
        sPathD = IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "R")
        NomFitxerLog = IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "LO") & "\Log_" & CDate(Now).ToString("yyyyMMddhhmmss") & ".txt"
        '''
        Dim storefile As Directory   ''''' exposició previa de les entrades provades.
        Dim directory As String    '''' fi de les entrades provades.
        Dim files As String()
        Dim File As String

        files = storefile.GetFiles(sPath, "*")

        Dim tipoAcuse As String
        Dim dataAcuse As String
        Dim RecDoc As String
        Dim DocRef As String
        Dim SenderID As String
        Dim RecID As String
        Dim DeliveryTime As String
        Dim DeliveryDate As String
        '
        Dim orecordset2 As ADODB.Recordset
        Dim Conn2 As New ADODB.Connection
        '
        Try
            '
            Dim sMissatge1 As String = "Empieza la carga de pedidos... "
            Escriure_Fitxer_TXT(NomFitxerLog, sMissatge1)
            '
            ' ''Dim proceso As System.Diagnostics.Process()
            ' ''Dim fechaInicioProceso As Date
            ' ''proceso = System.Diagnostics.Process.GetProcessesByName("SEI_Tokyo_VOXEL")
            ' ''fechaInicioProceso = Now  ''' Process.GetCurrentProcess.StartTime '''   Now 
            ' ''For Each opro As System.Diagnostics.Process In proceso
            ' ''    If opro.StartTime <= fechaInicioProceso Then
            ' ''        opro.Kill()
            ' ''    End If
            ' ''Next

            Conn2 = New ADODB.Connection
            Obre_Connexio_ADO(Conn2)
            '
            Dim hh As String = ""
            For Each File In files
                If File.Contains("Pedido") Then
                    hashliniaComentari.Clear()
                    '''''''''''''''' coses a fer al crear la comanda '''''''''''''''''''''''''''''
                    Dim oDocument As SAPbobsCOM.Documents = Nothing
                    oDocument = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                    '' ''oDocument.CardCode = "C" & Cliente
                    Dim fecha As String = CDate(Now).ToString("yyyyMMdd")
                    oDocument.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
                    oDocument.DocDate = fecha.Substring(6, 2) & "/" & fecha.Substring(4, 2) & "/" & fecha.Substring(0, 4)
                    oDocument.DocDueDate = fecha.Substring(6, 2) & "/" & fecha.Substring(4, 2) & "/" & fecha.Substring(0, 4)
                    oDocument.TaxDate = fecha.Substring(6, 2) & "/" & fecha.Substring(4, 2) & "/" & fecha.Substring(0, 4)
                    oXml = New XmlDocument
                    oXml.Load(File)

                    Dim sMissatge2 As String = "Cargando fichero... " & File
                    Escriure_Fitxer_TXT(NomFitxerLog, sMissatge2)

                    oItem = oXml.SelectNodes("//GeneralData")
                    If Not IsNothing(oItem.Item(0)) Then

                        tipoAcuse = ""
                        If Not IsNothing(oItem.Item(0).Attributes("Type")) Then
                            tipoAcuse = oItem.Item(0).Attributes("Type").InnerText
                        End If

                        dataAcuse = ""
                        If Not IsNothing(oItem.Item(0).Attributes("Date")) Then
                            dataAcuse = oItem.Item(0).Attributes("Date").InnerText
                        End If

                        If Not IsNothing(oItem.Item(0).Attributes("Ref")) Then
                            oDocument.NumAtCard = oItem.Item(0).Attributes("Ref").InnerText
                        End If
                    End If

                    oItem = oXml.SelectNodes("//Supplier")
                    If Not IsNothing(oItem.Item(0)) Then
                        hh = ""
                    End If

                    oItem = oXml.SelectNodes("//Client")   ''això és el client que se li envia la factura el cient del document
                    If Not IsNothing(oItem) Then
                        Dim codiIC As String = ""
                        If Not IsNothing(oItem.Item(0).Attributes("ClientID")) Then
                            codiIC = oItem.Item(0).Attributes("ClientID").InnerText
                        End If

                        Dim cif As String = ""
                        If Not IsNothing(oItem.Item(0).Attributes("CIF")) Then
                            cif = oItem.Item(0).Attributes("CIF").InnerText
                        End If

                        Dim nomClient As String = ""
                        If Not IsNothing(oItem.Item(0).Attributes("Company")) Then
                            nomClient = oItem.Item(0).Attributes("Company").InnerText
                        End If

                    End If
                    oItem = oXml.SelectNodes("//Customers/Customer")  '' això es la direcció de entrega 
                    If Not IsNothing(oItem.Item(0)) Then
                        Dim linies As Integer = oItem.Count
                        For iFila = 0 To linies - 1

                            Dim codiGLN As String = ""
                            If Not IsNothing(oItem.Item(0).Attributes("CustomerID")) Then
                                codiGLN = oItem.Item(0).Attributes("CustomerID").InnerText
                            End If

                            Dim codiDelproveidor As String = ""
                            If Not IsNothing(oItem.Item(0).Attributes("SupplierClientID")) Then
                                codiDelproveidor = oItem.Item(0).Attributes("SupplierClientID").InnerText   ''' codi del client al sap be benito álvarez.
                            End If

                            '''    oDocument.CardCode = "C0" & codiDelproveidor
                            oDocument.CardCode = codiDelproveidor

                            Dim SupplierCustomerID As String = ""
                            If Not IsNothing(oItem.Item(0).Attributes("SupplierCustomerID")) Then
                                SupplierCustomerID = oItem.Item(0).Attributes("SupplierCustomerID").InnerText
                            End If

                            Dim nomDireccio As String = ""
                            If Not IsNothing(oItem.Item(0).Attributes("Customer")) Then
                                nomDireccio = oItem.Item(0).Attributes("Customer").InnerText
                            End If

                            Dim direccio As String = ""
                            If Not IsNothing(oItem.Item(0).Attributes("Address")) Then
                                direccio = oItem.Item(0).Attributes("Address").InnerText
                            End If

                            Dim ciutat As String = ""
                            If Not IsNothing(oItem.Item(0).Attributes("City")) Then
                                ciutat = oItem.Item(0).Attributes("City").InnerText
                            End If

                            Dim CodiPostal As String = ""
                            If Not IsNothing(oItem.Item(0).Attributes("PC")) Then
                                CodiPostal = oItem.Item(0).Attributes("PC").InnerText
                            End If

                            Dim Provincia As String = ""
                            If Not IsNothing(oItem.Item(0).Attributes("Province")) Then
                                Provincia = oItem.Item(0).Attributes("Province").InnerText
                            End If

                            Dim pais As String = ""
                            If Not IsNothing(oItem.Item(0).Attributes("Country")) Then
                                pais = oItem.Item(0).Attributes("Country").InnerText
                            End If

                        Next
                    End If
                    '
                    oItem = oXml.SelectNodes("//realsender")
                    If Not IsNothing(oItem.Item(0)) Then
                        hh = ""
                    End If
                    oItem = oXml.SelectNodes("//Comments/Comment")
                    If Not IsNothing(oItem.Item(0)) Then
                        Dim linies As Integer = oItem.Count
                        For iFila = 0 To linies - 1

                            Dim comentariText As String = ""
                            If Not IsNothing(oItem.Item(iFila).Attributes("Subject")) Then
                                comentariText = oItem.Item(iFila).Attributes("Subject").InnerText
                            End If
                            Dim missatge As String = ""
                            If Not IsNothing(oItem.Item(iFila).Attributes("Msg")) Then
                                missatge = oItem.Item(iFila).Attributes("Msg").InnerText
                            End If

                            If Trim(oDocument.Comments.ToString) <> "" Then
                                oDocument.Comments = oDocument.Comments & vbNewLine & comentariText & " - " & missatge
                            Else
                                oDocument.Comments = comentariText & " - " & missatge
                            End If

                        Next
                    End If
                    oItem = oXml.SelectNodes("//References/Reference")
                    If Not IsNothing(oItem.Item(0)) Then
                        hh = ""
                    End If
                    oItem = oXml.SelectNodes("//Attachments/CAttachment")
                    If Not IsNothing(oItem.Item(0)) Then
                        hh = ""
                    End If
                    ''''''''''''''''''''
                    ''''''''''''''''''''
                    oItem = oXml.SelectNodes("//ProductList/Product/Remarks/Remark")
                    If Not IsNothing(oItem.Item(0)) Then
                        Dim linies As Integer = oItem.Count
                        For iFila = 0 To linies - 1
                            Dim comentariText As String = ""
                            If Not IsNothing(oItem.Item(iFila).Attributes("Subject")) Then
                                comentariText = oItem.Item(iFila).Attributes("Subject").InnerText
                            End If
                            Dim liniaComentada As String = ""
                            If Not IsNothing(oItem.Item(iFila).Attributes("Msg")) Then
                                liniaComentada = oItem.Item(iFila).Attributes("Msg").InnerText
                            End If
                            ''''' hashliniaComentari(liniaComentada) = oItem.Item(iFila).Attributes("Subject").InnerText
                            If Trim(liniaComentada) <> "" Then
                                hashliniaComentari(iFila) = liniaComentada    ''''oItem.Item(iFila).Attributes("Msg").InnerText
                            Else
                                hashliniaComentari(iFila) = oItem.Item(iFila).Attributes("Subject").InnerText
                            End If
                        Next
                    End If
                    '''''''''''''''''''
                    '''''''''''''''''''
                    oItem = oXml.SelectNodes("//ProductList/Product")
                    If Not IsNothing(oItem.Item(0)) Then
                        Dim linies As Integer = oItem.Count
                        For iFila = 0 To linies - 1
                            Dim provepidorCI As String = ""
                            If Not IsNothing(oItem.Item(iFila).Attributes("SupplierSKU")) Then
                                provepidorCI = oItem.Item(iFila).Attributes("SupplierSKU").InnerText   ''' codi article al sap
                            End If
                            Dim clientCI As String = ""
                            If Not IsNothing(oItem.Item(iFila).Attributes("CustomerSKU")) Then
                                clientCI = oItem.Item(iFila).Attributes("CustomerSKU").InnerText  '' codi de l'article en el client de sap
                            End If
                            Dim NomArticle As String = oItem.Item(iFila).Attributes("Item").InnerText
                            Dim quantitat As Double = CDbl(Replace(oItem.Item(iFila).Attributes("Qty").InnerText, ".", ","))
                            Dim unitatsMesura As String = donaUnitatMesura(oItem.Item(iFila).Attributes("MU").InnerText)
                            Dim preubrutunitari As Double = CDbl(Replace(oItem.Item(iFila).Attributes("UP").InnerText, ".", ","))
                            Dim lineTotal As Double = CDbl(Replace(oItem.Item(iFila).Attributes("Total").InnerText, ".", ","))
                            Dim importnet As Double = 0
                            If Not IsNothing(oItem.Item(iFila).Attributes("NetAmount")) Then
                                importnet = CDbl(Replace(oItem.Item(iFila).Attributes("NetAmount").InnerText, ".", ","))
                            End If
                            '''''  
                            Dim commentaris As String = ""
                            If Not IsNothing(oItem.Item(iFila).Attributes("Comment")) Then
                                commentaris = oItem.Item(iFila).Attributes("Comment").InnerText
                            End If
                            '''''
                            Dim linenum As Integer = 0
                            If Not IsNothing(oItem.Item(iFila).Attributes("SourceLineNumber")) Then
                                linenum = oItem.Item(iFila).Attributes("SourceLineNumber").InnerText
                            End If
                            ''''' 
                            If oDocument.Lines.ItemCode <> "" Then oDocument.Lines.Add()
                            oDocument.Lines.ItemCode = provepidorCI
                            oDocument.Lines.Quantity = quantitat
                            ''oDocument.Lines.UnitPrice = preubrutunitari
                            ''If importnet <> 0 Then
                            ''    oDocument.Lines.LineTotal = importnet
                            ''End If
                            oDocument.Lines.MeasureUnit = unitatsMesura
                            '''''  oDocument.Lines.UserFields.Fields.Item("U_SEIQtyPedida").Value = quantitat
                            ''''' 
                            ''''If hashliniaComentari.ContainsKey(linenum) Then
                            Dim oitem2 As XmlNodeList = oItem.Item(iFila).SelectNodes("Remarks/Remark")
                            If Not IsNothing(oitem2.Item(0)) Then
                                Dim linies2 As Integer = oitem2.Count
                                For iFila2 As Integer = 0 To linies2 - 1
                                    Dim comentariText2 As String = ""
                                    If Not IsNothing(oitem2.Item(iFila2).Attributes("Subject")) Then
                                        comentariText2 = oitem2.Item(iFila2).Attributes("Subject").InnerText
                                    End If
                                    Dim liniaComentada2 As String = ""
                                    If Not IsNothing(oitem2.Item(iFila2).Attributes("Msg")) Then
                                        liniaComentada2 = oitem2.Item(iFila2).Attributes("Msg").InnerText
                                    End If
                                    If Trim(liniaComentada2) <> "" Then
                                        oDocument.Lines.FreeText = liniaComentada2
                                    Else
                                        If Trim(comentariText2) <> "" Then
                                            oDocument.Lines.FreeText = comentariText2
                                        End If
                                    End If
                                Next
                            End If

                            '''''
                            ''If hashliniaComentari.ContainsKey(iFila) Then
                            ''    oDocument.Lines.FreeText = hashliniaComentari(iFila)
                            ''End If
                            '''''
                        Next
                    End If

                    oItem = oXml.SelectNodes("//ProductList/Product/Taxes/Tax")
                    If Not IsNothing(oItem.Item(0)) Then
                        '''    Dim import As Double = oItem.Item(iFila).Attributes("Amount").InnerText
                    End If
                    ''
                    oItem = oXml.SelectNodes("//ProductList/Product/Discounts/Discount")
                    If Not IsNothing(oItem.Item(0)) Then
                        hh = ""
                    End If
                    ''
                    oItem = oXml.SelectNodes("//ProductList/Product/Fees/Fee")
                    If Not IsNothing(oItem.Item(0)) Then
                        hh = ""
                    End If
                    '' 
                    oItem = oXml.SelectNodes("//ProductList/Product/References/Reference")
                    If Not IsNothing(oItem.Item(0)) Then
                        hh = ""
                    End If
                    '' 
                    oItem = oXml.SelectNodes("//ProductList/Product/ServicesData/ServiceData")
                    If Not IsNothing(oItem.Item(0)) Then
                        hh = ""
                    End If
                    oItem = oXml.SelectNodes("//ProductList/Product/References/Reference")
                    If Not IsNothing(oItem.Item(0)) Then
                        hh = ""
                    End If
                    oItem = oXml.SelectNodes("//GlobalDiscounts/Discount")
                    If Not IsNothing(oItem.Item(0)) Then
                        hh = ""
                    End If
                    oItem = oXml.SelectNodes("//TaxSummary/Tax")
                    If Not IsNothing(oItem.Item(0)) Then
                        hh = ""
                    End If
                    oItem = oXml.SelectNodes("//FeesSummary/Fee")
                    If Not IsNothing(oItem.Item(0)) Then
                        hh = ""
                    End If
                    oItem = oXml.SelectNodes("//DueDates/DueDate")
                    If Not IsNothing(oItem.Item(0)) Then
                        hh = ""
                    End If
                    oItem = oXml.SelectNodes("//DueDates/DueDate")
                    If Not IsNothing(oItem.Item(0)) Then
                        hh = ""
                    End If
                    oItem = oXml.SelectNodes("//TotalSummary")
                    If Not IsNothing(oItem.Item(0)) Then
                        hh = ""
                    End If

                    Dim sMissatge3 As String = "Voy a crear del documento correspondiente a... " & File
                    Escriure_Fitxer_TXT(NomFitxerLog, sMissatge2)
                    Dim haFallat As Boolean = False
                    If oDocument.Lines.ItemCode <> "" Then
                        If oDocument.Add <> 0 Then
                            haFallat = True
                            Dim sError As String
                            sError = RecuperarErrorSap()
                            Escriure_Fitxer_TXT(NomFitxerLog, sError)
                        Else
                            Dim sMissatge As String = "Se ha creado el pedido: " & oCompany.GetNewObjectKey
                            Escriure_Fitxer_TXT(NomFitxerLog, sMissatge)

                            Try
                                Dim sMissatge4 As String = "voy a mover el fichero : " & File & " a la carpeta de ficheros procesados " & sPathD & File.ToString.Replace(sPath, "")
                                Escriure_Fitxer_TXT(NomFitxerLog, sMissatge4)
                                System.IO.File.Copy(File, sPathD & File.ToString.Replace(sPath, ""), True)
                                System.IO.File.Delete(File)
                            Catch
                            End Try
                        End If
                    Else
                        Dim sMissatge As String = "No se va a crear el pedido no hay lineas. "
                        Escriure_Fitxer_TXT(NomFitxerLog, sMissatge)
                    End If
                End If
            Next
            '''''''' 
            go_conn.Close()
            ''''''''
        Catch ex As Exception
            Me.Form.lblmsg.Text = ex.Message
            Dim sMissatge As String = "Error1 al crear el pedido: " & ex.Message
            Escriure_Fitxer_TXT(NomFitxerLog, sMissatge)
        End Try

        Try

        Catch ex As Exception
            Me.Form.lblmsg.Text = ex.Message
            Dim sMissatge As String = "Error2 al crear el pedido: " & ex.Message
            Escriure_Fitxer_TXT(NomFitxerLog, sMissatge)
        Finally
            If Not IsNothing(oDataReader) Then
                oDataReader.Close()
            End If
        End Try

    End Sub
    '
    Private Sub UPDATE_SBO_CLIENTE(ByRef oDataReader As SqlClient.SqlDataReader)
        '''
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
                '''
                If oCliente.Update <> 0 Then
                    Throw New Exception(RecuperarErrorSap())
                End If
                '''
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            SEI_Globals.LiberarObjCOM(oCliente)
        End Try
        '''
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
        '
    End Function


    Private Sub Obre_Connexio_ADO(ByRef Conn1 As ADODB.Connection)
        '''   Dim Conn1 As New ADODB.Connection
        Dim Cmd1 As New ADODB.Command
        Dim Errs1 As ADODB.Errors
        Dim Rs1 As New ADODB.Recordset
        Dim i As Integer
        Dim AccessConnect As String
        ' Error Handling Variables
        Dim strTmp As String
        On Error GoTo AdoError  ' Full Error Handling which traverses
        ' Connection object
        Dim sSQLstring As String

        sSQLstring = "Provider=SQLNCLI11;" & _
                            "Server=" & IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "S") & ";" & _
                            "Database=" & IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "D") & ";" & _
                            "UID=" & IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "U") & ";" & _
                            "PWD=" & IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "P") & ";"

        Conn1.Open(sSQLstring)

Done:
        Rs1 = Nothing
        Cmd1 = Nothing
        Exit Sub

AdoError:
        i = 1
        On Error Resume Next

        ' Enumerate Errors collection and display properties of
        ' each Error object (if Errors Collection is filled out)
        Errs1 = Conn1.Errors
        Dim errLoop As ADODB.Error

        'tret pa no se de quina clase es tracat errLoop
        For Each errLoop In Errs1
            With errLoop
                strTmp = strTmp & vbCrLf & "ADO Error # " & i & ":"
                strTmp = strTmp & vbCrLf & "   ADO Error   # " & .Number
                strTmp = strTmp & vbCrLf & "   Description   " & .Description
                strTmp = strTmp & vbCrLf & "   Source        " & .Source
                i = i + 1
            End With
        Next

AdoErrorLite:
        ' Get VB Error Object's information
        strTmp = strTmp & vbCrLf & "VB Error # " & Str(Err.Number)
        strTmp = strTmp & vbCrLf & "   Generated by " & Err.Source
        strTmp = strTmp & vbCrLf & "   Description  " & Err.Description
        'MsgBox(strTmp)
        ' Clean up gracefully without risking infinite loop in error handler
        On Error GoTo 0
        GoTo Done
        '
    End Sub


    Public Sub Escriure_Fitxer_TXT(ByVal sRutaFitxer As String, ByVal sTexteLinia As String)
        '
        Dim oFitxer As New System.IO.StreamWriter(sRutaFitxer, True, System.Text.Encoding.GetEncoding(1252))
        ' 
        oFitxer.WriteLine(sTexteLinia, System.Text.Encoding.GetEncoding(1252))
        oFitxer.Flush()
        oFitxer.Close()
        '
    End Sub

End Class
