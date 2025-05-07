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

Public Class SEI_AcceptacioF
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
    Public Sub LLEGIR_ACCEPTACIONSF_TLY()
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

        sPath = IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "L")
        sPathD = IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "R")

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
        Conn2 = New ADODB.Connection
        Obre_Connexio_ADO(Conn2)
        '
        For Each File In files
            If File.Contains("ReturnReceipt") Then
                oXml = New XmlDocument
                oXml.Load(File)
                oItem = oXml.SelectNodes("//GeneralData")
                '' oXml.GetElementsByTagName("Lines")

                tipoAcuse = oItem.Item(0).Attributes("Type").InnerText
                dataAcuse = oItem.Item(0).Attributes("Date").InnerText

                oItem = oXml.SelectNodes("//ReturnReceipt")
                RecDoc = oItem.Item(0).Attributes("RecDoc").InnerText
                DocRef = oItem.Item(0).Attributes("DocRef").InnerText
                SenderID = oItem.Item(0).Attributes("SenderID").InnerText

                RecID = oItem.Item(0).Attributes("RecID").InnerText
                DeliveryDate = oItem.Item(0).Attributes("DeliveryDate").InnerText
                DeliveryTime = oItem.Item(0).Attributes("DeliveryTime").InnerText
                ls = ""
                If RecDoc.Contains("Factura") Then
                    ls = ls & " update  OINV  set  U_SEIacusv = 'S' where docnum  = " & DocRef
                Else
                    ls = ls & " update  ODLN  set  U_SEIacusv = 'S' where docnum  = " & DocRef
                End If
                '
                '' ''oDataReader = oSqlcomand.ExecuteReader()
                orecordset2 = Nothing
                orecordset2 = Conn2.Execute(ls)
                '
                Try
                    System.IO.File.Copy(File, sPathD & File.ToString.Replace(sPath, ""))
                    System.IO.File.Delete(File)
                Catch
                End Try
            End If
        Next
        ''''''' fi agafar dades cbg 
        go_conn.Close()
        '
        Try

        Catch ex As Exception
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
        '' MsgBox(strTmp)
        ' Clean up gracefully without risking infinite loop in error handler
        On Error GoTo 0
        GoTo Done

    End Sub
#End Region

End Class
