Option Explicit On
'
Imports System.Data.SqlClient
Imports System.Text
Imports System.IO
Imports SAPbobsCOM.BoDataServerTypes
Imports SAPbobsCOM.BoSuppLangs
Imports SAPbobsCOM.BoFieldTypes
Imports SAPbobsCOM.BoUTBTableType
Imports SAPbobsCOM.BoObjectTypes
Imports SAPbobsCOM.BoFldSubTypes

Public Class SEI_SRV_VOXEL
    ''v2
    Private dInicio As Date
    Private dFinal As Date
    Private lResultado As Long

    Private Sub btnEjecutar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEjecutar.Click
        ''''  
        Me.lblmsg.Text = ""
        ''''  
        dInicio = Now

        ''''  
        If Not Me.ConectarSQLNative(go_conn) Then
            MsgBox("No se ha podido conectar a la BBDD")
            Exit Sub
        End If

        oCompany = New SAPbobsCOM.Company
        oCompany.Server = IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "S")    ' Server


        If Not ConectarSBO() Then
            MsgBox("Ha fallado la conexión a SBO")
            Exit Sub
        Else

        End If
        ''''  
        Dim crearcamps As Boolean = IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "CC")
        If crearcamps Then
            addUserFields_Documentos()
            addUserFields_Articles()
            addUserFields_BusinessPartner()
            addUserFields_Empresa()
        End If
        ''''  
        If Me.chkFacturacion.Checked Then
            txtProceso.Text = "Facturación"
            FacturacionElectronica()
        End If
        ''''
        'If Me.chkPedidos.Checked Then
        '    txtProceso.Text = "Leyendo ficheros de pedidos "
        '    CargarPedidosElectronicos()
        'End If
        ' ''''  
        If Me.chkConfirmacions.Checked Then
            txtProceso.Text = "Leyendo ficheros de aceptación de facturas "
            AcetpacionFacturasElectronicas()
        End If
        ''''  
        dFinal = Now
        lResultado = DateDiff(DateInterval.Minute, dInicio, dFinal)
        Me.lblmsg.Text = "Facturación finalizada. Minutos: " & lResultado.ToString
        '
        '-------------------------------------------------------------------------------
        ' DESECONEXIONES DE LAS BASES DE DATOS   '--------------------------------------
        '-------------------------------------------------------------------------------
        '
        Me.DesconectarSQLNative()
        ' '' ''Me.DesconectarSB0()
        Me.Dispose()
        '
    End Sub

    Public Function ConectarSQLNative(ByRef go_connaux As SqlConnection) As Boolean
        '
        Dim ls As String = ""
        '
        ConectarSQLNative = False
        Try
            ''''
            ls = "Server=" & IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "S") & ";" & _
                 "Database=" & IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "D") & ";" & _
                 "User id=" & IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "U") & ";" & _
                 "Password=" & IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "P") & ";"
            '''' 
            go_connaux = New SqlConnection
            go_connaux.ConnectionString = ls
            go_connaux.Open()
            '
            Me.txtUsuario.Text = IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "U")
            Me.txtempresa.Text = IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "D")
            '
            ConectarSQLNative = True
            '
        Catch ex As Exception
            ''''  MsgBox(ex.Message)
        End Try
        '
    End Function
    '
    Private Function ConectarSBO() As Boolean

        Dim lRetCode As Long
        Dim lErrCode As Long
        Dim sErrMsg As String
        '
        ConectarSBO = False
        lRetCode = 0
        sErrMsg = ""
        '
        oCompany = New SAPbobsCOM.Company
        oCompany.Server = IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "S")    ' Server
        oCompany.CompanyDB = IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "D") ' Base de Dades
        oCompany.UserName = IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "US")     'User SBO
        oCompany.Password = IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "PS")    'Password SBO
        oCompany.DbUserName = IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "U")       'User BD
        oCompany.DbPassword = IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "P")       'Password BD
        oCompany.UseTrusted = IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "Trusted") 'Trusted
        oCompany.language = ln_Spanish
        oCompany.DbServerType = dst_MSSQL2012
        '
        '// Connecting to a company DB
        lRetCode = oCompany.Connect

        If lRetCode = 0 Then
            Me.txtempresa.Text = oCompany.CompanyName
            Me.txtUsuario.Text = oCompany.UserName
            Application.DoEvents()
            Return True
        Else
            '
            'Me.txtempresa.Text = oCompany.CompanyName
            'Me.txtUsuario.Text = oCompany.UserName
            '
            oCompany.GetLastError(lErrCode, sErrMsg)
            'grabar log de errores
            '
            Me.lblmsg.Text = "Error: " & lErrCode.ToString & " " & sErrMsg
            Me.lblmsg.BackColor = Color.Red
            Application.DoEvents()
            Return False
        End If
        '// 
        '// Use Windows authentication for database server.
        '// True for NT server authentication,
        '// False for database server authentication.
        '// oCompany.UseTrusted = True
        '// 
    End Function
    '
    Private Sub DesconectarSQLNative()
        go_conn.Close()
    End Sub

    Private Sub DesconectarSB0()
        oCompany.Disconnect()
    End Sub
    '
    Private Sub FacturacionElectronica()
        Dim oFacturas As SEI_Facturas
        oFacturas = New SEI_Facturas(Me)
        oFacturas.GENERAR_FACTURES_TLY()

        Dim oAbonamentsFactures As SEI_AbonoFactura
        oAbonamentsFactures = New SEI_AbonoFactura(Me)
        oAbonamentsFactures.GENERAR_ABONAMENTSFACTURES_TLY()
    End Sub


    Private Sub AbonamentFacturacionElectronica()
        Dim oAbonamentsFactures As SEI_AbonoFactura
        oAbonamentsFactures = New SEI_AbonoFactura(Me)
        oAbonamentsFactures.GENERAR_ABONAMENTSFACTURES_TLY()
    End Sub

    '
    ' ''Private Sub AlbaransElectronics()
    ' ''    Dim oAlbarans As SEI_Albarans
    ' ''    oAlbarans = New SEI_Albarans(Me)
    ' ''    oAlbarans.GENERAR_ALBRANS_TLY()
    ' ''End Sub
    '
    Private Sub AcetpacionFacturasElectronicas()
        Dim oAcceptacio As SEI_AcceptacioF
        oAcceptacio = New SEI_AcceptacioF(Me)
        oAcceptacio.LLEGIR_ACCEPTACIONSF_TLY()
    End Sub
    '

    Private Sub CargarPedidosElectronicos()
        Dim comandes As SEI_CreaComandes
        comandes = New SEI_CreaComandes(Me)
        comandes.LLEGIR_COMANDES_TLY()
    End Sub

    Private Sub SEI_SRV_VOXEL_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ''''CheckBoxAlbarans.Checked = True
        chkFacturacion.Checked = True
        chkConfirmacions.Checked = True
        ''''   chkPedidos.Checked = True
    End Sub
    '
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkConfirmacions.CheckedChanged
    End Sub
    '
    Private Sub SEI_SRV_VOXEL_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        If IniGet(Application.StartupPath & "\S_SEI_Tokyo_VOXEL.ini", "Parametros", "A") = "S" Then
            btnEjecutar_Click(sender, e)
        End If
    End Sub
    '
    Private Sub btnSalir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSalir.Click
        Me.Dispose()
    End Sub

    Private Sub chkFacturacion_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkFacturacion.CheckedChanged

    End Sub

#Region "crearcamps"

    Private Sub addUserFields_Documentos()
        '
        Dim strTabla As String
        strTabla = "OINV"
        '
        If UserFieldsExist(strTabla, "SEIFiVox") = False Then
            Me.addCampAlpha(strTabla, "SEIFiVox", "Fichero Voxel", "254")
        End If
        '
    End Sub


    Private Sub addUserFields_Articles()
        '
        Dim strTabla As String
        strTabla = "OITM"
        '
        If UserFieldsExist(strTabla, "SEIArtVox") = False Then
            Me.addCampAlpha(strTabla, "SEIArtVox", "Articulo Voxel", "30")
        End If
        '
    End Sub

    Private Sub addUserFields_BusinessPartner()
        '
        Dim strTabla As String
        strTabla = "OCRD"
        '
        If UserFieldsExist(strTabla, "SEICliVox") = False Then
            Me.addCampAlpha(strTabla, "SEICliVox", "Cliente Voxel", "50")
        End If
        '
    End Sub



    Private Sub addUserFields_Empresa()
        '
        Dim strTabla As String
        strTabla = "OADM"
        '
        If UserFieldsExist(strTabla, "SEIDirVox") = False Then
            Me.addCampAlpha(strTabla, "SEIDirVox", "Directorio Voxel", "249")
        End If
        '
    End Sub



    Private Function AddUserTable(ByVal Nom As String, ByVal Descripcio As String, ByVal Tipus As SAPbobsCOM.BoUTBTableType) As Long
        '
        Dim lREtCode As Integer
        Dim lErrCode As String
        Dim sErrMsg As String
        If Not UserTablesExist(Nom) Then
            Dim oUTables As SAPbobsCOM.UserTablesMD = Nothing
            oUTables = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
            '
            oUTables.TableName = Nom
            oUTables.TableDescription = Descripcio
            oUTables.TableType = Tipus
            lREtCode = oUTables.Add
            '
            If lREtCode <> 0 Then
                oCompany.GetLastError(lErrCode, sErrMsg)
            End If
            '
            SEI_Globals.LiberarObjCOM(oUTables)
            '
        End If
        '
    End Function


    Public Function UserTablesExist(ByVal sTableName As String) As Boolean
        Dim oTmpRecordset As SAPbobsCOM.Recordset
        oTmpRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTmpRecordset.DoQuery("Select Count(*) From OUTB Where TableName = '" & sTableName & "'")
        If oTmpRecordset.Fields.Item(0).Value > 0 Then
            UserTablesExist = True
        Else
            UserTablesExist = False
        End If
        SEI_Globals.LiberarObjCOM(oTmpRecordset)

    End Function


    Private Function addCampData(ByVal vTaula As String, _
                                ByVal vCamp As String, _
                                ByVal vDesc As String, _
                                Optional ByVal vSubtipus As SAPbobsCOM.BoFldSubTypes = st_None, _
                                Optional ByVal vValorDefecte As String = "") As Long
        '
        'Per afegir un camp de data
        'Per exemple:
        '    addCampData "SEIEDI", "SEI_Data", "Data exportacio"
        '
        Dim lRetCode As Integer
        Dim lErrCode, sErrMsg As String

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD = Nothing
        Dim bTablaUsuario As Boolean
        '
        If InStr(vTaula, "@") <> 0 Then
            vTaula = Replace(vTaula, "@", "")
            bTablaUsuario = True
        End If
        '
        If Not bTablaUsuario Then
            '
            If UserFieldsExist(vTaula, vCamp) = False Then
                oUserFieldsMD = oCompany.GetBusinessObject(oUserFields)
                oUserFieldsMD.TableName = vTaula
                oUserFieldsMD.Name = vCamp
                oUserFieldsMD.Description = Mid(vDesc, 1, 30)
                oUserFieldsMD.Type = db_Date
                oUserFieldsMD.SubType = vSubtipus
                If vValorDefecte <> "" Then
                    oUserFieldsMD.DefaultValue = vValorDefecte
                End If
                lRetCode = oUserFieldsMD.Add
                addCampData = lRetCode
                If lRetCode <> 0 Then
                    oCompany.GetLastError(lErrCode, sErrMsg)
                    ''''      MsgBox(sErrMsg)
                End If
                SEI_Globals.LiberarObjCOM(oUserFieldsMD)
            End If
        Else
            '
            If UserFieldsExistT(vTaula, vCamp) = False Then
                oUserFieldsMD = oCompany.GetBusinessObject(oUserFields)
                oUserFieldsMD.TableName = vTaula
                oUserFieldsMD.Name = vCamp
                oUserFieldsMD.Description = Mid(vDesc, 1, 30)
                oUserFieldsMD.Type = db_Date
                oUserFieldsMD.SubType = vSubtipus
                If vValorDefecte <> "" Then
                    oUserFieldsMD.DefaultValue = vValorDefecte
                End If

                lRetCode = oUserFieldsMD.Add
                addCampData = lRetCode
                If lRetCode <> 0 Then
                    oCompany.GetLastError(lErrCode, sErrMsg)
                    '''     MsgBox(sErrMsg)
                End If
                SEI_Globals.LiberarObjCOM(oUserFieldsMD)
            End If
        End If
        '
    End Function


    Private Function UserFieldsExistT(ByVal sTableName As String, ByVal sFieldName As String) As Boolean
        Dim oTmpRecordset As SAPbobsCOM.Recordset
        oTmpRecordset = oCompany.GetBusinessObject(BoRecordset)
        If Mid(sTableName, 1, 1) <> "@" Then
            sTableName = "@" & sTableName
        End If
        ''
        oTmpRecordset.DoQuery("Select Count(*) From CUFD Where TableId = '" & sTableName & "' And AliasID = '" & sFieldName & "'")
        ''
        If oTmpRecordset.Fields.Item(0).Value > 0 Then
            UserFieldsExistT = True
        Else
            UserFieldsExistT = False
        End If
        SEI_Globals.LiberarObjCOM(oTmpRecordset)
        ''
    End Function


    Private Function addCampNumericFloat(ByVal vTaula As String, _
                 ByVal vCamp As String, _
                 ByVal vDesc As String, _
                 ByVal vSubtipus As SAPbobsCOM.BoFldSubTypes, _
                 Optional ByVal vValorDefecte As String = "") As Long
        '
        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD = Nothing
        Dim bTablaUsuario As Boolean
        '
        Dim lRetCode As Integer
        Dim lErrCode, sErrMsg As String


        If InStr(vTaula, "@") <> 0 Then
            vTaula = Replace(vTaula, "@", "")
            bTablaUsuario = True
        End If
        '
        If Not bTablaUsuario Then
            '
            If UserFieldsExist(vTaula, vCamp) = False Then
                oUserFieldsMD = oCompany.GetBusinessObject(oUserFields)
                oUserFieldsMD.TableName = vTaula
                oUserFieldsMD.Name = vCamp
                oUserFieldsMD.Description = Mid(vDesc, 1, 30)
                oUserFieldsMD.Type = db_Float
                oUserFieldsMD.SubType = vSubtipus
                If vValorDefecte <> "" Then
                    oUserFieldsMD.DefaultValue = vValorDefecte
                End If
                ''''
                lRetCode = oUserFieldsMD.Add
                addCampNumericFloat = lRetCode
                If lRetCode <> 0 Then
                    oCompany.GetLastError(lErrCode, sErrMsg)
                    '''      MsgBox(sErrMsg)
                End If
                ''''
                SEI_Globals.LiberarObjCOM(oUserFieldsMD)
                ''''
            End If
        Else
            '
            If UserFieldsExistT(vTaula, vCamp) = False Then
                oUserFieldsMD = oCompany.GetBusinessObject(oUserFields)
                oUserFieldsMD.TableName = vTaula
                oUserFieldsMD.Name = vCamp
                oUserFieldsMD.Description = Mid(vDesc, 1, 30)
                oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
                oUserFieldsMD.SubType = vSubtipus
                If vValorDefecte <> "" Then
                    oUserFieldsMD.DefaultValue = vValorDefecte
                End If
                ''''
                lRetCode = oUserFieldsMD.Add
                addCampNumericFloat = lRetCode
                If lRetCode <> 0 Then
                    oCompany.GetLastError(lErrCode, sErrMsg)
                    ''' MsgBox(sErrMsg)
                End If
                SEI_Globals.LiberarObjCOM(oUserFieldsMD)
                ''''
            End If
        End If
        '
    End Function


    Private Function UserFieldsExist(ByVal sTableName As String, ByVal sFieldName As String) As Boolean
        Dim oTmpRecordset As SAPbobsCOM.Recordset
        Dim ls As String
        '''''
        oTmpRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '''''
        ls = "Select Count(*) From CUFD Where TableId = '" & sTableName & "' And AliasID = '" & sFieldName & "'"
        oTmpRecordset.DoQuery(ls)
        '''''
        If oTmpRecordset.Fields.Item(0).Value > 0 Then
            UserFieldsExist = True
        Else
            UserFieldsExist = False
        End If
        SEI_Globals.LiberarObjCOM(oTmpRecordset)
        '''''
    End Function

    Function Explode(ByVal vCadena As String, ByVal vSeparador As String) As String()
        '
        'Retorna un array de valors de cadenes separades
        'exemple: explode("1|Un;2|Dos",";") retorna un array de 2 elements "1|Un" i "2|Dos"
        '
        Dim i As Integer
        Dim a() As String
        Dim fl As Boolean
        ReDim a(0)
        ' Destronat per
        'Retorn un array de valor de cadenes separades 
        ' 
        fl = False
        a(0) = ""
        Do
            i = InStr(vCadena, vSeparador)
            If i > 0 Then
                If Not fl Then
                    fl = True
                Else
                    ReDim Preserve a(0 To UBound(a) + 1)
                End If
                a(UBound(a)) = Mid(vCadena, 1, i - 1)
                If Len(vCadena) = Len(vSeparador) Then
                    'L'últim element és buit. l'afegim a l'arrai
                    ReDim Preserve a(0 To UBound(a) + 1)
                    a(UBound(a)) = ""
                End If
                vCadena = Mid(vCadena, i + (Len(vSeparador)))
            ElseIf Len(vCadena) > 0 Then
                If Not fl Then
                    fl = True
                Else
                    ReDim Preserve a(0 To UBound(a) + 1)
                End If
                a(UBound(a)) = vCadena
            End If
        Loop Until i = 0
        Explode = a
    End Function


    Private Function UserFieldID(ByVal sTableName As String, ByVal sFieldName As String) As Long
        '
        Dim oTmpRecordset As SAPbobsCOM.Recordset
        '
        oTmpRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '
        oTmpRecordset.DoQuery("Select FieldID From CUFD Where TableId = '" & sTableName & "' And AliasID = '" & sFieldName & "'")
        '
        UserFieldID = -1
        '
        If Not oTmpRecordset.EoF Then
            UserFieldID = oTmpRecordset.Fields.Item("FieldID").Value
        End If
        '
        SEI_Globals.LiberarObjCOM(oTmpRecordset)
        '
    End Function


    Private Function UserFieldsListExist(ByVal sTableName As String, _
                                       ByVal iFieldID As Integer, _
                                       ByVal sFldValue As String) As Boolean
        Dim oRcs As SAPbobsCOM.Recordset
        Dim ls As String
        '
        oRcs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'IndexID
        '
        ls = ""
        ls = ls & " SELECT TableID,FieldID,IndexID, FldValue    "
        ls = ls & " FROM UFD1 "
        ls = ls & " WHERE TableId = '" & sTableName & "'"
        ls = ls & " AND   FieldID = " & iFieldID.ToString
        ls = ls & " AND   FldValue = '" & sFldValue & "'"
        '
        oRcs.DoQuery(ls)
        '
        If Not oRcs.EoF Then
            UserFieldsListExist = True
        Else
            UserFieldsListExist = False
        End If
        SEI_Globals.LiberarObjCOM(oRcs)

    End Function

    Private Function addCampAlpha(ByVal vTaula As String, _
                ByVal vCamp As String, _
                ByVal vDesc As String, _
                ByVal vLong As Long, _
                Optional ByVal vLlistaValors As String = "", _
                Optional ByVal vSubtipus As SAPbobsCOM.BoFldSubTypes = st_None, _
                Optional ByVal vValorDefecte As String = "", _
                Optional ByVal sLinkTable As String = "") As Long
        '
        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        Dim a, b, i
        Dim iFieldID As Integer = 0
        Dim bListaPendiente As Boolean = False

        Dim lRetCode As Integer
        Dim lErrCode As String
        Dim sErrMsg As String

        If UserFieldsExist(vTaula, vCamp) = False Then
            SEI_Globals.LiberarObjCOM(oUserFieldsMD)
            '
            oUserFieldsMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            oUserFieldsMD.TableName = vTaula
            oUserFieldsMD.Name = vCamp
            oUserFieldsMD.Description = Mid(vDesc, 1, 30)
            oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
            oUserFieldsMD.SubType = vSubtipus
            oUserFieldsMD.EditSize = vLong
            oUserFieldsMD.Size = vLong
            '
            If Len(vLlistaValors) > 0 Then
                a = Explode(vLlistaValors, ";")
                For i = 0 To UBound(a)
                    b = Explode(a(i), "|")
                    oUserFieldsMD.ValidValues.Value = b(0)
                    oUserFieldsMD.ValidValues.Description = b(1)
                    oUserFieldsMD.ValidValues.Add()
                Next i
            End If
            '
            If vValorDefecte <> "" Then
                oUserFieldsMD.DefaultValue = vValorDefecte
            End If
            '
            If sLinkTable.Trim <> "" Then
                oUserFieldsMD.LinkedTable = sLinkTable
            End If
            '
            lRetCode = oUserFieldsMD.Add
            addCampAlpha = lRetCode
            If lRetCode <> 0 Then
                oCompany.GetLastError(lErrCode, sErrMsg)
                '''  MsgBox(sErrMsg)
            End If
            '
            SEI_Globals.LiberarObjCOM(oUserFieldsMD)
            '
        Else
            If Len(vLlistaValors) > 0 Then
                '
                iFieldID = UserFieldID(vTaula, vCamp)
                oUserFieldsMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                If oUserFieldsMD.GetByKey(vTaula, iFieldID) Then
                    ' Validar la Llista
                    a = Explode(vLlistaValors, ";")
                    For i = 0 To UBound(a)
                        b = Explode(a(i), "|")
                        If Not UserFieldsListExist(vTaula, iFieldID.ToString, b(0)) Then
                            bListaPendiente = True
                            oUserFieldsMD.ValidValues.Add()
                            oUserFieldsMD.ValidValues.Value = b(0)
                            oUserFieldsMD.ValidValues.Description = b(1)
                        End If
                    Next i
                    If bListaPendiente Then
                        lRetCode = oUserFieldsMD.Update
                        addCampAlpha = lRetCode
                        If lRetCode <> 0 Then
                            oCompany.GetLastError(lErrCode, sErrMsg)
                        End If
                    End If
                End If
                SEI_Globals.LiberarObjCOM(oUserFieldsMD)
            End If
        End If
    End Function

#End Region

    Private Sub chkPedidos_CheckedChanged(sender As System.Object, e As System.EventArgs)

    End Sub
End Class
