'
Imports System.Data.SqlClient
Imports System.Data.OleDb
'
Module SEI_Globals

    Public go_connOLEDB As OleDbConnection
    Public go_conn As SqlConnection
    Public oCompany As SAPbobsCOM.Company

#Region "Funciones Fichero.INI"

    '
    ' Leer una clave de un fichero INI
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Integer, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    '
    Public Function IniGet(ByVal sFileName As String, ByVal sSection As String, ByVal sKeyName As String, Optional ByVal sDefault As String = "") As String
        '--------------------------------------------------------------------------
        ' Devuelve el valor de una clave de un fichero INI
        ' Los parámetros son:
        '   sFileName   El fichero INI
        '   sSection    La sección de la que se quiere leer
        '   sKeyName    Clave
        '   sDefault    Valor opcional que devolverá si no se encuentra la clave
        '--------------------------------------------------------------------------
        ' sSection ->   "Parametros"
        ' sKeyName ->   "U" , "I" , "P"
        '
        ' [Parametros]
        ' U = sa
        ' I = IG
        ' P =seidor.65

        Dim ret As Integer
        Dim sRetVal As String
        '
        sRetVal = New String(Chr(0), 255)
        '
        ret = GetPrivateProfileString(sSection, sKeyName, sDefault, sRetVal, Len(sRetVal), sFileName)
        If ret = 0 Then
            Return sDefault
        Else
            Return Left(sRetVal, ret)
        End If
    End Function

#End Region


#Region "Funciones Tipos de Datos"

    Function Formato_Decimales_IG(ByVal Valor As Object) As String

        Valor = Valor.ToString.Replace(".", "")
        Valor = Valor.ToString.Replace(",", ".")
        Return Valor.ToString

    End Function

    Function NullToText(ByVal Valor As Object) As String

        If IsDBNull(Valor) Or Valor.ToString.Trim = "" Then
            Return " "
        Else
            Return Valor.ToString
        End If

    End Function

    Function NullToInt(ByVal Valor As Object) As Integer

        If IsDBNull(Valor) Or Valor.ToString.Trim = "" Then
            Return 0
        Else
            Return Convert.ToInt32(Valor.ToString)   ' Pasar a integer
        End If

    End Function

    Function NullToDoble(ByRef Valor As Object) As Double

        If IsDBNull(Valor) Or Trim(Valor.ToString) = "" Then
            Return 0
        Else
            Return Convert.ToDouble(Valor.ToString)  ' Pasar a double
        End If

    End Function

    Function NullToData(ByVal Valor As Object) As String

        If IsDBNull(Valor) Or Valor.GetType.ToString = "" Then
            Return "NULL"
        Else
            Return String.Format("{0:d}", Valor)
        End If

    End Function

    Function NullToHora(ByVal Valor As Object) As String

        If IsDBNull(Valor) Or Valor.GetType.ToString = "" Then
            Return "NULL"
        Else
            Return String.Format("{0:t}", Valor)
        End If

    End Function

    Function NullToLong(ByVal Valor As Object) As Long

        If IsDBNull(Valor) Or Trim(Valor.ToString) = "" Then
            Return 0
        Else
            Return CType(Valor, Long)   ' Pasar a Long
        End If

    End Function

    Function NullToSiNo(ByVal Valor As Object) As String

        If IsDBNull(Valor) Or Valor.GetType.ToString = "" Then
            Return "N"
        Else
            Return "Y"
        End If

    End Function

    Function IntToBooleanS_N(ByVal Valor As Object) As String

        If IsDBNull(Valor) Then
            Return "N"
        ElseIf Valor = 0 Then
            Return "N"
        Else
            Return "S"
        End If

    End Function

    Function IntToBoolean(ByVal Valor As Object) As String

        If IsDBNull(Valor) Then
            Return "N"
        ElseIf Valor = 0 Then
            Return "N"
        Else
            Return "Y"
        End If

    End Function

    Function BooleanToInt(ByVal Valor As Object) As Integer

        If Valor = "True" Then
            Return "1"
        Else
            Return "0"
        End If

    End Function
    '
    Public Function NowDateToString() As String
        NowDateToString = Now.Date.ToString("yyyyMMdd")
    End Function
    '
    ' Poner un valor entre comillas
    Public Function sC(ByVal sValor As String) As String
        sC = "'" & sValor.Replace("'", "''") & "'"
    End Function

#End Region

    '
    Public Sub LiberarObjCOM(ByRef oObjCOM As Object, Optional ByVal bCollect As Boolean = False)
        '
        'Liberar y destruir Objecto com 
        ' En los UDO'S es necesario utilizar GC.Collect  para eliminarlos de la memoria
        If Not IsNothing(oObjCOM) Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oObjCOM)
            oObjCOM = Nothing
            If bCollect Then
                GC.Collect()
            End If
        End If

    End Sub
    '
    Public Function RecuperarErrorSap() As String

        Dim sError As String
        Dim lErrCode As Long
        Dim sErrMsg As String
        '
        lErrCode = 0
        sErrMsg = ""
        oCompany.GetLastError(lErrCode, sErrMsg)
        sError = "Error: " & lErrCode.ToString & " " & sErrMsg
        '
        Return sError
        '
    End Function



    Public Function DonaArticleVoxel(ByVal xxCodiSAP As String) As String

        Dim oTmpRecordset As SAPbobsCOM.Recordset

        oTmpRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        oTmpRecordset.DoQuery("Select U_SEIArtVox From OITM Where ItemCode = '" & xxCodiSAP & "'")

        If Not IsNothing(oTmpRecordset.Fields.Item(0)) Then
            If oTmpRecordset.Fields.Item(0).Value <> "" Then
                DonaArticleVoxel = oTmpRecordset.Fields.Item("U_SEIArtVox").Value
            Else
                DonaArticleVoxel = xxCodiSAP   ''' si no hi ha codi voxel retorna el codi article del sap
            End If
        Else
            DonaArticleVoxel = xxCodiSAP   ''' si no hi ha codi voxel retorna el codi article del sap
        End If

        SEI_Globals.LiberarObjCOM(oTmpRecordset)

    End Function



    Public Function DonaArticleSAP(ByVal xxCodiVoxel As String) As String
        Dim oTmpRecordset As SAPbobsCOM.Recordset

        oTmpRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        oTmpRecordset.DoQuery("Select  ItemCode From OITM Where U_SEIArtVox = '" & xxCodiVoxel & "'")

        If oTmpRecordset.Fields.Item(0).Value > 0 Then
            DonaArticleSAP = oTmpRecordset.Fields.Item("ItemCode").Value
        Else
            DonaArticleSAP = ""
        End If
        SEI_Globals.LiberarObjCOM(oTmpRecordset)
    End Function


    Public Function DonaClientVoxel(ByVal codiProveidor As String) As String
        Dim oTmpRecordset As SAPbobsCOM.Recordset
        oTmpRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'Dim ls As String = " select isnull(crd1.glbllocnum,'') as GLN  from ocrd  "
        'ls = ls & " left join crd1 on ocrd.cardcode = crd1.cardcode  and crd1.AdresType = 'B'  and ocrd.Billtodef = crd1.Address  "
        'ls = ls & "  where isnull(crd1.glbllocnum,'') <> '' and crd1.AdresType = 'B' and ocrd.CardCode = '" & codiProveidor & "' "


        Dim ls As String = " select isnull(U_SEICliVox,'') as Clivox  from ocrd  "
        ls = ls & "  where   ocrd.CardCode = '" & codiProveidor & "' "

        oTmpRecordset.DoQuery(ls)

        If Not IsNothing(oTmpRecordset.Fields.Item(0)) Then
            If oTmpRecordset.Fields.Item(0).Value <> "" Then
                DonaClientVoxel = oTmpRecordset.Fields.Item("Clivox").Value
            Else
                DonaClientVoxel = ""
            End If
        Else
            DonaClientVoxel = ""
        End If

        SEI_Globals.LiberarObjCOM(oTmpRecordset)
    End Function


    Public Function donaUnitatMesura(ByVal codiUnitat As String) As String
        Dim oTmpRecordset As SAPbobsCOM.Recordset


        Select Case codiUnitat
            Case Is = ""
                Return ""
            Case Is = "Unidades"
                Return "Un"
            Case Is = "Kgs"
                Return "Kgs"
            Case Is = "Lts"
                Return "Lts"
            Case Is = "Lbs"
                Return "Lbs"
            Case Is = "Cajas"
                Return "Cajas"
            Case Is = "Bultos"
                Return "Bultos"
            Case Is = "Palets"
                Return "Palets"
            Case Is = "Horas"
                Return "Horas"
            Case Is = "Metros"
                Return "Metros"
            Case Is = "MetrosCuadrados"
                Return "MetrosCuadrados"
            Case Is = "Contenedores"
                Return "Contenedores"
            Case Is = "Otros"
                Return ""
            Case Else
                Return ""
        End Select


        SEI_Globals.LiberarObjCOM(oTmpRecordset)
    End Function

    Private Function DonaProveidorSAPBO(ByVal codiProveidorVoxel As String) As String
        Dim oTmpRecordset As SAPbobsCOM.Recordset
        oTmpRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim ls As String = "   Select ocrd.cardcode "
        ls = ls & " from crd1 "
        ls = ls & " left join ocrd on  ocrd.cardcode = crd1.cardcode  and crd1.AdresType = 'B'   "
        ls = ls & "where  isnull(crd1.glbllocnum,'') = '" & codiProveidorVoxel & "'"
        oTmpRecordset.DoQuery(ls)

        If oTmpRecordset.Fields.Item(0).Value > 0 Then
            DonaProveidorSAPBO = oTmpRecordset.Fields.Item("cardcode").Value
        Else
            DonaProveidorSAPBO = ""
        End If
        SEI_Globals.LiberarObjCOM(oTmpRecordset)
    End Function

    Private Function DonaEnviamentSAPBO(ByVal codiProveidorVoxel As String) As String
        Dim oTmpRecordset As SAPbobsCOM.Recordset
        oTmpRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
         

        Dim ls As String = "  Select Address"
        ls = ls & "      from crd1 "
        ls = ls & "    where  isnull(crd1.glbllocnum,'') = '" & codiProveidorVoxel & "' and crd1.AdresType = 'S'"
        oTmpRecordset.DoQuery(ls)

        If oTmpRecordset.Fields.Item(0).Value > 0 Then
            DonaEnviamentSAPBO = oTmpRecordset.Fields.Item("Address").Value
        Else
            DonaEnviamentSAPBO = ""
        End If
        SEI_Globals.LiberarObjCOM(oTmpRecordset)
    End Function

End Module
