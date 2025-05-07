Attribute VB_Name = "SEI_Functions"
Option Explicit
'
Public Sub Descuentos_Documento(sInterlocutor As String, sFecha As String, _
                                ByRef oApplication As SAPbouiCOM.Application, ByRef oMatrix As SAPbouiCOM.Matrix, _
                                iTipoDocumento As TIPOVENTASoCOMPRAS)
    '
    Dim oDescuento As DESCUENTOS
    Dim oLinees() As String
    Dim iFila As Long
    Dim bAplicarLinea As Boolean
    Dim bAplicarDocumento As Boolean
    Const lArticulo As String = "1"
    Const lCodigoDescuentoDocumento As String = "U_SEIDescD"
    Const lCodigoDescuentoLinea As String = "U_SEIDescL"
    Const lDescuento1 As String = "U_SEIDesc1"
    Const lDescuento2 As String = "U_SEIDesc2"
    Const lDescuento3 As String = "U_SEIDesc3"
    Const lDescuento4 As String = "U_SEIDesc4"
    Const lDescuento5 As String = "U_SEIDesc5"
    Const lDescuento As String = "15"
    Const lPrecio As String = "17"
    '
    'Comprovo que no hi hagi variables en blanc
    If sInterlocutor = "" Then
        oApplication.StatusBar.SetText "Debes seleccionar un cliente para poder calcular el descuento", bmt_Short, smt_Error
        Exit Sub
    End If
    If sFecha = "" Then
        oApplication.StatusBar.SetText "Debes seleccionar una fecha para poder calcular el descuento", bmt_Short, smt_Error
        Exit Sub
    End If
    'Comprovo que les columnes necessàries estiguin visibles i editables
    If (Not oMatrix.Columns(lPrecio).Visible) Or (Not oMatrix.Columns(lPrecio).Editable) Then
        oApplication.StatusBar.SetText "Para calcular el descuento, la columna del 'Precio' debe estar visible y editable", bmt_Short, smt_Error
        Exit Sub
    End If
    If (Not oMatrix.Columns(lDescuento).Visible) Or (Not oMatrix.Columns(lDescuento).Editable) Then
        oApplication.StatusBar.SetText "Para calcular el descuento, la columna del '% de descuento' debe estar visible y editable", bmt_Short, smt_Error
        Exit Sub
    End If
    If (Not oMatrix.Columns(lCodigoDescuentoDocumento).Visible) Or (Not oMatrix.Columns(lCodigoDescuentoDocumento).Editable) Then
        oApplication.StatusBar.SetText "Para calcular el descuento, la columna del 'Código Descuento Documento' debe estar visible y editable", bmt_Short, smt_Error
        Exit Sub
    End If
    If (Not oMatrix.Columns(lDescuento1).Visible) Or (Not oMatrix.Columns(lDescuento1).Editable) Then
        oApplication.StatusBar.SetText "Para calcular el descuento, la columna del 'Descuento 1' debe estar visible y editable", bmt_Short, smt_Error
        Exit Sub
    End If
    If (Not oMatrix.Columns(lDescuento2).Visible) Or (Not oMatrix.Columns(lDescuento2).Editable) Then
        oApplication.StatusBar.SetText "Para calcular el descuento, la columna del 'Descuento 2' debe estar visible y editable", bmt_Short, smt_Error
        Exit Sub
    End If
    If (Not oMatrix.Columns(lDescuento3).Visible) Or (Not oMatrix.Columns(lDescuento3).Editable) Then
        oApplication.StatusBar.SetText "Para calcular el descuento, la columna del 'Descuento 3' debe estar visible y editable", bmt_Short, smt_Error
        Exit Sub
    End If
    If (Not oMatrix.Columns(lDescuento4).Visible) Or (Not oMatrix.Columns(lDescuento4).Editable) Then
        oApplication.StatusBar.SetText "Para calcular el descuento, la columna del 'Descuento 4' debe estar visible y editable", bmt_Short, smt_Error
        Exit Sub
    End If
    If (Not oMatrix.Columns(lDescuento5).Visible) Or (Not oMatrix.Columns(lDescuento5).Editable) Then
        oApplication.StatusBar.SetText "Para calcular el descuento, la columna del 'Descuento 5' debe estar visible y editable", bmt_Short, smt_Error
        Exit Sub
    End If
    '
    'En el cas de compres demano si volen calcular els descomptes per document
    ' Eliminar validacion 02-10-2006 Esther Dorca
    'If iTipoDocumento = TCompras Then
    '    If oApplication.MessageBox("Quieres calcular los descuentos por documento?", 1, "Sí", "No") = 2 Then
    '        Exit Sub
    '    End If
    'End If
    '
    bAplicarLinea = False
    bAplicarDocumento = False
    '
    Set oDescuento = Nothing
    Set oDescuento = New DESCUENTOS
    oLinees = oDescuento.DESCUENTO_DOCUMENTO(sInterlocutor, sFecha, iTipoDocumento, TUserInterface, oMatrix)
    For iFila = 1 To UBound(oLinees)
        If oLinees(iFila, 0) <> "" Then
            'Si el codi de descompte de linea coincideix amb el de descompte per document no haig de guardar res
            'D'això se'n fa responsable en GABRI
            If NullToText(oLinees(iFila, 0)) <> NullToText(oMatrix.Columns(lCodigoDescuentoLinea).Cells(iFila).Specific.Value) Then
                bAplicarLinea = False
                If Not bAplicarDocumento Then
                    Select Case oApplication.MessageBox("Quieres aplicar bonificación por cantidad en el artículo '" & _
                                oMatrix.Columns(lArticulo).Cells(iFila).Specific.Value & "' de la línea " & iFila & "?", 1, "Sí", "Sí a Todos", "No")
                        Case Is = 1
                            bAplicarLinea = True
                        Case Is = 2
                            bAplicarDocumento = True
                        Case Is = 3
                            bAplicarLinea = False
                            bAplicarDocumento = False
                    End Select
                End If
                If bAplicarLinea Or bAplicarDocumento Then
                    oMatrix.Columns(lDescuento1).Cells(iFila).Specific.Value = Replace(oLinees(iFila, 1), ",", ".")
                    oMatrix.Columns(lDescuento2).Cells(iFila).Specific.Value = Replace(oLinees(iFila, 2), ",", ".")
                    oMatrix.Columns(lDescuento3).Cells(iFila).Specific.Value = Replace(oLinees(iFila, 3), ",", ".")
                    oMatrix.Columns(lDescuento4).Cells(iFila).Specific.Value = Replace(oLinees(iFila, 4), ",", ".")
                    oMatrix.Columns(lDescuento5).Cells(iFila).Specific.Value = Replace(oLinees(iFila, 5), ",", ".")
                    oMatrix.Columns(lCodigoDescuentoDocumento).Cells(iFila).Specific.Value = oLinees(iFila, 0)
                End If
            End If
        End If
    Next iFila
    '
    Set oDescuento = Nothing
    '
End Sub
'
Public Sub Descuentos_Linea(sInterlocutor As String, sArticulo As String, sFecha As String, dCantidad As Double, _
                                    ByRef oApplication As SAPbouiCOM.Application, ByRef oMatrix As SAPbouiCOM.Matrix, _
                                    iFila As Long, iTipoDocumento As TIPOVENTASoCOMPRAS)
    '
    Dim oDescuento As DESCUENTOS
    Dim iRegistro As Long
    Dim ls As String
    Dim oRecordset As SAPbobsCOM.Recordset
    Const lCodigoDescuento As String = "U_SEIDescL"
    Const lDescuento1 As String = "U_SEIDesc1"
    Const lDescuento2 As String = "U_SEIDesc2"
    Const lDescuento3 As String = "U_SEIDesc3"
    Const lDescuento4 As String = "U_SEIDesc4"
    Const lDescuento5 As String = "U_SEIDesc5"
    Const lDescuento As String = "15"
    Const lPrecio As String = "17"
    Const lPrecioNeto As String = "U_SEIPrice"
    Const lListaMateriales As String = "39"
    '
    'Comprovo que no hi hagi variables en blanc
    If sInterlocutor = "" Then
        oApplication.StatusBar.SetText "Debes seleccionar un cliente para poder calcular el descuento", bmt_Short, smt_Error
        Exit Sub
    End If
    If sArticulo = "" Then
        oApplication.StatusBar.SetText "Debes seleccionar un artículo para poder calcular el descuento", bmt_Short, smt_Error
        Exit Sub
    End If
    If sFecha = "" Then
        oApplication.StatusBar.SetText "Debes seleccionar una fecha para poder calcular el descuento", bmt_Short, smt_Error
        Exit Sub
    End If
    If dCantidad <= 0 Then
        oApplication.StatusBar.SetText "La cantidad debe ser superior a 0 para poder calcular el descuento", bmt_Short, smt_Error
        Exit Sub
    End If
    'Comprovo que les columnes necessàries estiguin visibles i editables
    If (Not oMatrix.Columns(lPrecio).Visible) Or (Not oMatrix.Columns(lPrecio).Editable) Then
        oApplication.StatusBar.SetText "Para calcular el descuento, la columna del 'Precio' debe estar visible y editable", bmt_Short, smt_Error
        Exit Sub
    End If
    If (Not oMatrix.Columns(lDescuento).Visible) Or (Not oMatrix.Columns(lDescuento).Editable) Then
        oApplication.StatusBar.SetText "Para calcular el descuento, la columna del '% de descuento' debe estar visible y editable", bmt_Short, smt_Error
        Exit Sub
    End If
    If (Not oMatrix.Columns(lCodigoDescuento).Visible) Or (Not oMatrix.Columns(lCodigoDescuento).Editable) Then
        oApplication.StatusBar.SetText "Para calcular el descuento, la columna del 'Código Descuento' debe estar visible y editable", bmt_Short, smt_Error
        Exit Sub
    End If
    If (Not oMatrix.Columns(lDescuento1).Visible) Or (Not oMatrix.Columns(lDescuento1).Editable) Then
        oApplication.StatusBar.SetText "Para calcular el descuento, la columna del 'Descuento 1' debe estar visible y editable", bmt_Short, smt_Error
        Exit Sub
    End If
    If (Not oMatrix.Columns(lDescuento2).Visible) Or (Not oMatrix.Columns(lDescuento2).Editable) Then
        oApplication.StatusBar.SetText "Para calcular el descuento, la columna del 'Descuento 2' debe estar visible y editable", bmt_Short, smt_Error
        Exit Sub
    End If
    If (Not oMatrix.Columns(lDescuento3).Visible) Or (Not oMatrix.Columns(lDescuento3).Editable) Then
        oApplication.StatusBar.SetText "Para calcular el descuento, la columna del 'Descuento 3' debe estar visible y editable", bmt_Short, smt_Error
        Exit Sub
    End If
    If (Not oMatrix.Columns(lDescuento4).Visible) Or (Not oMatrix.Columns(lDescuento4).Editable) Then
        oApplication.StatusBar.SetText "Para calcular el descuento, la columna del 'Descuento 4' debe estar visible y editable", bmt_Short, smt_Error
        Exit Sub
    End If
    If (Not oMatrix.Columns(lDescuento5).Visible) Or (Not oMatrix.Columns(lDescuento5).Editable) Then
        oApplication.StatusBar.SetText "Para calcular el descuento, la columna del 'Descuento 5' debe estar visible y editable", bmt_Short, smt_Error
        Exit Sub
    End If
    '
    'Comprovo que no sigui un fill d'una llista de materials
    If oMatrix.Columns(lListaMateriales).Cells(iFila).Specific.Selected.Value = "I" Then
        oMatrix.Columns(lDescuento1).Cells(iFila).Specific.Value = 0
        oMatrix.Columns(lDescuento2).Cells(iFila).Specific.Value = 0
        oMatrix.Columns(lDescuento3).Cells(iFila).Specific.Value = 0
        oMatrix.Columns(lDescuento4).Cells(iFila).Specific.Value = 0
        oMatrix.Columns(lDescuento5).Cells(iFila).Specific.Value = 0
        oMatrix.Columns(lCodigoDescuento).Cells(iFila).Specific.Value = ""
    Else
        Set oDescuento = Nothing
        Set oDescuento = New DESCUENTOS
        iRegistro = oDescuento.DESCUENTO_LINIA(sArticulo, sInterlocutor, sFecha, dCantidad, iTipoDocumento)
        If iTipoDocumento = TCompras Then
            ls = "SELECT * FROM [@SEIDESCUENTOSPRO] WHERE DocEntry = " & iRegistro
        ElseIf iTipoDocumento = TVentas Then
            ls = "SELECT * FROM [@SEIDESCUENTOSCLI] WHERE DocEntry = " & iRegistro
        End If
        Set oRecordset = Nothing
        Set oRecordset = oCompany.GetBusinessObject(BoRecordset)
        oRecordset.DoQuery ls
        If oRecordset.EOF Then
        
            'possible millora rendiment pantalla
            If ("" <> Trim(oMatrix.Columns(lCodigoDescuento).Cells(iFila).Specific.Value)) Or _
               (NullToDoble(oMatrix.Columns(lDescuento1).Cells(iFila).Specific.String) <> 0) Or _
               (NullToDoble(oMatrix.Columns(lDescuento2).Cells(iFila).Specific.String) <> 0) Or _
               (NullToDoble(oMatrix.Columns(lDescuento3).Cells(iFila).Specific.String) <> 0) Or _
               (NullToDoble(oMatrix.Columns(lDescuento4).Cells(iFila).Specific.String) <> 0) Or _
               (NullToDoble(oMatrix.Columns(lDescuento5).Cells(iFila).Specific.String) <> 0) Then
        
                esdeveniment = True
                oMatrix.Columns(lDescuento1).Cells(iFila).Specific.Value = 0
                oMatrix.Columns(lDescuento2).Cells(iFila).Specific.Value = 0
                oMatrix.Columns(lDescuento3).Cells(iFila).Specific.Value = 0
                oMatrix.Columns(lDescuento4).Cells(iFila).Specific.Value = 0
                oMatrix.Columns(lDescuento5).Cells(iFila).Specific.Value = 0
                oMatrix.Columns(lCodigoDescuento).Cells(iFila).Specific.Value = ""
                esdeveniment = False

            End If
        Else
            If oRecordset(lPrecioNeto).Value = 0 Then 'Aplicar descomptes
            
                'possible millora rendiment pantalla
                If (Trim(oRecordset("Code").Value) <> Trim(oMatrix.Columns(lCodigoDescuento).Cells(iFila).Specific.Value)) Or _
                   (NullToDoble(oMatrix.Columns(lDescuento1).Cells(iFila).Specific.String) <> NullToDoble(oRecordset(lDescuento1).Value)) Or _
                   (NullToDoble(oMatrix.Columns(lDescuento2).Cells(iFila).Specific.String) <> NullToDoble(oRecordset(lDescuento2).Value)) Or _
                   (NullToDoble(oMatrix.Columns(lDescuento3).Cells(iFila).Specific.String) <> NullToDoble(oRecordset(lDescuento3).Value)) Or _
                   (NullToDoble(oMatrix.Columns(lDescuento4).Cells(iFila).Specific.String) <> NullToDoble(oRecordset(lDescuento4).Value)) Or _
                   (NullToDoble(oMatrix.Columns(lDescuento5).Cells(iFila).Specific.String) <> NullToDoble(oRecordset(lDescuento5).Value)) Then
                    
                    esdeveniment = True
                    oMatrix.Columns(lDescuento1).Cells(iFila).Specific.Value = Replace(oRecordset(lDescuento1).Value, ",", ".")
                    oMatrix.Columns(lDescuento2).Cells(iFila).Specific.Value = Replace(oRecordset(lDescuento2).Value, ",", ".")
                    oMatrix.Columns(lDescuento3).Cells(iFila).Specific.Value = Replace(oRecordset(lDescuento3).Value, ",", ".")
                    oMatrix.Columns(lDescuento4).Cells(iFila).Specific.Value = Replace(oRecordset(lDescuento4).Value, ",", ".")
                    oMatrix.Columns(lDescuento5).Cells(iFila).Specific.Value = Replace(oRecordset(lDescuento5).Value, ",", ".")
                    '
                    oMatrix.Columns(lCodigoDescuento).Cells(iFila).Specific.Value = oRecordset("Code").Value
                    esdeveniment = False
                End If
            Else 'Aplicar preu net
                esdeveniment = True
                oMatrix.Columns(lDescuento1).Cells(iFila).Specific.Value = 0
                oMatrix.Columns(lDescuento2).Cells(iFila).Specific.Value = 0
                oMatrix.Columns(lDescuento3).Cells(iFila).Specific.Value = 0
                oMatrix.Columns(lDescuento4).Cells(iFila).Specific.Value = 0
                oMatrix.Columns(lDescuento5).Cells(iFila).Specific.Value = 0
                oMatrix.Columns(lPrecio).Cells(iFila).Specific.Value = oRecordset(lPrecioNeto).Value
                oMatrix.Columns(lCodigoDescuento).Cells(iFila).Specific.Value = oRecordset("Code").Value
                esdeveniment = False
            End If
        End If
    End If
    '
    Set oDescuento = Nothing
    Set oRecordset = Nothing
    '
End Sub
'
Public Function Calculo_Descuento_Total(dDescuento1 As Double, dDescuento2 As Double, dDescuento3 As Double, _
                                    dDescuento4 As Double, dDescuento5 As Double) As Double
    '
    'A partir de 5 descomptes retorna el seu únic descompte equivalent
    '
    Dim lDescuento1 As Double
    Dim lDescuento2 As Double
    Dim lDescuento3 As Double
    Dim lDescuento4 As Double
    Dim lDescuento5 As Double
    Dim lDescuento As Double
    '
    'Calculo_Descuento_Total = dDescuentoActual
    '
    'If dDescuento1 = 0 And dDescuento2 = 0 And dDescuento3 = 0 And dDescuento4 = 0 And dDescuento5 = 0 Then Exit Function
    '
    lDescuento1 = 1 - (dDescuento1 / 100)
    lDescuento2 = 1 - (dDescuento2 / 100)
    lDescuento3 = 1 - (dDescuento3 / 100)
    lDescuento4 = 1 - (dDescuento4 / 100)
    lDescuento5 = 1 - (dDescuento5 / 100)
    lDescuento = (1 - (lDescuento1 * lDescuento2 * lDescuento3 * lDescuento4 * lDescuento5)) * 100
    '
    Calculo_Descuento_Total = lDescuento
    '
End Function
'
Public Function AdmiteOperacionesTraspaso(sInterlocutor As String) As Boolean
    '
    
    
    If (RecuperaValor("U_SEI015", "OCRD", Array("CardCode"), Array(Trim(sInterlocutor)), "") = "S") Then
       AdmiteOperacionesTraspaso = True
    Else
       AdmiteOperacionesTraspaso = False
    End If
    
    'Dim oInterlocutor As SAPbobsCOM.BusinessPartners
    '
    
    'Set oInterlocutor = Nothing
    'Set oInterlocutor = oCompany.GetBusinessObject(oBusinessPartners)
    'If oInterlocutor.GetByKey(sInterlocutor) Then
    '    If oInterlocutor.UserFields("U_SEI015").Value = "S" Then
    '        AdmiteOperacionesTraspaso = True
    '    Else
    '        AdmiteOperacionesTraspaso = False
    '    End If
    'Else
    '    AdmiteOperacionesTraspaso = False
    'End If
    ''
    'Set oInterlocutor = Nothing
    '
End Function
'
Public Function AveragePrice(sArticulo As String, sAlmacen As String) As Double
    '
    'Dim oItem As SAPbobsCOM.Items
    Dim ls As String
    Dim oRecordset As SAPbobsCOM.Recordset
    Dim i As Long
    '
    'Set oItem = Nothing
    'Set oItem = oCompany.GetBusinessObject(oItems)
    'If oItem.GetByKey(sArticulo) Then
        'For i = 0 To oItem.WhsInfo.Count - 1
            'oItem.WhsInfo.SetCurrentLine i
            'If oItem.WhsInfo.WarehouseCode = sAlmacen Then
                'AveragePrice = oItem.WhsInfo.StandardAveragePrice
                'Exit For
            'End If
        'Next i
        'If i >= oItem.WhsInfo.Count Then
            'AveragePrice = 0
        'End If
    'Else
        'AveragePrice = 0
    'End If
    '
    'ls = "SELECT * FROM OITW " & _
        " WHERE ItemCode = '" & sArticulo & "' " & _
        " AND WhsCode = '" & sAlmacen & "'"
    ls = "SELECT * FROM OITM WHERE ItemCode = '" & sArticulo & "'"
    Set oRecordset = Nothing
    Set oRecordset = oCompany.GetBusinessObject(BoRecordset)
    oRecordset.DoQuery ls
    If oRecordset.EOF Then
        AveragePrice = 0
    Else
        AveragePrice = oRecordset("AvgPrice").Value
    End If
    '
    'Set oItem = Nothing
    Set oRecordset = Nothing
    '
End Function
'
Public Function Usuario_Permiso_Almacen(sAlmacen As String) As Boolean
    '
    Dim ls As String
    Dim oRecordset As SAPbobsCOM.Recordset
    Const lAlmacen1 As String = "U_SEIAlm1"
    Const lAlmacen2 As String = "U_SEIAlm2"
    Const lAlmacen3 As String = "U_SEIAlm3"
    Const lAlmacen4 As String = "U_SEIAlm4"
    Const lAlmacen5 As String = "U_SEIAlm5"
    Const lAlmacen6 As String = "U_SEIAlm6"
    Const lAlmacen7 As String = "U_SEIAlm7"
    Const lAlmacen8 As String = "U_SEIAlm8"
    Const lAlmacen9 As String = "U_SEIAlm9"
    Const lAlmacen10 As String = "U_SEIAlm10"
    '
    Usuario_Permiso_Almacen = True
    '
    ls = "SELECT * FROM OUSR WHERE Internal_K = " & oCompany.UserSignature
    Set oRecordset = Nothing
    Set oRecordset = oCompany.GetBusinessObject(BoRecordset)
    oRecordset.DoQuery ls
    If Not oRecordset.EOF Then
        If UCase(sAlmacen) <> UCase(oRecordset(lAlmacen1).Value) _
        And UCase(sAlmacen) <> UCase(oRecordset(lAlmacen2).Value) _
        And UCase(sAlmacen) <> UCase(oRecordset(lAlmacen3).Value) _
        And UCase(sAlmacen) <> UCase(oRecordset(lAlmacen4).Value) _
        And UCase(sAlmacen) <> UCase(oRecordset(lAlmacen5).Value) _
        And UCase(sAlmacen) <> UCase(oRecordset(lAlmacen6).Value) _
        And UCase(sAlmacen) <> UCase(oRecordset(lAlmacen7).Value) _
        And UCase(sAlmacen) <> UCase(oRecordset(lAlmacen8).Value) _
        And UCase(sAlmacen) <> UCase(oRecordset(lAlmacen9).Value) _
        And UCase(sAlmacen) <> UCase(oRecordset(lAlmacen10).Value) Then
            Usuario_Permiso_Almacen = False
        End If
    End If
    '
    Set oRecordset = Nothing
    '
End Function
'
Public Function StockDisponible(sArticulo As String, sAlmacen As String) As Double
    '
    Dim ls As String
    Dim oRecordset As SAPbobsCOM.Recordset
    '
    ls = "SELECT (OnHand - IsCommited) AS Disponible FROM OITW " & _
        " WHERE ItemCode = '" & sArticulo & "' " & _
        " AND WhsCode = '" & sAlmacen & "'"
    Set oRecordset = Nothing
    Set oRecordset = oCompany.GetBusinessObject(BoRecordset)
    oRecordset.DoQuery ls
    If oRecordset.EOF Then
        StockDisponible = 0
    Else
        StockDisponible = oRecordset("Disponible").Value
    End If
    '
    Set oRecordset = Nothing
    '
End Function
'
Public Function AlmacenDefectoUsuario() As String
    '
    Dim ls As String
    Dim oRecordset As SAPbobsCOM.Recordset
    '
    ls = "SELECT WareHouse FROM OUDG " & _
        " WHERE Code = " & _
        " (SELECT DfltsGroup FROM OUSR WHERE Internal_K = " & oCompany.UserSignature & ")"
    Set oRecordset = Nothing
    Set oRecordset = oCompany.GetBusinessObject(BoRecordset)
    oRecordset.DoQuery ls
    If oRecordset.EOF Then
        AlmacenDefectoUsuario = ""
    Else
        AlmacenDefectoUsuario = oRecordset("WareHouse").Value
    End If
    '
    Set oRecordset = Nothing
    '
End Function
'
Public Function TienePermisoDocumentosCredito() As Boolean
    '
    Dim ls As String
    Dim oRecordset As SAPbobsCOM.Recordset
    '
    ls = "SELECT * FROM OUSR WHERE Internal_K = " & oCompany.UserSignature
    Set oRecordset = Nothing
    Set oRecordset = oCompany.GetBusinessObject(BoRecordset)
    oRecordset.DoQuery ls
    If oRecordset.EOF Then
        TienePermisoDocumentosCredito = False
    Else
        If oRecordset("U_SEI001").Value = "S" Then
            TienePermisoDocumentosCredito = True
        Else
            TienePermisoDocumentosCredito = False
        End If
    End If
    '
    Set oRecordset = Nothing
    '
End Function
'
Public Function TienePermisoCancelarDocumentos() As Boolean
    '
    Dim ls As String
    Dim oRecordset As SAPbobsCOM.Recordset
    '
    ls = "SELECT * FROM OUSR WHERE Internal_K = " & oCompany.UserSignature
    Set oRecordset = Nothing
    Set oRecordset = oCompany.GetBusinessObject(BoRecordset)
    oRecordset.DoQuery ls
    If oRecordset.EOF Then
        TienePermisoCancelarDocumentos = False
    Else
        If oRecordset("U_SEI002").Value = "S" Then
            TienePermisoCancelarDocumentos = True
        Else
            TienePermisoCancelarDocumentos = False
        End If
    End If
    '
    Set oRecordset = Nothing
    '
End Function
'
Public Function PrecioMinimoVenta(sArticulo As String) As Double
    '
    Dim ls As String
    Dim oRecordset As SAPbobsCOM.Recordset
    '
    ls = "SELECT * FROM ITM1 " & _
        " WHERE ItemCode = '" & sArticulo & "' " & _
        " AND PriceList = 3"
    Set oRecordset = Nothing
    Set oRecordset = oCompany.GetBusinessObject(BoRecordset)
    oRecordset.DoQuery ls
    If oRecordset.EOF Then
        PrecioMinimoVenta = 0
    Else
        PrecioMinimoVenta = oRecordset("Price").Value
    End If
    '
    Set oRecordset = Nothing
    '
End Function
'
Public Sub RecalcularMargen(ByRef oApplication As SAPbouiCOM.Application, ByRef oMatrix As SAPbouiCOM.Matrix, ByRef pVal As SAPbouiCOM.ItemEvent)
    '
    Dim dPrecioNeto As Double
    Dim dPrecioCoste As Double
    Dim dMargen As Double
    Dim dPrecioBruto As Double
    Dim dValor As Double
    Const lPrecioNeto As String = "17"
    Const lPrecioBruto As String = "14"
    Const lPrecioCoste As String = "U_SEI010"
    Const lMargen As String = "U_SEI009"
    Const lDescuento1 As String = "U_SEIDesc1"
    Const lDescuento2 As String = "U_SEIDesc2"
    Const lDescuento3 As String = "U_SEIDesc3"
    Const lDescuento4 As String = "U_SEIDesc4"
    Const lDescuento5 As String = "U_SEIDesc5"
    Const lArticulo As String = "1"
    Const lListaMateriales As String = "39"
    '
    'Només aplicar-ho a les ofertes!!
    If pVal.FormTypeEx <> SEI_EventsOfertaVenta.OfertaVenta_FormType Then Exit Sub
    If pVal.Row = 0 Then Exit Sub
    If oMatrix.Columns(lListaMateriales).Cells(pVal.Row).Specific.Selected.Value = "I" Then Exit Sub
    '
    dPrecioNeto = NullToDoble(Format(Val(Replace(Replace(oMatrix.Columns(lPrecioNeto).Cells(pVal.Row).Specific.Value, ".", ""), ",", ".")), "0.000"))
    dPrecioBruto = NullToDoble(Format(Val(Replace(Replace(oMatrix.Columns(lPrecioBruto).Cells(pVal.Row).Specific.Value, ".", ""), ",", ".")), "0.000"))
    dMargen = NullToDoble(Format(Replace(oMatrix.Columns(lMargen).Cells(pVal.Row).Specific.Value, ".", ","), "0.000"))
    dPrecioCoste = NullToDoble(Format(Replace(oMatrix.Columns(lPrecioCoste).Cells(pVal.Row).Specific.Value, ".", ","), "0.000"))
    '
    If oMatrix.Columns(lArticulo).Cells(pVal.Row).Specific.Value <> "" Then
        'Recalcular descompte 1 per aplicar marge
        If (pVal.ColUID = lPrecioCoste Or pVal.ColUID = lMargen) And (pVal.EventType = et_DOUBLE_CLICK) Then
            If Format(dMargen, "0.000") <> -100 Then
                If Format(dPrecioNeto, "0.000") <> Format((1 + (dMargen / 100)) * dPrecioCoste, "0.000") Then
                    dPrecioNeto = Format((1 + (dMargen / 100)) * dPrecioCoste, "0.000")
                    dValor = Format(((dPrecioBruto - dPrecioNeto) / dPrecioBruto) * 100, "0.000")
                    If dValor > 100 Or dValor < -100 Then
                        oApplication.MessageBox "El descuento resultante supera el 100% o el -100%. Operación cancelada"
                    Else
                        oMatrix.Columns(lDescuento1).Cells(pVal.Row).Specific.Value = 0
                        oMatrix.Columns(lDescuento2).Cells(pVal.Row).Specific.Value = 0
                        oMatrix.Columns(lDescuento3).Cells(pVal.Row).Specific.Value = 0
                        oMatrix.Columns(lDescuento4).Cells(pVal.Row).Specific.Value = 0
                        oMatrix.Columns(lDescuento5).Cells(pVal.Row).Specific.Value = 0
                        oMatrix.Columns(lDescuento1).Cells(pVal.Row).Specific.Value = Replace(dValor, ",", ".")
                    End If
                End If
            End If
        Else
            If dPrecioCoste = 0 Then
                oMatrix.Columns(lMargen).Cells(pVal.Row).Specific.Value = 100
            Else
                If Format(dMargen, "0.000") <> Format((((dPrecioNeto - dPrecioCoste) / dPrecioCoste) * 100), "0.000") Then
                    dValor = Format((((dPrecioNeto - dPrecioCoste) / dPrecioCoste) * 100), "0.000")
                    oMatrix.Columns(lMargen).Cells(pVal.Row).Specific.Value = Replace(dValor, ",", ".")
                End If
            End If
        End If
    End If
    '
End Sub
'
Public Sub Imprimir_Informe(ByRef oApplication As SAPbouiCOM.Application, sNombreInforme As String, _
                            bCapturar As Boolean, sParametro As String, bImprimir As Boolean)
    '
    Dim oInformes As SAPbouiCOM.Form
    Dim oMatrix As SAPbouiCOM.Matrix
    Dim i As Long
    Const cInformesUsuarios As String = "4868"
    Const cTipoFormularioInforme As String = "4666"
    Const cTipoFormularioParametros As String = "4000"
    '
    'Obro el formulari d'informes d'usuari
    On Error GoTo CONTINUAR
    Set oInformes = oApplication.Forms.GetForm(cTipoFormularioInforme, 0)
    If Not oInformes Is Nothing Then oInformes.Close
CONTINUAR:
    oApplication.ActivateMenuItem cInformesUsuarios
    'Capturo el formulari i el poso als núvols
    If oApplication.Forms.ActiveForm.TypeEx = cTipoFormularioInforme Then
        Set oInformes = Nothing
        Set oInformes = oApplication.Forms.ActiveForm
        oInformes.Items("3").Click ct_Regular
        'oInformes.Top = -5000
    Else
        Exit Sub
    End If
    'Busco l'informe i el selecciono
    Set oMatrix = Nothing
    Set oMatrix = oInformes.Items("5").Specific
    For i = 1 To oMatrix.VisualRowCount
        If oMatrix.Columns("1").Cells(i).Specific.Value = sNombreInforme Then
            oMatrix.Columns("1").Cells(i).Click ct_Regular
            Exit For
        End If
    Next i
    If i > oMatrix.VisualRowCount Then
        oApplication.StatusBar.SetText "No se ha encontrado el informe '" & sNombreInforme & "'", bmt_Short, smt_Warning
        oInformes.Close
        Exit Sub
    End If
    'Crido al informe
    oInformes.Select
    '
    If bCapturar Then
        SEI_EventsParametros.Parametros_Capturar = True
        SEI_EventsParametros.Parametros_Valor1 = sParametro
    End If
    '
    If bImprimir Then
        If oApplication.Menus(mnu_Imprimir).Enabled Then oApplication.ActivateMenuItem mnu_Imprimir
    Else
        If oApplication.Menus(mnu_PresentacionPreliminar).Enabled Then oApplication.ActivateMenuItem mnu_PresentacionPreliminar
    End If
    '
    oInformes.Close
    '
    Set oMatrix = Nothing
    Set oInformes = Nothing
    '
End Sub

