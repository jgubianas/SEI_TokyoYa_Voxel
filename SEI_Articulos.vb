'
Option Explicit On
'
Imports System.Data.SqlClient
Imports System.Text
Imports System.IO
Imports SAPbobsCOM.BoObjectTypes

Public Class SEI_Articulos
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
    Public Sub Replicar_CLIENTES_TLY()

        Dim ls As String
        Dim oSqlcomand As SqlCommand
        Dim oDataReader As SqlClient.SqlDataReader = Nothing
        '
        ls = ""
        ls = ls & " Select [ID]"
        ls = ls & " ,[Create_Date]"
        ls = ls & " ,[Create_User]"
        ls = ls & " ,[Modify_Date]"
        ls = ls & " ,[Modify_User]"
        ls = ls & " ,[status_mobile]"
        ls = ls & " ,[Code]"
        ls = ls & " ,[Email]"
        ls = ls & " ,[Fax]"
        ls = ls & " ,[Phone]"
        ls = ls & " ,[Phone2]"
        ls = ls & " ,[Movil]"
        ls = ls & " FROM [CLIENTS_TLY] "
        '
        Try
            oSqlcomand = New SqlCommand(ls, go_conn)
            oDataReader = oSqlcomand.ExecuteReader()

            While oDataReader.Read()

                Me.Form.lblmsg.Text = "Cliente: " & oDataReader("Code").ToString
                Application.DoEvents()

                Select Case oDataReader("status_mobile").ToString
                    Case "01"
                        ' ALTA
                    Case "02"
                        ' MODIFICACIÓN
                        UPDATE_SBO_CLIENTE(oDataReader)

                End Select

            End While

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
#End Region

End Class
