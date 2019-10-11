Imports System.Data.Sql
Imports System.Data.SqlClient
Module Module1
    Public Conn As SqlConnection
    Public Ds As DataSet
    Public Rd As SqlDataReader
    Public Da As SqlDataAdapter
    Public Cmd As SqlCommand
    Public Dt As DataTable
    Public sql As String

    Public Sub buka()

        sql = "Data Source=192.168.100.11,1433;Network Library=DBMSSOCN;Initial Catalog=DB_APLIKASI;User ID=sa;Password=ilyvm;"
        Conn = New SqlConnection(sql)
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub buka2()

        sql = "Data Source=192.168.100.11,1433;Network Library=DBMSSOCN;Initial Catalog=D:\LKMBENANGEMBROMISDATA\TRANSAKSI.MDF;User ID=sa;Password=ilyvm;"
        Conn = New SqlConnection(sql)
        Try
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Module