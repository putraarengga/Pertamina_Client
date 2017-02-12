Imports System.Data
Imports System.Data.Odbc
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports MySql.Data.MySqlClient

Module MdlKoneksi
    Public konek As OdbcConnection
    Public DA As OdbcDataAdapter
    Public DS As DataSet
    Public DT As DataTable
    Public DR As OdbcDataReader
    Public cmd As OdbcCommand
    Sub bukaDB()
        Try
            konek = New OdbcConnection("Dsn=" + FormMenu.arrValue(0) + ";server=" + FormMenu.arrValue(1) + ";userid=admin_idsf;password=123456;database=idsf;port=3306")
            'konek = New OdbcConnection("Dsn=idsf;server=db4free.net;userid=idsf_putra;password=ansel3128;database=idsf;port=3306")
            'konek = New OdbcConnection("Dsn=idsf;server=localhost;userid=root;password=r7pqv6s6Xc9QbZKK;database=idsf;port=3306")
            If konek.State = ConnectionState.Closed Then
                konek.Open()
            End If
        Catch ex As Exception
            MsgBox("Koneksi DataBase Bermasalah, Silahkan Periksa Koneksi Anda!")
        End Try
    End Sub
End Module
