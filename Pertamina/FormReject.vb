Imports System.Data.OleDb
Imports MySql.Data.MySqlClient
Imports System.Net

Public Class FormReject

    Dim connect As MySqlConnection
    Dim command As MySqlCommand
    Dim provider As String
    Dim dataFile As String
    Dim connString As String
    Dim simpan As String
    Dim tmpAbsensi, tmpDate, tmpTime As String
    Dim countAbsensi As Integer
    Dim tmpCount As Integer
    Dim selectDataBase As String

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btn_login.Click

       
    End Sub

    Private Sub FormReject_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
