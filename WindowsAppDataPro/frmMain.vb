Imports System
Imports System.Windows.Forms
Imports System.Data
Public Class frmMain
    Public ObjDBF As New ClassCnnDBase
    Public ObjMysql As New ClassCnnMysql
    Public cmdMYSQL As New ADODB.Command
    Public rstMYSQL As New ADODB.Recordset

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ObjDBF.Ruta = My.Settings.RutaDBF.ToString
        ObjDBF.CnnOpenDBF(ObjDBF.Ruta)
        ObjMysql.CnnOpenMysql(My.Settings.Mysql_IP, My.Settings.Mysql_User, My.Settings.Mysql_Pass, My.Settings.Mysql_DB, My.Settings.Mysql_Version)
        BrowserPrrUpdate()
    End Sub
    Sub BrowserPrrUpdate()
        Dim rec As ADODB.Recordset
        Dim Conn As New ADODB.Connection
        Dim sql As String
        Try
            Conn = ObjMysql.cnnMYSQL
            rec = New ADODB.Recordset
            sql = " SELECT DPINV.INV_CODIGO "
            sql = sql + "FROM DPINV "
            sql = sql + " WHERE DPINV.INV_FCHACT BETWEEN DATE_SUB(CURRENT_DATE(),INTERVAL 1 DAY) AND CURRENT_DATE() "
            sql = sql + " ORDER BY DPINV.INV_CODIGO "


            rec.Open(sql, Conn)
            Do While Not (rec.EOF)
                Console.WriteLine(rec("INV_CODIGO").Value)
                rec.MoveNext()
            Loop

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
End Class