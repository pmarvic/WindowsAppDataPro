Imports System
Imports System.Windows.Forms
Imports System.Data
Imports MySql.Data
Imports MySql.Data.Types
Imports MySql.Data.MySqlClient
Imports System.Globalization

Public Class ClassCnnMysql
    Implements IDisposable
    Public Connection As MySqlConnection
    Public Command As MySqlCommand
    Public cnnMYSQL As New ADODB.Connection
    Public cmdMYSQL As New ADODB.Command
    Public rstMYSQL As New ADODB.Recordset


    Public Sub CnnOdbcMysql(Server As String, User As String, Clave As String, BD As String, Drivers As String)
        If Drivers = "3.51" Then
            cnnMYSQL.ConnectionString = "driver={mysql odbc 3.51 driver};" _
                         & "user=" & User & ";" _
                         & "password=" & Clave & ";" _
                         & "server=" + Server + ";" _
                         & "database=" & BD & ";"
        Else
            cnnMYSQL.ConnectionString = "DSN=datapro"
        End If

        Try
            cnnMYSQL.Open()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Sub CloseOdbcMyql()
        cnnMYSQL.Close()
    End Sub

    Public Sub CnnNetMysql()

        Try
            Connection = New MySqlConnection()
            Connection.ConnectionString = "server=127.0.0.1;" _
            & "uid=root;" _
            & "pwd=s3rv3r;" _
            & "port=3306;" _
            & "database=vscyrm;"

            Connection.Open()

        Catch ex As Exception
            MsgBox("Error al conectar al servidor MySQL " &
        vbCrLf & vbCrLf & ex.Message,
        MsgBoxStyle.OkOnly + MsgBoxStyle.Critical)
        Finally
            '   Connection.Dispose()
        End Try

    End Sub
    Public Sub CloseNetMysql()
        Connection.Close()
    End Sub
    Protected Overridable Overloads Sub Dispose(disposing As Boolean)
        If disposing Then
            ' dispose managed resources
            Connection.Close()
        End If
    End Sub 'Dispose


    Public Overloads Sub Dispose() Implements IDisposable.Dispose

        Dispose(True)
        GC.SuppressFinalize(Me)

    End Sub 'Dispose



End Class
