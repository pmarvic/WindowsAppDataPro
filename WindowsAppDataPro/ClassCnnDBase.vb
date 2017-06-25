Imports System
Imports System.Windows.Forms
Imports System.Data
Public Class ClassCnnDBase
    Public cnnDBF As New ADODB.Connection
    Public cmdDBF As New ADODB.Command
    Public rstDBF As New ADODB.Recordset
    Public Ruta As String
    Public File As String
    Public Function CnnOpenDBF(cRutaCaja As String) As ADODB.Connection
        Dim FileName As String = "SIPRR.DBF"
        Try
            If IO.File.Exists(FileName) Then
                cnnDBF.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cRutaCaja & ";" & "Extended Properties=""DBASE IV;"";"
                cnnDBF.Open()



            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Return cnnDBF
    End Function
    Sub CloseDBF()
        cnnDBF.Close()
    End Sub
End Class
