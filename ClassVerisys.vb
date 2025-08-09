Imports System.Data
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Imports System.Runtime.InteropServices
Imports Microsoft.Reporting.WinForms
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.Office.Interop.Excel
Imports System.Drawing
Public Class ClassVerisys
    Private connDBverisys As SqlConnection
    Private cmdVerisys As SqlCommand
    Private readerVerisys As SqlDataReader
    Private cmdLoc As SqlCommand
    Private readerLoc As SqlDataReader
    Public Sub openConn()
        connDBverisys = New SqlConnection("Server=" + ServerName + ";Database=DB2023VERS;Trusted_Connection=True;")
        'connDBverisys.Open()
        If connDBverisys.State = ConnectionState.Closed Then
            connDBverisys.Open()
        End If
    End Sub
    Public Sub insertLocation(ByVal location As String)
        Dim insertQuery As String = "insert into TB2023LOC (Location) VALUE(" + location + ")"
        Dim recordExists As Boolean = CheckRecordExists("TB2023LOC", "Location", location)

        If recordExists Then
            ' Record exists
            MsgBox("Record already exists.")
        Else
            ' Record does not exist

            cmdLoc = New SqlCommand(insertQuery, connDBverisys)
            openConn()
            Try
                cmdLoc.ExecuteNonQuery()
                connDBverisys.Close()
            Catch ex As Exception
                MsgBox("Error during inserting location" + vbCrLf + Err.Description, MsgBoxStyle.Information)
                connDBverisys.Close()
            End Try
        End If
        

    End Sub

    Public Sub UpdateLocation()

    End Sub
    Public Sub deleteLocation()

    End Sub
    Public Sub insertVerisys()

    End Sub
    Public Sub updateVerisys()

    End Sub
    Public Sub deleteVerisys()

    End Sub
    Public Function CheckRecordExists(tableName As String, columnName As String, value As Object) As Boolean
        Dim recordExists As Boolean = False

        ' Create the SQL connection
        openConn()

        ' Create the SQL command
        Dim commandText As String = "SELECT COUNT(*) FROM " + tableName + " WHERE " + columnName + " = @Value"
        Using command As New SqlCommand(commandText, connDBverisys)
            ' Add the parameter
            command.Parameters.AddWithValue("@Value", value)

            ' Execute the command and get the count
            Dim count As Integer = CInt(command.ExecuteScalar())
            connDBverisys.Close()
            command.Dispose()
            ' Check if the count is greater than 0
            If count > 0 Then
                recordExists = True
            End If
        End Using


        Return recordExists
    End Function
End Class
