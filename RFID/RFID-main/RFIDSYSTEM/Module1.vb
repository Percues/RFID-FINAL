Imports System.Data.OleDb
Module Database
    Public cn As New OleDbConnection
    Public cmd As OleDbCommand
    Public sql As String
    Public dr As OleDbDataReader

    Public Sub connection()
        cn.Close()
        cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\RFIDSYSTEMFOLDER\RFID\RFID-main\RFIDSYSTEM\bin\Debug\RFIDdatabase.mdb"
        cn.Open()
        MsgBox("Connection success", MsgBoxStyle.Information, "Database Connection")
    End Sub
End Module