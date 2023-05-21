Imports System.Data.OleDb
Imports System.IO.File
Imports System.IO.FileStream
Imports Microsoft.VisualBasic.Devices
Imports System.Net
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Reflection.Emit
Imports System.IO
Public Class Form1
    Dim Conn As OleDbConnection

    Private Sub Populate()
        sql = "Select RFID, NAME, LRN, GRADE, DEPARTMENT, IMAGE from StudentInfo"
        cmd = New OleDbCommand(sql, cn)
        Dim da As New OleDbDataAdapter(cmd)
        Dim dt As New DataTable
        da.Fill(dt)
        DBTABLE.DataSource = dt
        DBTABLE2.DataSource = dt
    End Sub

    Private Sub ButtonNew_Click(sender As Object, e As EventArgs)
        StudentInfoBindingSource.AddNew()
    End Sub

    Private Sub ButtonSave_Click(sender As Object, e As EventArgs) Handles ButtonSave.Click
        sql = "INSERT INTO [StudentInfo] (RFID, [NAME], LRN, [GRADE], DEPARTMENT, [IMAGE]) Values('" & TextBoxRFID.Text & "','" & TextBoxName.Text & "','" & TextBoxLRN.Text & "','" & TextBoxGrade.Text & "','" & TextBoxDepartment.Text & "', '" & TextBoxImage.Text & "')"
        cmd = New OleDbCommand(sql, cn)
        cmd.ExecuteNonQuery()
        MsgBox("Information Successfully Saved", MsgBoxStyle.Information, "Student Information")

        Call Populate()
    End Sub

    Private Sub ButtonRemove_Click(sender As Object, e As EventArgs) Handles ButtonRemove.Click
        If DBTABLE.SelectedRows.Count > 0 Then
            Dim selectedRowIndex As Integer = DBTABLE.SelectedRows(0).Index
            Dim rfidValue As String = DBTABLE.Rows(selectedRowIndex).Cells(0).Value.ToString()

            Dim result As DialogResult = MessageBox.Show("Are you sure you want to delete this record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                Dim commandText As String = "DELETE FROM [StudentInfo] WHERE RFID = @RFID"
                cmd = New OleDbCommand(commandText, cn)
                cmd.Parameters.AddWithValue("@RFID", rfidValue)
                cmd.ExecuteNonQuery()
                MsgBox("Information Successfully Deleted", MsgBoxStyle.Information, "Student Information")
                Call Populate()
            End If
        Else
            MessageBox.Show("No record selected to delete.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub ButtonBrowseImage_Click(sender As Object, e As EventArgs) Handles ButtonBrowseImage.Click
        OpenFileDialog1.ShowDialog()
        TextBoxImage.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub TextBoxImage_TextChanged(sender As Object, e As EventArgs) Handles TextBoxImage.TextChanged
        If (System.IO.File.Exists(TextBoxImage.Text)) Then
            PictureBoxImageInput.Image = Image.FromFile(TextBoxImage.Text)
        End If
        If TextBoxImage.Text = "" Then
            PictureBoxImageInput.Hide()
        Else
            PictureBoxImageInput.Show()
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call connection()
        Call Populate()
    End Sub

    Private Sub ButtonConnection_Click(sender As Object, e As EventArgs)
        PanelUserData.Visible = False
        PanelRegistrationEditUserData.Visible = False
        PanelMasterlist.Visible = False
        PanelConnection.Visible = True
    End Sub

    Private Sub ButtonUserData_Click(sender As Object, e As EventArgs) Handles ButtonUserData.Click
        PanelConnection.Visible = False
        PanelRegistrationEditUserData.Visible = False
        PanelMasterlist.Visible = False
        PanelUserData.Visible = True
    End Sub

    Private Sub ButtonRegistration_Click(sender As Object, e As EventArgs) Handles ButtonRegistration.Click
        PanelConnection.Visible = False
        PanelUserData.Visible = False
        PanelMasterlist.Visible = False
        PanelRegistrationEditUserData.Visible = True
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        PanelConnection.Visible = False
        PanelUserData.Visible = False
        PanelRegistrationEditUserData.Visible = False
        PanelMasterlist.Visible = True
    End Sub

    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        GroupBoxImage.Location = New Point((PanelUserData.Width / 2) - (GroupBoxImage.Width / 2), GroupBoxImage.Top)
    End Sub

    Private Sub PanelRegistrationEditUserData_Paint(sender As Object, e As PaintEventArgs) Handles PanelRegistrationEditUserData.Paint
        e.Graphics.DrawRectangle(New Pen(Color.LightGray, 2), PanelConnection.ClientRectangle)
    End Sub

    Private Sub PanelRegistrationEditUserData_Resize(sender As Object, e As EventArgs) Handles PanelRegistrationEditUserData.Resize
        PanelRegistrationEditUserData.Invalidate()
    End Sub

    Private Sub ButtonClearForm_Click(sender As Object, e As EventArgs) Handles ButtonClearForm.Click
        TextBoxRFID.Clear()
        TextBoxName.Clear()
        TextBoxLRN.Clear()
        TextBoxGrade.Clear()
        TextBoxDepartment.Clear()
        TextBoxImage.Clear()
    End Sub
    Private Sub DBTABLE_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DBTABLE.CellClick
        TextBoxRFID.Text = DBTABLE.Rows(e.RowIndex).Cells(0).Value.ToString
        TextBoxName.Text = DBTABLE.Rows(e.RowIndex).Cells(1).Value.ToString
        TextBoxLRN.Text = DBTABLE.Rows(e.RowIndex).Cells(2).Value.ToString
        TextBoxGrade.Text = DBTABLE.Rows(e.RowIndex).Cells(3).Value.ToString
        TextBoxDepartment.Text = DBTABLE.Rows(e.RowIndex).Cells(4).Value.ToString
        TextBoxImage.Text = DBTABLE.Rows(e.RowIndex).Cells(5).Value.ToString
    End Sub

    Private Sub TextBoxImage2_TextChanged(sender As Object, e As EventArgs) Handles TextBoxImage2.TextChanged
        If (System.IO.File.Exists(TextBoxImage2.Text)) Then
            PictureBoxImageInput2.Image = Image.FromFile(TextBoxImage2.Text)
        End If
        If TextBoxImage2.Text = "" Then
            PictureBoxImageInput2.Hide()
        Else
            PictureBoxImageInput2.Show()
        End If
    End Sub

    Private Sub DBTABLE2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DBTABLE2.CellClick
        TextBoxRFID2.Text = DBTABLE2.Rows(e.RowIndex).Cells(0).Value.ToString
        TextBoxName2.Text = DBTABLE2.Rows(e.RowIndex).Cells(1).Value.ToString
        TextBoxLRN2.Text = DBTABLE2.Rows(e.RowIndex).Cells(2).Value.ToString
        TextBoxGrade2.Text = DBTABLE2.Rows(e.RowIndex).Cells(3).Value.ToString
        TextBoxDepartment2.Text = DBTABLE2.Rows(e.RowIndex).Cells(4).Value.ToString
        TextBoxImage2.Text = DBTABLE2.Rows(e.RowIndex).Cells(5).Value.ToString
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBoxSearchBar.TextChanged
        sql = "Select RFID,NAME,LRN,GRADE,DEPARTMENT from StudentInfo where [RFID] Like '%" & TextBoxSearchBar.Text & "%' Or [NAME] Like'%" & TextBoxSearchBar.Text & "%' Or [LRN] Like'%" & TextBoxSearchBar.Text & "%'"
        cmd = New OleDbCommand(sql, cn)
        Dim da As New OleDbDataAdapter(cmd)
        Dim dt As New DataTable
        da.Fill(dt)
        DBTABLE2.DataSource = dt

        If TextBoxSearchBar.Text = "" Then
            Call Populate()
        End If
    End Sub

    Private Sub TextBoxSearch_TextChanged(sender As Object, e As EventArgs) Handles TextBoxSearch.TextChanged
        sql = "SELECT RFID, NAME, LRN, GRADE, DEPARTMENT, IMAGE FROM StudentInfo WHERE [RFID] LIKE '%" & TextBoxSearch.Text & "%' OR [NAME] LIKE '%" & TextBoxSearch.Text & "%' OR [LRN] LIKE '%" & TextBoxSearch.Text & "%'"
        cmd = New OleDbCommand(sql, cn)
        dr = cmd.ExecuteReader()

        If dr.HasRows Then
            dr.Read()
            LabelName.Text = " " & dr("NAME").ToString()
            LabelLRN.Text = " " & dr("LRN").ToString()
            LabelGrade.Text = " " & dr("GRADE").ToString()
            LabelDepartment.Text = " " & dr("DEPARTMENT").ToString()

            Dim imagePath As String = dr("IMAGE").ToString()
            If Not String.IsNullOrEmpty(imagePath) AndAlso File.Exists(imagePath) Then
                PictureBoxUserImage.Image = Image.FromFile(imagePath)
            Else
                PictureBoxUserImage.Image = Nothing
            End If
        Else
            LabelName.Text = "Waiting..."
            LabelLRN.Text = "Waiting..."
            LabelGrade.Text = "Waiting..."
            LabelDepartment.Text = "Waiting..."
            PictureBoxUserImage.Image = Nothing
        End If

        dr.Close()
    End Sub
    Private Sub ButtonClear_Click(sender As Object, e As EventArgs) Handles ButtonClear.Click
        LabelName.Text = "Waiting..."
        LabelLRN.Text = "Waiting..."
        LabelGrade.Text = "Waiting..."
        LabelDepartment.Text = "Waiting..."
        PictureBoxUserImage.Image = Nothing
        TextBoxSearch.Clear()
    End Sub



End Class