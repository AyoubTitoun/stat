Imports System.Data.OleDb
Imports Microsoft.Office.Interop


Public Class Form_s
    Public connection As New System.Data.OleDb.OleDbConnection("provider=Microsoft.Ace.OLEDB.12.0;" & "data source=" & Application.StartupPath & "\DataBase\DataBase.accdb")
    Dim da As New OleDbDataAdapter
    Dim da1 As New OleDbDataAdapter
    Dim table1 As New System.Data.DataTable
    Dim table2 As New System.Data.DataTable






    Sub Data_search()

        table1.Clear()
        table2.Clear()

        da = New OleDbDataAdapter("SELECT * FROM [MOYENNES] where Option like '" & ComboBox1.Text & "' AND [NIVEAU] like '" & ComboBox2.Text & "' AND [ANNEE_SCOLAIRE] like '" & TextBox3.Text & "'", connection)

        da1 = New OleDbDataAdapter("SELECT * FROM [RATRAPAGE] where OPTION like '" & ComboBox1.Text & "' AND [NIVEAU] like '" & ComboBox2.Text & "' AND [ANNE] like '" & TextBox3.Text & "'", connection)


        da.Fill(table1)
        da1.Fill(table2)
        DataGridView3.DataSource = table1
        DataGridView4.DataSource = table2


        Dim i As Integer

        Label6.Text = DataGridView3.RowCount()
        Label11.Text = DataGridView4.RowCount()
        i = (((DataGridView4.RowCount() * 100) / DataGridView3.RowCount()))
        BunifuCircleProgressbar2.Value = i
    End Sub







    Private Sub Button1_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Data_search()
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Form_s_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub PictureBox7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub Label4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click

    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click

    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub BunifuImageButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BunifuImageButton8.Click
        Me.Close()
    End Sub

    Private Sub BunifuImageButton9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BunifuImageButton9.Click
        Me.Hide()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Data_search()
    End Sub
End Class
