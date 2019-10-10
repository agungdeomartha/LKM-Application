Public Class LoginForm

    Private Sub UserBindingNavigatorSaveItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Validate()
        Me.UserBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(Me.DB_APLIKASIDataSet)
    End Sub
    Sub KosongkanData()
        UsernameTextBox.Text = ""
        PasswordTextBox.Text = ""
    End Sub
    Private Sub LoginForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'DB_APLIKASIDataSet.User' table. You can move, or remove it, as needed.
        Me.UserTableAdapter.Fill(Me.DB_APLIKASIDataSet.User)
        Call KosongkanData()
    End Sub
    
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim dtamin = UserTableAdapter.LoginQuery(UsernameTextBox.Text, PasswordTextBox.Text, "admin")
        Dim dtkaryawan = UserTableAdapter.LoginQuery(UsernameTextBox.Text, PasswordTextBox.Text, "karyawan")
        If dtamin = 1 Then
            Me.Close()
            Form1.Show()
            MenuUtama.Hide()
        ElseIf dtkaryawan = 1 Then
            Me.Close()
            Form2.Show()
            MenuUtama.Hide()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class