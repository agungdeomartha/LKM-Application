Imports System.Data.SqlClient

Public Class Form1
    Dim Conn As SqlConnection
    Dim Da As SqlDataAdapter
    Dim Cmd As SqlCommand
    Dim Rd As SqlDataReader
    Dim Ds As DataSet
    Dim MyDB As String
    Sub Koneksi()
        'data source=Computer_Server;initial catalog=DB_APLIKASI;integrated security=true
        'Data Source=192.168.100.6,1433;Network Library=DBMSSOCN;Initial Catalog=DB_APLIKASI;User ID=sa;Password=ilyvm;
        MyDB = "Data Source=192.168.100.6,1433;Network Library=DBMSSOCN;Initial Catalog=DB_APLIKASI;User ID=sa;Password=ilyvm;"
        Conn = New SqlConnection(MyDB)
        If Conn.State = ConnectionState.Closed Then Conn.Open()
    End Sub
    Sub KosongkanData()
        TextBox1.Text = ""
        TextBox2.Text = ""
        ComboBox1.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
    End Sub
    Sub KondisiAwal()
        Call Koneksi()
        Da = New SqlDataAdapter("Select * from TBL_MAHASISWA order by NIM desc", Conn)
        Ds = New DataSet
        Ds.Clear()
        Da.Fill(Ds, "TBL_MAHASISWA")
        DataGridView1.DataSource = (Ds.Tables("TBL_MAHASISWA"))
        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("PRIA")
        ComboBox1.Items.Add("WANITA")
        Call KosongkanData()
        TextBox1.MaxLength = 17
        TextBox2.MaxLength = 50
        TextBox3.MaxLength = 100
        TextBox4.MaxLength = 20

        Button1.Text = "INPUT"
        Button2.Text = "EDIT"
        Button3.Text = "DELETE"
        Button4.Text = "TUTUP"
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        ComboBox1.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False

    End Sub
    Sub SiapIsi()
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        ComboBox1.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
    End Sub
    Sub SiapHapus()
        TextBox1.Enabled = True
        TextBox2.Enabled = False
        ComboBox1.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call KondisiAwal()
        Me.KeyPreview = True
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        If Button4.Text = "TUTUP" Then
            Me.Close()
            MenuUtama.Show()
        Else
            Call KondisiAwal()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Button1.Text = "INPUT" Then
            Button1.Text = "SIMPAN"
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Text = "BATAL"
            Call SiapIsi()
            
        Else
            If TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or ComboBox1.Text = "" Then
                MsgBox("Data Belum Lengkap!, Silahkan isi semua Field")
            Else
                Call Koneksi()
                Cmd = New SqlCommand("Select * from TBL_MAHASISWA where NIM in (select max(NIM) from TBL_MAHASISWA)", Conn)
                Dim urutan As String
                Dim hitung As Long
                Dim MyDateTime As DateTime = Now()
                Dim MyString As String
                MyString = MyDateTime.ToString("yyyy/MM/")
                Rd = Cmd.ExecuteReader
                Rd.Read()

                If Not Rd.HasRows Then
                    urutan = "NIM" + MyString + "000001"
                Else
                    hitung = Microsoft.VisualBasic.Right(Rd.GetString(0), 3) + 1
                    urutan = "NIM" + MyString + Microsoft.VisualBasic.Right("000000" & hitung, 6)

                    Dim InputData As String = "insert into TBL_MAHASISWA values ('" & urutan & "','" & TextBox2.Text & "','" & ComboBox1.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "')"
                    Rd.Close()
                    Cmd = New SqlCommand(InputData, Conn)
                    Cmd.ExecuteNonQuery()
                    MsgBox("Data Berhasil Diinput")
                    Call KondisiAwal()
                End If
            End If
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Button2.Text = "EDIT" Then
            Button2.Text = "SIMPAN"
            Button1.Enabled = False
            Button3.Enabled = False
            Button4.Text = "BATAL"
            Call SiapIsi()

        Else
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or ComboBox1.Text = "" Then
                MsgBox("Data Belum Lengkap!, Silahkan isi semua Field")
            Else
                Call Koneksi()
                Dim EditData As String = "Update TBL_MAHASISWA set NamaMahasiswa='" & TextBox2.Text & "', JenisKelamin='" & ComboBox1.Text & "', AlamatMahasiswa='" & TextBox3.Text & "', TelpMahasiswa='" & TextBox4.Text & "' where NIM='" & TextBox1.Text & "'"
                Cmd = New SqlCommand(EditData, Conn)
                Cmd.ExecuteNonQuery()
                MsgBox("Data Berhasil DiEdit")
                Call KondisiAwal()
            End If
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If Button3.Text = "DELETE" Then
            Button3.Text = "HAPUS"
            Button1.Enabled = False
            Button2.Enabled = False
            Button4.Text = "BATAL"
            Call SiapHapus()
        Else
            If TextBox1.Text = "" Then
                MsgBox("Data Belum Lengkap!, Silahkan isi semua Field")
            Else
                Call Koneksi()
                Dim HapusData As String = "Delete TBL_MAHASISWA where NIM='" & TextBox1.Text & "'"
                Cmd = New SqlCommand(HapusData, Conn)
                Cmd.ExecuteNonQuery()
                MsgBox("Data Berhasil DiHapus")
                Call KondisiAwal()
            End If
        End If
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            Cmd = New SqlCommand("Select * From TBL_MAHASISWA Where NIM='" & TextBox1.Text & "'", Conn)
            Rd = Cmd.ExecuteReader
            Rd.Read()
            If Rd.HasRows Then
                TextBox2.Text = Rd.Item("NamaMahasiswa")
                ComboBox1.Text = Rd.Item("JenisKelamin")
                TextBox3.Text = Rd.Item("AlamatMahasiswa")
                TextBox4.Text = Rd.Item("TelpMahasiswa")
            Else
                MsgBox("Data Tidak Ada")
            End If
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Call Koneksi()
        Dim HapusData As String = "Delete TBL_MAHASISWA where NIM='" & TextBox1.Text & "'"
        Cmd = New SqlCommand(HapusData, Conn)
        Cmd.ExecuteNonQuery()
        MsgBox("Data Berhasil DiHapus")
        Call KondisiAwal()
    End Sub
    Private Sub LogoutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogoutToolStripMenuItem.Click
        Me.Hide()
        MenuUtama.Show()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Close()
    End Sub

    Private Sub DataGridView1_CellMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        If e.RowIndex >= 0 Then
            TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(0).Value
            TextBox2.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value
            ComboBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(2).Value
            TextBox3.Text = DataGridView1.Rows(e.RowIndex).Cells(3).Value
            TextBox4.Text = DataGridView1.Rows(e.RowIndex).Cells(4).Value
        End If
    End Sub
    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        'jika ke-4 button dibawah ini aktif
        If Button1.Enabled = True And Button2.Enabled = True And Button3.Enabled = True And Button4.Enabled = True Then
            'maka shorcut delete akan aktif
            Select Case e.KeyCode
                Case Keys.Delete
                    Call Koneksi()
                    Dim HapusData As String = "Delete TBL_MAHASISWA where NIM='" & TextBox1.Text & "'"
                    Cmd = New SqlCommand(HapusData, Conn)
                    Cmd.ExecuteNonQuery()
                    MsgBox("Data Berhasil DiHapus")
                    Call KondisiAwal()

            End Select
        End If
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Clipboard.SetText(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        TextBox1.Text = My.Computer.Clipboard.GetText()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Call KondisiAwal()
    End Sub
End Class
