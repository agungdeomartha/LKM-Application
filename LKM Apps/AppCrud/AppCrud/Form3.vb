Imports System.Data.SqlClient
Public Class Form3
    Public Sub New()
        ' required by form designer
        InitializeComponent()
        ' inisialisasi datetimepicker
        InitializeDatePicker()
        ' inisialisas  i combobox jenis transaksi
        InitializeComboBox()
        ' refresh grid data
        RefreshGrid()
    End Sub

    Private Sub InitializeDatePicker()
        ' set value date time picker ke awal dan akhir tahun
        DateTimePicker1.Value = "2014-01-01"
        DateTimePicker2.Value = "2014-12-31"
    End Sub

    Private Sub InitializeComboBox()
        Call buka()
        ' siapkan koneksi database
        ' siapkan data adapter untuk data retrieval
        Da = New SqlDataAdapter("select * from OrderTypes", Conn)
        ' siapkan datatable untuk menampung data dari database
        Dt = New DataTable
        ' enclose di dalam try-catch block
        ' untuk menghindari crash jika terjadi kesalahan database
        Try
            ' ambil data dari database
            Da.Fill(Dt)
            ' bind data ke combobox
            ComboBox1.DataSource = Dt
            ComboBox1.ValueMember = "ID"
            ComboBox1.DisplayMember = "Description"
            ' DONE!!!
        Catch ex As Exception
            ' tampilkan pesan error
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub RefreshGrid()
        Call buka()
        ' siapkan koneksi database
        ' siapkan data adapter untuk data retrieval
        Da = New SqlDataAdapter("SELECT A.*, B.Description AS TypeName " & _
                                     "FROM Orders A JOIN OrderTypes B ON A.TypeID=B.ID " & _
                                     "WHERE OrderNum LIKE @p1 AND TypeID = @t1 " & _
                                     "AND OrderDate BETWEEN @d1 AND @d2 ", Conn)
        Da.SelectCommand.Parameters.AddWithValue("@p1", "%" & TextBox1.Text & "%")
        Da.SelectCommand.Parameters.AddWithValue("@t1", ComboBox1.SelectedValue)
        Da.SelectCommand.Parameters.AddWithValue("@d1", DateTimePicker1.Value.ToString("yyyy-MM-dd"))
        Da.SelectCommand.Parameters.AddWithValue("@d2", DateTimePicker2.Value.ToString("yyyy-MM-dd"))
        ' siapkan datatable untuk menampung data dari database
        Dt = New DataTable
        ' enclose di dalam try-catch block
        ' untuk menghindari crash jika terjadi kesalahan database
        Try
            ' ambil data dari database
            Da.Fill(Dt)
            ' bind data ke combobox
            DataGridView1.DataSource = Dt
            ' DONE!!!
        Catch ex As Exception
            ' tampilkan pesan error
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click
        ' refresh grid data
        RefreshGrid()
    End Sub


    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class