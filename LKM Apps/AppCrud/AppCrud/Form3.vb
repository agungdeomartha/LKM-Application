Imports System.Data.SqlClient

Public Class Form3
    Private Const _cnString As String = "data source=Computer_Server;initial catalog=DB_APLIKASI;integrated security=true"

    Public Sub New()
        ' required by form designer
        InitializeComponent()
        ' inisialisasi datetimepicker
        InitializeDatePicker()
        ' inisialisasi combobox jenis transaksi
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
        ' siapkan koneksi database
        Dim cn As New SqlConnection(_cnString)
        ' siapkan data adapter untuk data retrieval
        Dim da As New SqlDataAdapter("select * from OrderTypes", cn)
        ' siapkan datatable untuk menampung data dari database
        Dim dt As New DataTable
        ' enclose di dalam try-catch block
        ' untuk menghindari crash jika terjadi kesalahan database
        Try
            ' ambil data dari database
            da.Fill(dt)
            ' bind data ke combobox
            ComboBox1.DataSource = dt
            ComboBox1.ValueMember = "ID"
            ComboBox1.DisplayMember = "Description"
            ' DONE!!!
        Catch ex As Exception
            ' tampilkan pesan error
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub RefreshGrid()
        ' siapkan koneksi database
        Dim cn As New SqlConnection(_cnString)
        ' siapkan data adapter untuk data retrieval
        Dim da As New SqlDataAdapter("SELECT A.*, B.Description AS TypeName " & _
                                     "FROM Orders A JOIN OrderTypes B ON A.TypeID=B.ID " & _
                                     "WHERE OrderNum LIKE @p1 AND TypeID = @t1 " & _
                                     "AND OrderDate BETWEEN @d1 AND @d2 ", cn)
        da.SelectCommand.Parameters.AddWithValue("@p1", "%" & TextBox1.Text & "%")
        da.SelectCommand.Parameters.AddWithValue("@t1", ComboBox1.SelectedValue)
        da.SelectCommand.Parameters.AddWithValue("@d1", DateTimePicker1.Value.ToString("yyyy-MM-dd"))
        da.SelectCommand.Parameters.AddWithValue("@d2", DateTimePicker2.Value.ToString("yyyy-MM-dd"))
        ' siapkan datatable untuk menampung data dari database
        Dim dt As New DataTable
        ' enclose di dalam try-catch block
        ' untuk menghindari crash jika terjadi kesalahan database
        Try
            ' ambil data dari database
            da.Fill(dt)
            ' bind data ke combobox
            DataGridView1.DataSource = dt
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