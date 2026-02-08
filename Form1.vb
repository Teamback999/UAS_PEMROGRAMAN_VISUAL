Imports System.IO

Public Class FrmStokBarang
    Dim modeEdit As Boolean = False
    Dim indexEdit As Integer = -1


    Private Sub FrmStokBarang_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dgvStok.ReadOnly = False
        dgvStok.AllowUserToAddRows = False
        dgvStok.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvStok.MultiSelect = False
    End Sub

    ' =========================
    ' SIMPAN DATA BARU
    ' =========================
    Private Sub btnSimpan_Click(sender As Object, e As EventArgs) Handles btnSimpan.Click
        Dim kode As String = txtKode.Text
        Dim nama As String = txtNama.Text
        Dim jumlah As String = txtJumlah.Text
        Dim harga As String = txtHarga.Text

        If kode = "" Or nama = "" Or jumlah = "" Or harga = "" Then
            MessageBox.Show("Lengkapi semua data!")
            Exit Sub
        End If

        If Not IsNumeric(jumlah) Or Not IsNumeric(harga) Then
            MessageBox.Show("Jumlah dan Harga harus angka!")
            Exit Sub
        End If

        Dim baris As String = kode & "|" & nama & "|" & jumlah & "|" & harga
        If modeEdit = False Then
            File.AppendAllText("stok_barang.txt", baris & Environment.NewLine)
        Else
            dgvStok.Rows(indexEdit).Cells(0).Value = kode
            dgvStok.Rows(indexEdit).Cells(1).Value = nama
            dgvStok.Rows(indexEdit).Cells(2).Value = jumlah
            dgvStok.Rows(indexEdit).Cells(3).Value = harga

            SimpanKeFile()

            modeEdit = False
            indexEdit = -1

            MessageBox.Show("Data berhasil diperbarui")
            btnBrowse.PerformClick()
            Exit Sub
        End If


        MessageBox.Show("Data berhasil disimpan")

        txtKode.Clear()
        txtNama.Clear()
        txtJumlah.Clear()
        txtHarga.Clear()

        btnBrowse.PerformClick()
    End Sub

    ' =========================
    ' TAMPILKAN DATA
    ' =========================
    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        dgvStok.Rows.Clear()

        If File.Exists("stok_barang.txt") = False Then
            MessageBox.Show("File data belum ada")
            Exit Sub
        End If

        Dim lines() As String = File.ReadAllLines("stok_barang.txt")

        For Each line As String In lines
            If line.Trim() = "" Then Continue For

            Dim data() As String = line.Split("|"c)

            If data.Length >= 4 Then
                dgvStok.Rows.Add(data(0), data(1), data(2), data(3))
            End If
        Next
    End Sub

    ' =========================
    ' EDIT LANGSUNG DI GRID
    ' =========================
    Private Sub dgvStok_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgvStok.CellEndEdit
        SimpanKeFile()
    End Sub

    ' =========================
    ' VALIDASI ANGKA GRID
    ' =========================
    Private Sub dgvStok_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles dgvStok.CellValidating
        If e.ColumnIndex = 2 Or e.ColumnIndex = 3 Then
            If Not IsNumeric(e.FormattedValue.ToString()) Then
                MessageBox.Show("Jumlah dan Harga harus angka!")
                e.Cancel = True
            End If
        End If
    End Sub

    ' =========================
    ' SIMPAN SEMUA DATA KE FILE
    ' =========================
    Private Sub SimpanKeFile()
        Dim hasil As New List(Of String)

        For Each row As DataGridViewRow In dgvStok.Rows
            If row.IsNewRow Then Continue For

            Dim kode As String = row.Cells(0).Value?.ToString()
            Dim nama As String = row.Cells(1).Value?.ToString()
            Dim jumlah As String = row.Cells(2).Value?.ToString()
            Dim harga As String = row.Cells(3).Value?.ToString()

            If kode <> "" Then
                hasil.Add(kode & "|" & nama & "|" & jumlah & "|" & harga)
            End If
        Next

        File.WriteAllLines("stok_barang.txt", hasil)
    End Sub

    ' =========================
    ' HAPUS DATA BARIS
    ' =========================
    Private Sub btnHapus_Click(sender As Object, e As EventArgs) Handles btnHapus.Click

        If dgvStok.SelectedRows.Count = 0 Then
            MessageBox.Show("Pilih baris yang ingin dihapus!")
            Exit Sub
        End If

        Dim konfirmasi = MessageBox.Show("Yakin ingin menghapus data ini?",
                                         "Konfirmasi",
                                         MessageBoxButtons.YesNo,
                                         MessageBoxIcon.Warning)

        If konfirmasi = DialogResult.Yes Then
            dgvStok.Rows.Remove(dgvStok.SelectedRows(0))
            SimpanKeFile()
            MessageBox.Show("Data berhasil dihapus")
        End If

    End Sub

    ' =========================
    ' CARI DATA
    ' =========================
    Private Sub btnCari_Click(sender As Object, e As EventArgs) Handles btnCari.Click
        Dim keyword As String = txtCari.Text.ToLower()

        For Each row As DataGridViewRow In dgvStok.Rows
            row.Visible = False

            If row.IsNewRow Then Continue For

            For i = 0 To row.Cells.Count - 1
                If row.Cells(i).Value IsNot Nothing AndAlso
                   row.Cells(i).Value.ToString().ToLower().Contains(keyword) Then
                    row.Visible = True
                    Exit For
                End If
            Next
        Next

    End Sub

    ' =========================
    ' LAPORAN STOK
    ' =========================
    Private Sub btnLaporan_Click(sender As Object, e As EventArgs) Handles btnLaporan.Click

        If File.Exists("stok_barang.txt") = False Then
            MessageBox.Show("Data belum tersedia")
            Exit Sub
        End If

        Dim totalJenis As Integer = 0
        Dim totalJumlah As Integer = 0
        Dim totalNilai As Double = 0

        Dim lines() As String = File.ReadAllLines("stok_barang.txt")
        totalJenis = lines.Length

        For Each line As String In lines
            If line.Trim() = "" Then Continue For

            Dim data() As String = line.Split("|"c)

            If data.Length < 4 Then Continue For

            If IsNumeric(data(2)) And IsNumeric(data(3)) Then
                totalJumlah += CInt(data(2))
                totalNilai += CInt(data(2)) * CDbl(data(3))
            End If
        Next

        txtLaporan.Text =
            "LAPORAN STOK BARANG" & vbCrLf &
            "----------------------------" & vbCrLf &
            "Total Jenis Barang : " & totalJenis & vbCrLf &
            "Total Jumlah Barang: " & totalJumlah & vbCrLf &
            "Total Nilai Stok   : Rp " & totalNilai.ToString("N0")

    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click

        If dgvStok.SelectedRows.Count = 0 Then
            MessageBox.Show("Pilih data yang ingin diedit!")
            Exit Sub
        End If

        Dim row As DataGridViewRow = dgvStok.SelectedRows(0)

        txtKode.Text = row.Cells(0).Value.ToString()
        txtNama.Text = row.Cells(1).Value.ToString()
        txtJumlah.Text = row.Cells(2).Value.ToString()
        txtHarga.Text = row.Cells(3).Value.ToString()

        modeEdit = True
        indexEdit = row.Index

        MessageBox.Show("Data siap diedit, silakan ubah lalu klik SIMPAN")

    End Sub

End Class
