Imports System.IO

Public Class FrmStokBarang
    Dim filePath As String = ""


    Private Sub FrmStokBarang_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Kosong dulu
    End Sub

    Private Sub btnSimpan_Click(sender As Object, e As EventArgs) Handles btnSimpan.Click
        Dim kode As String = txtKode.Text
        Dim nama As String = txtNama.Text
        Dim jumlah As String = txtJumlah.Text
        Dim harga As String = txtHarga.Text

        Dim baris As String = kode & "|" & nama & "|" & jumlah & "|" & harga

        File.AppendAllText("stok_barang.txt", baris & Environment.NewLine)

        MessageBox.Show("Data berhasil disimpan ke file")

        txtKode.Clear()
        txtNama.Clear()
        txtJumlah.Clear()
        txtHarga.Clear()

        btnBrowse.PerformClick()
    End Sub

    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click

        dgvStok.Rows.Clear()

        If File.Exists("stok_barang.txt") = False Then
            MessageBox.Show("File data belum ada")
            Exit Sub
        End If

        Dim lines() As String = File.ReadAllLines("stok_barang.txt")

        For Each line As String In lines
            Dim data() As String = line.Split("|"c)
            dgvStok.Rows.Add(data(0), data(1), data(2), data(3))
        Next

    End Sub

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
            "-------------------------" & vbCrLf &
            "Total Jenis Barang : " & totalJenis & vbCrLf &
            "Total Jumlah Barang: " & totalJumlah & vbCrLf &
            "Total Nilai Stok   : Rp " & totalNilai.ToString("N0")

    End Sub

End Class
