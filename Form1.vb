Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Form1
    Private Structure DataBarang
        Dim Nama As String
        Dim Harga As String
        Dim Jumlah As String
        Dim Diskon As String
        Dim TotalBayar As String
    End Structure
    Dim Indeks As Integer
    Dim Data() As DataBarang
    Sub DaftarBarang()
        Data(Indeks).Nama = TextBox2.Text
        Data(Indeks).Harga = TextBox3.Text
        Data(Indeks).Jumlah = TextBox4.Text
        Data(Indeks).Diskon = TextBox5.Text
        Data(Indeks).TotalBayar = TextBox6.Text
    End Sub

    Sub HapusDaftar()
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox2.Focus()
    End Sub
    Sub TampilanData()
        MsgBox("Informasi barang" & Indeks & Chr(10) &
        "Nama Barang         : " & Data(Indeks).Nama & Chr(10) &
        "Harga                   : " & Data(Indeks).Harga & Chr(10) &
        "Jumlah                  : " & Data(Indeks).Jumlah & Chr(10) &
        "Diskon                  : " & Data(Indeks).Diskon & Chr(10) &
        "Totalbayar              : " & Data(Indeks).TotalBayar, , "Data Barang")
    End Sub
    Sub TampilanFrom()
        TextBox2.Text = Data(Indeks).Nama
        TextBox3.Text = Data(Indeks).Harga
        TextBox4.Text = Data(Indeks).Jumlah
        TextBox5.Text = Data(Indeks).Diskon
        TextBox6.Text = Data(Indeks).TotalBayar
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Indeks = 1
        ReDim Data(Indeks)
        TextBox1.Text = Indeks
        TextBox2.Focus()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        HapusDaftar()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim totalHarga As Double
        Dim totalDiskon As Double

        ' hitung total harga
        totalHarga = Double.Parse(TextBox3.Text) * Double.Parse(TextBox4.Text)

        totalDiskon = Double.Parse(TextBox3.Text) * Double.Parse(TextBox4.Text) * Double.Parse(TextBox5.Text) / 100
        ' hitung total bayar
        TextBox6.Text = totalHarga - totalDiskon
        DaftarBarang()
        TampilanData()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Indeks > LBound(Data) Then
            DaftarBarang()
            Indeks = Indeks - 1
            TampilanFrom()
        End If
        If Indeks = 0 Then Indeks = 1
        TextBox1.Text = Indeks
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim totalHarga As Double
        Dim totalDiskon As Double

        ' hitung total harga
        totalHarga = Double.Parse(TextBox3.Text) * Double.Parse(TextBox4.Text)

        totalDiskon = Double.Parse(TextBox3.Text) * Double.Parse(TextBox4.Text) * Double.Parse(TextBox5.Text) / 100

        ' hitung total bayar
        TextBox6.Text = totalHarga - totalDiskon
        TextBox1.Text = Indeks
        If Indeks = UBound(Data) Then
            ReDim Preserve Data(Indeks + 1)
        End If
        DaftarBarang()
        Indeks = Indeks + 1
        TextBox1.Text = Indeks
        TampilanFrom()
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Close()
    End Sub
End Class
