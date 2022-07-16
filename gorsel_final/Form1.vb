Imports System.Data.OleDb

Public Class Form1
    Dim vtbag As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Application.StartupPath + "\\gorsel_final.mdb")
    Dim da As New OleDbDataAdapter("SELECT * FROM hesap", vtbag)
    Dim sorgu As New OleDbCommand()
    Dim oku As OleDbDataReader


    Dim Sonuc As New hesapla()

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Label1.Text = "Ürün ID: "
        Label2.Text = "Ürün Adı: "
        Label3.Text = "Ortalama Sabit Maliyet: "
        Label4.Text = "Ortalama Değişken Maliyet: "
        Label5.Text = "Birim Sayısı: "
        Button1.Text = "EKLE"
        Button2.Text = "GÜNCELLE"
        Button3.Text = "SİL"

        listele()

    End Sub
    Private Sub listele()
        Dim verial As New DataSet()
        da.Fill(verial, "sanaltablo")
        DataGridView1.DataSource = verial
        DataGridView1.DataMember = verial.Tables(0).TableName
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            vtbag.Open()
            sorgu.Connection = vtbag
            sorgu.CommandType = CommandType.Text
            sorgu.CommandText = "INSERT INTO hesap (urun_ad, ortalama_sabit_maliyet, ortalama_degisken_maliyet, birim_sayisi, toplam_maliyet) VALUES ('" + TextBox2.Text + "','" + TextBox3.Text + "','" + TextBox4.Text + "','" + TextBox5.Text + "','" + hesaplama()
            sorgu.ExecuteNonQuery()
            vtbag.Close()
            listele()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim urun_id As Integer
        urun_id = InputBox("Silmek istediğiniz ürün id giriniz:")
        While oku.Read()
            If oku("urun_id") = urun_id Then
                vtbag.Open()
                sorgu.Connection = vtbag
                sorgu.CommandType = CommandType.Text
                sorgu.CommandText = "DELETE FROM hesap WHERE urun_id='" + urun_id + "'"
                sorgu.ExecuteNonQuery()
                vtbag.Close()
            End If
            MessageBox.Show(urun_id + " NUMARALI ÜRÜN SİLİNMİŞTİR.")
            listele()
        End While
    End Sub
End Class
