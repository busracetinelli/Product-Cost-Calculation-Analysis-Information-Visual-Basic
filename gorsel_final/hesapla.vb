Public Class hesapla
    Public Function hesaplama(ByVal ortalama_sabit_maliyet As Integer, ByVal ortalama_degisken_maliyet As Integer, ByVal birim_sayisi As Integer)
        Dim Sonuc As Double
        Sonuc = (ortalama_sabit_maliyet + ortalama_degisken_maliyet) * birim_sayisi
        Return Sonuc
    End Function
End Class
