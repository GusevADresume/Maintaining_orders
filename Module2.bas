Attribute VB_Name = "Module2"
Function coursive(n As Double) As String
 
 Dim Nums1, Nums2, Nums3, Nums4 As Variant
 
 Nums1 = Array("", "один ", "два ", "три ", "четыре ", "п€ть ", "шесть ", "семь ", "восемь ", "дев€ть ")
 Nums2 = Array("", "дес€ть ", "двадцать ", "тридцать ", "сорок ", "пятьдес€т ", "шестьдес€т ", "семьдес€т ", _
                        "восемьдес€т ", "дев€носто ")
 Nums3 = Array("", "сто ", "двести ", "триста ", "четыреста ", "пятьсот ", "шестьсот ", "семьсот ", _
                        "восемьсот ", "дев€тьсот ")
 Nums4 = Array("", "одна ", "две ", "три ", "четыре ", "пять ", "шесть ", "семь ", "восемь ", "девять ")
 Nums5 = Array("дес€ть ", "одиннадцать ", "двенадцать ", "тринадцать ", "четырнадцать ", _
                        "пятнадцать ", "шестнадцать ", "семнадцать ", "восемнадцать ", "дев€тнадцать ")
 
 If n <= 0 Then
   СУММАПРОПИСЬЮ = "ноль"
   Exit Function
 End If
 'разделяем число на разряды, используя вспомогательную функцию Class
 ed = Class(n, 1)
 dec = Class(n, 2)
 sot = Class(n, 3)
 tys = Class(n, 4)
 dectys = Class(n, 5)
 sottys = Class(n, 6)
 mil = Class(n, 7)
 decmil = Class(n, 8)
 
 'проверяем миллионы
 Select Case decmil
   Case 1
     mil_txt = Nums5(mil) & "миллионов "
     GoTo www
   Case 2 To 9
     decmil_txt = Nums2(decmil)
 End Select
 Select Case mil
   Case 1
     mil_txt = Nums1(mil) & "миллион "
   Case 2, 3, 4
     mil_txt = Nums1(mil) & "миллиона "
   Case 5 To 20
     mil_txt = Nums1(mil) & "миллионов "
 End Select
www:
 sottys_txt = Nums3(sottys)
 'проверяем тысячи
 Select Case dectys
   Case 1
     tys_txt = Nums5(tys) & "тыс€ч "
     GoTo eee
   Case 2 To 9
     dectys_txt = Nums2(dectys)
 End Select
 Select Case tys
   Case 0
     If dectys > 0 Then tys_txt = Nums4(tys) & "тыс€ч "
   Case 1
     tys_txt = Nums4(tys) & "тыс€ча "
   Case 2, 3, 4
     tys_txt = Nums4(tys) & "тыс€чи "
   Case 5 To 9
     tys_txt = Nums4(tys) & "тыс€ч "
 End Select
 If dectys = 0 And tys = 0 And sottys <> 0 Then sottys_txt = sottys_txt & " тыс€ч "
eee:
 sot_txt = Nums3(sot)
 'проверяем десятки
 Select Case dec
   Case 1
     ed_txt = Nums5(ed)
     GoTo rrr
   Case 2 To 9
     dec_txt = Nums2(dec)
 End Select
 
 ed_txt = Nums1(ed)
rrr:
 'формируем итоговую строку
 СУММАПРОПИСЬЮ = decmil_txt & mil_txt & sottys_txt & dectys_txt & tys_txt & sot_txt & dec_txt & ed_txt
End Function
 
'вспомогательная функция для выделения из числа разрядов
Private Function Class(M, I)
  Class = Int(Int(M - (10 ^ I) * Int(M / (10 ^ I))) / 10 ^ (I - 1))
End Function



