Attribute VB_Name = "Module2"
Function coursive(n As Double) As String
 
 Dim Nums1, Nums2, Nums3, Nums4 As Variant
 
 Nums1 = Array("", "���� ", "��� ", "��� ", "������ ", "���� ", "����� ", "���� ", "������ ", "������ ")
 Nums2 = Array("", "������ ", "�������� ", "�������� ", "����� ", "��������� ", "���������� ", "��������� ", _
                        "����������� ", "��������� ")
 Nums3 = Array("", "��� ", "������ ", "������ ", "��������� ", "������� ", "�������� ", "������� ", _
                        "��������� ", "��������� ")
 Nums4 = Array("", "���� ", "��� ", "��� ", "������ ", "���� ", "����� ", "���� ", "������ ", "������ ")
 Nums5 = Array("������ ", "����������� ", "���������� ", "���������� ", "������������ ", _
                        "���������� ", "����������� ", "���������� ", "������������ ", "������������ ")
 
 If n <= 0 Then
   ������������� = "����"
   Exit Function
 End If
 '��������� ����� �� �������, ��������� ��������������� ������� Class
 ed = Class(n, 1)
 dec = Class(n, 2)
 sot = Class(n, 3)
 tys = Class(n, 4)
 dectys = Class(n, 5)
 sottys = Class(n, 6)
 mil = Class(n, 7)
 decmil = Class(n, 8)
 
 '��������� ��������
 Select Case decmil
   Case 1
     mil_txt = Nums5(mil) & "��������� "
     GoTo www
   Case 2 To 9
     decmil_txt = Nums2(decmil)
 End Select
 Select Case mil
   Case 1
     mil_txt = Nums1(mil) & "������� "
   Case 2, 3, 4
     mil_txt = Nums1(mil) & "�������� "
   Case 5 To 20
     mil_txt = Nums1(mil) & "��������� "
 End Select
www:
 sottys_txt = Nums3(sottys)
 '��������� ������
 Select Case dectys
   Case 1
     tys_txt = Nums5(tys) & "����� "
     GoTo eee
   Case 2 To 9
     dectys_txt = Nums2(dectys)
 End Select
 Select Case tys
   Case 0
     If dectys > 0 Then tys_txt = Nums4(tys) & "����� "
   Case 1
     tys_txt = Nums4(tys) & "������ "
   Case 2, 3, 4
     tys_txt = Nums4(tys) & "������ "
   Case 5 To 9
     tys_txt = Nums4(tys) & "����� "
 End Select
 If dectys = 0 And tys = 0 And sottys <> 0 Then sottys_txt = sottys_txt & " ����� "
eee:
 sot_txt = Nums3(sot)
 '��������� �������
 Select Case dec
   Case 1
     ed_txt = Nums5(ed)
     GoTo rrr
   Case 2 To 9
     dec_txt = Nums2(dec)
 End Select
 
 ed_txt = Nums1(ed)
rrr:
 '��������� �������� ������
 ������������� = decmil_txt & mil_txt & sottys_txt & dectys_txt & tys_txt & sot_txt & dec_txt & ed_txt
End Function
 
'��������������� ������� ��� ��������� �� ����� ��������
Private Function Class(M, I)
  Class = Int(Int(M - (10 ^ I) * Int(M / (10 ^ I))) / 10 ^ (I - 1))
End Function



