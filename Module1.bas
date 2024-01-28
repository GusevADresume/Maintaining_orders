Attribute VB_Name = "Module1"
Option Explicit
Global self_employed, master_name, master_phone, master_id, fl, d, claim, client_name, client_addres, client_phone, master_cash, order_price, path, iPath, build_date, furn_list, cuirsive_summ As String
Global fcell As Variant
Dim x, y As Integer

'MAIN TABLE!!!!!

Sub define_string()
Application.ScreenUpdating = False
Application.EnableEvents = False
d = Range("Q2").End(xlDown).Row
For y = 17 To 17
        For x = 2 To d
            If Cells(x, y).Value = "новая" Then
                claim = Cells(x, y - 15).Value
                client_name = Cells(x, y - 12).Value
                client_addres = Cells(x, y - 10).Value
                client_phone = Cells(x, y - 11).Value
                master_cash = Cells(x, y - 5).Value
                build_date = Cells(x, y - 14).Value
                order_price = Cells(x, y - 7).Value
                master_name = Cells(x, y - 8).Value
                Sheets("Тех3").Select
                Set fcell = Columns("B:B").Find(master_name)
                If Not fcell Is Nothing Then
                    master_id = Cells(fcell.Row, 1).Value
                    master_phone = Cells(fcell.Row, 3).Value
                    master_name = Cells(fcell.Row, 2).Value
                        Sheets("Заявки").Select
                            Cells(x, y - 8).Interior.Color = RGB(255, 0, 255)
                            Cells(x, y + 4).Value = 1
                        Sheets("Тех3").Select
                    create_registry
                End If
                If master_cash < order_price Then
                    order_price = ""
                End If
                Sheets("Заявки").Select
                furn_list = Cells(x, y + 3).Value
                cuirsive_summ = coursive((Cells(x, y - 7).Value)) + "рублей. 00 копеек"
                create_path
                note_to_the_master
                uploadword
                Cells(x, y).Value = "передана мастеру"
            End If
        Next
    Next
    download_payments_registry
End Sub

Sub download_payments_registry()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim r As Variant
    Dim d, dt As String
    dt = Date
    d = Range("G1").End(xlDown).Row
    Sheets("Тех4").Select
    r = Range("A1:G1000")
    Workbooks.Add
    Range("A1:G1000") = r
    ActiveWorkbook.SaveAs (ThisWorkbook.path & "\payment_registers\" & dt & ".xlsx")
    ActiveWorkbook.Close
    Range("A2:G1000").Value = ""
    Sheets("Заявки").Select
    
End Sub

Sub create_registry()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim d As String
    Sheets("Тех4").Select
    d = Range("A1").End(xlDown).Row
    If d = 1048576 Then
        d = 1
    End If
    Cells(d + 1, 1).Value = master_id
    Cells(d + 1, 2).Value = master_name
    Cells(d + 1, 3).Value = master_phone
    Cells(d + 1, 4).Value = "Новое"
    Cells(d + 1, 5).Value = master_cash
    Cells(d + 1, 6).Value = "Новое"
    Cells(d + 1, 7).Value = Date
End Sub

Sub create_path()
Application.EnableEvents = False
iPath = ThisWorkbook.path
Set path = CreateObject("Scripting.FileSystemObject")
If Not path.FolderExists(iPath & "\" & "\Материалы\" & claim) Then
        path.CreateFolder (iPath & "\" & "\Материалы\" & claim)
    End If

End Sub

Sub note_to_the_master()
Application.ScreenUpdating = False
Application.EnableEvents = False
iPath = ThisWorkbook.path
fl = ActiveWorkbook.path & "\Материалы\" & "\" & claim & "\" & claim & ".txt"
    Open fl For Output As 1
         Print #1, "Клиент: " & client_name
         Print #1, "Мебель: " & furn_list
         Print #1, "Адрес: " & client_addres
         Print #1, "Телефон: " & client_phone
         Print #1, "ПРОШУ ВАС СОЗВОНИТЬСЯ С КЛИЕНТОМ И НАЗНАЧИТЬ ВРЕМЯ ВИЗИТА, ТАК ЖЕ ПРОШУ ПО ВОЗМОЖНОСТИ ОСУЩЕСТВИТЬ ЗВОНОК КЛИЕНТУ ЗА ЧАС ДО ВИЗИТА, при наличии мусорки рядом с домом, просим вас выбросить коробки"
         Print #1, "Оплата: " & master_cash
    Close 1

End Sub

Sub uploadword()
Dim wda As Object
Application.ScreenUpdating = False
Application.EnableEvents = False

Set wda = CreateObject("Word.Application")
With wda
.Visible = True
.Documents.Open ActiveWorkbook.path & "\template\template.docx"
End With
Application.ScreenUpdating = False
wda.ActiveDocument.Bookmarks("claim").Select
wda.Selection = claim
wda.ActiveDocument.Bookmarks("client_name").Select
wda.Selection = client_name
wda.ActiveDocument.Bookmarks("client_addres").Select
wda.Selection = client_addres
wda.ActiveDocument.Bookmarks("client_phone").Select
wda.Selection = client_phone
wda.ActiveDocument.Bookmarks("order_price").Select
wda.Selection = order_price
wda.ActiveDocument.Bookmarks("order_price1").Select
wda.Selection = order_price
wda.ActiveDocument.Bookmarks("order_price2").Select
wda.Selection = order_price
wda.ActiveDocument.Bookmarks("cuirsive_summ").Select
wda.Selection = cuirsive_summ
wda.ActiveDocument.Bookmarks("build_date").Select
wda.Selection = build_date

'save
wda.ActiveDocument.SaveAs ActiveWorkbook.path & "\Материалы\" & "\" & claim & "\" & claim & ".docx"
'savepdf
'wda.ActiveDocument.ExportAsFixedFormat OutputFileName:=ActiveWorkbook.path & "\Њатериалы\" & "\" & claim & "\" & claim & ".pdf", ExportFormat:=17
'wda.ActiveDocument.Close
GetObject(, "Word.Application").Quit


End Sub

Sub getData()
Application.ScreenUpdating = False
Application.EnableEvents = False
Dim x, y, cucle, act_summ As Integer
Dim summ As Variant
Dim d As String
Dim Folder, Filename, lastCell, number, assem_date, client, phone, addres, city, furn_type, old_number, product_group As String
Dim bkBook_1 As Excel.Workbook
Dim bkBook_2 As Excel.Workbook
Dim LArray() As String
Dim furn_list As String

Set bkBook_1 = ActiveWorkbook
lastCell = Range("A1").End(xlDown).Row


Folder = ActiveWorkbook.path & "\orig_file\"
Filename = Dir(Folder & "/", vbNormal)
Workbooks.Open Filename:=ActiveWorkbook.path & "\orig_file\" & Filename
d = Range("E10").End(xlDown).Row
Set bkBook_2 = ActiveWorkbook


For y = 5 To 5
        For x = 11 To d
        If d >= 10000 Then
            d = 11
        End If
            cucle = cucle + 1
            bkBook_2.Activate
            
            If Cells(x, y).Value = "" Then
                summ = summ + Cells(x, y + 8).Value
                furn_list = furn_list & " " & Cells(x, y + 16).Value & "=" & Cells(x, y + 17).Value
                
            End If
            
            If Cells(x, y).Value <> "" Then
            
            number = Cells(x, y).Value
            
            If old_number <> "" Then
                If old_number <> number Then
                    bkBook_1.Activate
                    If summ * 0.087 < 900 Then
                        summ = 986
                        If product_group <> "35" Or product_group <> "38" Or product_group <> "39" Or product_group <> "40" Then
                            summ = 1261.5
                        End If
                        Else
                        summ = summ * 0.087
                    End If
                    Cells(lastCell, 14).Value = summ
                    Cells(lastCell, 20).Value = furn_list
                    If Cells(lastCell, 14).Value < Cells(lastCell, 10).Value Then
                        Cells(lastCell, 10).Value = ""
                    End If
                    summ = 0
                    furn_list = ""
                    bkBook_2.Activate
                End If
            End If
                
            
            
            LArray = Split(Cells(x, y - 2).Value, "_")
            city = LArray(1)
            furn_type = LArray(2)
            assem_date = Cells(x, y - 3).Value
            client = Cells(x, y + 1).Value
            phone = Cells(x, y + 3).Value
            addres = Cells(x, y + 2).Value
            act_summ = Cells(x, y + 7).Value
            product_group = Cells(x, y + 12).Value
            
            
            summ = summ + Cells(x, y + 8).Value
            furn_list = furn_list & " " & Cells(x, y + 16).Value & "=" & Cells(x, y + 17).Value
            
            lastCell = lastCell + 1
            bkBook_1.Activate
            

            Cells(lastCell, 1).Value = Cells(lastCell - 1, 1).Value + 1
            Cells(lastCell, 2).Value = number
            Cells(lastCell, 3).Value = Format(Date, "Short Date")
            Cells(lastCell, 4).Value = Format(assem_date, "Short Date")
            Cells(lastCell, 5).Value = client
            Cells(lastCell, 6).Value = phone
            Cells(lastCell, 7).Value = addres
            Cells(lastCell, 8).Value = city
            Cells(lastCell, 10).Value = act_summ
            Cells(lastCell, 16).Value = furn_type
            Cells(lastCell, 17).Value = "новая"
            
            old_number = number
            'MsgBox (old_number)
            
           End If
           If Cells(x, y).Row = d Or d = 10000 Then
                bkBook_1.Activate
                If summ * 0.087 < 900 Then
                    summ = 986
                    If product_group <> "35" Or product_group <> "38" Or product_group <> "39" Or product_group <> "40" Then
                        summ = 1261.5
                    End If
                    Else
                        summ = summ * 0.087
                End If
                Cells(lastCell, 14).Value = summ
                Cells(lastCell, 20).Value = furn_list
                bkBook_2.Activate
            
            End If
           
           
        Next
    Next

ActiveWorkbook.Close

End Sub


Sub CloseOrder()
Dim x, y As Integer
Dim lastCell As String
Application.ScreenUpdating = False

lastCell = Range("Q1").End(xlDown).Row
For y = 17 To 17
    For x = 1 To lastCell
    If Cells(x, y).Value = "закрыта" And Cells(x, y + 1).Value = "" Then
        Range(Cells(x, y - 16), Cells(x, y - 4)).Interior.Color = RGB(0, 176, 240)
        Range(Cells(x, y - 2), Cells(x, y)).Interior.Color = RGB(0, 176, 240)
        Cells(x, y + 1).Value = Format(Now, "Short Date")
    End If
    If Cells(x, y).Value = "Отменен_клиентом" And Cells(x, y + 1).Value = "" Then
        Range(Cells(x, y - 16), Cells(x, y - 4)).Interior.Color = RGB(0, 176, 240)
        Range(Cells(x, y - 2), Cells(x, y)).Interior.Color = RGB(0, 176, 240)
        Cells(x, y + 1).Value = Format(Now, "Short Date")
    End If
    If Cells(x, y).Value = "Отменен_магазином" And Cells(x, y + 1).Value = "" Then
        Range(Cells(x, y - 16), Cells(x, y - 4)).Interior.Color = RGB(66, 245, 105)
        Range(Cells(x, y - 2), Cells(x, y)).Interior.Color = RGB(66, 245, 105)
        Cells(x, y + 1).Value = Format(Now, "Short Date")
    End If
    If Cells(x, y).Value = "рекламация" And Cells(x, y + 1).Value = "" Then
        Range(Cells(x, y - 16), Cells(x, y - 4)).Interior.Color = RGB(245, 150, 67)
        Range(Cells(x, y - 2), Cells(x, y)).Interior.Color = RGB(245, 150, 66)
        Cells(x, y + 1).Value = Format(Now, "Short Date")
    End If
    If Cells(x, y).Value = "претензия" And Cells(x, y + 1).Value = "" Then
        Range(Cells(x, y - 16), Cells(x, y - 4)).Interior.Color = RGB(245, 150, 67)
        Range(Cells(x, y - 2), Cells(x, y)).Interior.Color = RGB(245, 150, 66)
        Cells(x, y + 1).Value = Format(Now, "Short Date")
    End If
    If Cells(x, y).Value = "Отменен_КМ" And Cells(x, y + 1).Value = "" Then
        Range(Cells(x, y - 16), Cells(x, y - 4)).Interior.Color = RGB(245, 150, 67)
        Range(Cells(x, y - 2), Cells(x, y)).Interior.Color = RGB(245, 159, 68)
        Cells(x, y + 1).Value = Format(Now, "Short Date")
    End If
    If Cells(x, y).Value = "Отменен_нами" And Cells(x, y + 1).Value = "" Then
        Range(Cells(x, y - 16), Cells(x, y - 4)).Interior.Color = RGB(245, 150, 67)
        Range(Cells(x, y - 2), Cells(x, y)).Interior.Color = RGB(235, 159, 68)
        Cells(x, y + 1).Value = Format(Now, "Short Date")
    End If
        
    Next
Next


End Sub



Sub formatDate()
Dim x, y As Integer
Dim lastCell As String
Dim val, typ As String
lastCell = Range("D1").End(xlDown).Row


For y = 4 To 4
    For x = 2 To lastCell
    If TypeName(Cells(x, y).Value) = "String" Then
        Cells(x, y).Value = DateValue(Format(Cells(x, y).Value, "Short Date"))
    End If
    
    Next
Next

End Sub

Sub formatStatusDate()
Dim x, y As Integer
Dim lastCell As String
Dim val, typ As String
lastCell = Range("D1").End(xlDown).Row


For y = 18 To 18
    For x = 2 To lastCell
    If TypeName(Cells(x, y).Value) = "String" Then
        Cells(x, y).Value = DateValue(Format(Cells(x, y).Value, "Short Date"))
    End If
    
    Next
Next

End Sub


