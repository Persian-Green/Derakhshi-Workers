Sub CopyTableDataRowByRow()
    ' تعريف متغيرها
        Dim wsSource As Worksheet ' شيت مبدا
        Dim wsDest As Worksheet ' شيت مقصد
        Dim LastSourceRow As Long ' آخرين رديف مبدا
        Dim LastDestRow As Long ' آخرين رديف مقصد
        Dim i As Long ' براي حلقه

    ' شيت‌هاي مبدا و مقصد
        Set wsSource = ThisWorkbook.Sheets("Enter")
        Set wsDest = ThisWorkbook.Sheets("DB")

    ' ریفرش کردن سلولی که آخرین آی‌دی مقصد را نگه می‌دارد
        Range("P_LastID").Formula = "=ROUNDUP(MAX(T_DB[ID]), 0)"

    ' پيدا کردن اولين رديف خالي
        LastSourceRow = Range("P_LastSourceRow").Value
        LastDestRow = Range("P_LastDestRow").Value
    
    ' حلقه براي کپي کردن هر رديف از محدوده
        For i = 5 To LastSourceRow
            Dim EmptyRow As Boolean
            Dim j As Long
            
            ' فرض بر اين که رديف خالي است
                EmptyRow = True
            
            ' به استثنای ستون D بررسي سلول‌هاي ستون A تا T براي هر رديف
                For j = 1 To 3 ' ستون A تا ستون C (1 - 3)
                    If wsSource.Cells(i, j).Value <> "" Then
                        EmptyRow = False
                        Exit For
                    End If
                Next j

                For j = 5 To 20 ' ستون E تا ستون T (5 - 20)
                    If wsSource.Cells(i, j).Value <> "" Then
                        EmptyRow = False
                        Exit For
                    End If
                Next j
            
            ' اگر رديف خالي نبود و اصلا ردیفی ایجاد شده بود، آن را کپي کن
                If Not EmptyRow And LastSourceRow > 4 Then
                    ' به روزرساني آخرين رديف مقصد
                        LastDestRow = Range("P_LastDestRow").Value
                    ' کپي کردن رديف از شيت مبدا به شيت مقصد
                        wsDest.Range("A" & LastDestRow & ":Z" & LastDestRow).Value = wsSource.Range("A" & i & ":Z" & i).Value
                End If
        Next i
    ' پیغام
        MsgBox "اطلاعات با موفقيت ثبت شد"
End Sub