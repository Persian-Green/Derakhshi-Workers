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
    
    ' پيدا کردن اولين رديف خالي
        LastSourceRow = Range("P_LastSourceRow").Value
        LastDestRow = Range("P_LastDestRow").Value
    
    ' حلقه براي کپي کردن هر رديف از محدوده
        For i = 5 To LastSourceRow
            Dim EmptyRow As Boolean
            Dim j As Long
            
            ' فرض بر اين که رديف خالي است
            EmptyRow = True
            
            ' بررسي سلول‌هاي ستون A تا R براي هر رديف
            For j = 1 To 18 ' ستون A تا ستون R (1 تا 18)
                If wsSource.Cells(i, j).Value <> "" Then
                    EmptyRow = False
                    Exit For
                End If
            Next j
            
            ' اگر رديف خالي نبود، آن را کپي کن
            If Not EmptyRow  And LastSourceRow > 4 Then
                ' به روزرساني آخرين رديف مقصد
                LastDestRow = Range("P_LastDestRow").Value
                ' کپي کردن رديف از شيت مبدا به شيت مقصد
                wsDest.Range("A" & LastDestRow & ":T" & LastDestRow).Value = wsSource.Range("A" & i & ":T" & i).Value
            End If
        Next i

        MsgBox "اطلاعات با موفقيت ثبت شد"
End Sub
