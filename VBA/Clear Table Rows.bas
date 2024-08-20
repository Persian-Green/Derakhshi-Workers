Sub ClearTableRows()
    ' متغیرها
        Dim ws As Worksheet
        Dim tbl As ListObject
    
    ' تنظیم شیت و جدول
        Set ws = ThisWorkbook.Sheets("Enter") 
        Set tbl = ws.ListObjects("T_Form")
    
    ' بررسی تعداد ردیف‌ها و پاک کردن آنها
        If tbl.ListRows.Count > 0 Then
            tbl.DataBodyRange.Rows.Delete
        End If
End Sub
