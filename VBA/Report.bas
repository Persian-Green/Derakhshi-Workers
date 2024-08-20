Sub Report()
    ' ريفرش کوئري اول
        With ActiveWorkbook.Connections("Query - P_ReportFromDate").OLEDBConnection
            .BackgroundQuery = False
            .Refresh
        End With
    
    ' ريفرش کوئري دوم بعد از اتمام کوئري اول
            With ActiveWorkbook.Connections("Query - P_ReportToDate").OLEDBConnection
            .BackgroundQuery = False
            .Refresh
        End With

    ' ريفرش کوئري سوم بعد از اتمام کوئري دوم
        With ActiveWorkbook.Connections("Query - FinalRows").OLEDBConnection
            .BackgroundQuery = False
            .Refresh
        End With

    ' ريفرش کوئري چهارم بعد از اتمام کوئري سوم
        With ActiveWorkbook.Connections("Query - Report").OLEDBConnection
            .BackgroundQuery = False
            .Refresh
        End With

    ' تنظيم فرمت ستون‌ايس جدول
        Range("B5:D999,G5:G999,I5:R999").NumberFormat = "[h]:mm"
        Range("E5:F999,H5:H999").NumberFormat = "#,##0 ""تومان"""

    ' تنظيم عرض سه ستون اول
        Columns("B:D").ColumnWidth = 14
        Columns("F:H").ColumnWidth = 14
End Sub