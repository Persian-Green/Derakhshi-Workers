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

    ' تنظيم فرمت ستون‌هاي جدول
        Range("A5:A999").NumberFormat = "[$-fa-IR,16]yyyy/mm/dd;@"  
        Range("C5:E999, H5:H999, J5:K999, P5:Y999").NumberFormat = "[h]:mm"
        Range("F5:G999, I5:I999, L5:O999").NumberFormat = "#,##0 ""تومان"""
        Range("F5:G999, I5:I999, L5:O999").Select
            With Selection
                .HorizontalAlignment = xlLeft
                .ShrinkToFit = True
                .ReadingOrder = xlRTL
            End With
        Range("A5").Select

    ' تنظيم عرض ستون‌ها
        Columns("B:D").ColumnWidth = 14
        Columns("F:H").ColumnWidth = 14
End Sub
