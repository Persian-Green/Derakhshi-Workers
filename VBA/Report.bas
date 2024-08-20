Sub Report()
    ' ریفرش کوئری اول
        With ActiveWorkbook.Connections("Query - P_ReportDate").OLEDBConnection
            .BackgroundQuery = False
            .Refresh
        End With
    
    ' ریفرش کوئری دوم بعد از اتمام کوئری اول
        With ActiveWorkbook.Connections("Query - FinalRows").OLEDBConnection
            .BackgroundQuery = False
            .Refresh
        End With
    
    ' ریفرش کوئری سوم بعد از اتمام کوئری دوم
        With ActiveWorkbook.Connections("Query - Report").OLEDBConnection
            .BackgroundQuery = False
            .Refresh
        End With

    ' تنظیم عرض سه ستون اول
        Columns("C:C").Select
        Selection.ColumnWidth = 14

    ' انتخاب سلول A5
        Range("A5").Select
End Sub
