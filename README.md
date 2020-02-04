# AI
AI数据处理代码
Sub 动态图表()
Dim chart1 As Chart
Set chart1 = ActiveSheet.ChartObjects.Add(Range("I1").Left, Range("I1").Top, 400, 300).Chart
chart1.SetSourceData (ActiveSheet.Range("A1:H2"))
chart1.ChartType = xlLine
ActiveSheet.ChartObjects(1).Name = "动态图表"
chart1.SeriesCollection(1).Name = "第一系列"
Set chart1 = Nothing
End Sub
