Attribute VB_Name = "模块1"
Sub CreateScatterChartWithLine()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chartRange As Range
    
    Dim series As series
    Dim i As Integer
    Dim differenceThreshold As Double ' 差值阈值
    Dim lastPeakIndex As Integer ' 上一个峰值的索引
    Dim xColumn As Range ' X 列范围
    Dim dColumn As Range ' D 列范围

    
    ' 设置工作表
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' 设置图表范围
    Set chartRange = ws.Range("A1:B" & ws.Cells(Rows.Count, "A").End(xlUp).Row)
    
    ' 创建散点图
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=1000, Top:=20, Height:=500)
    chartObj.Chart.ChartType = xlXYScatterLinesNoMarkers ' 设置图表类型为散点图带直线
    
    ' 设置图表数据源
    chartObj.Chart.SetSourceData Source:=chartRange
    
    ' 可以自定义图表的其他属性，如标题、坐标轴等
    
    
    ' 设置差值阈值，根据需要自定义阈值大小
    differenceThreshold = 200 ' 差值阈值
    lastPeakIndex = -1 ' 初始化上一个峰值索引为-1
    Row = 1 ' 初始化 row 变量为 1
    
    ' 获取工作表对象
    Set ws = ActiveSheet
    
    ' 指定 X 列和 D 列的范围
    Set xColumn = ws.Range("A2:A" & ws.Cells(Rows.Count, "A").End(xlUp).Row)
    Set dColumn = ws.Range("D1")
    
    ' 遍历每个数据系列
    For Each series In chartObj.Chart.SeriesCollection
        ' 遍历数据点
        For i = 2 To series.Points.Count - 1 ' 从第二个数据点开始，避免越界
            ' 判断X轴的值是否大于250
            If series.XValues(i) > 250 Then
                ' 判断是否满足峰值条件
                If series.Values(i) >= series.Values(i - 1) And series.Values(i) >= series.Values(i + 1) Then
                    ' 计算当前数据点与前后数据点的差值
                    Dim diff1 As Double
                    Dim diff2 As Double
                    diff1 = Abs(series.Values(i) - series.Values(i - 1))
                    diff2 = Abs(series.Values(i) - series.Values(i + 1))
                    
                    ' 如果差值满足阈值条件，认为是峰值
                    If diff1 > differenceThreshold Or diff2 > differenceThreshold Then
                        ' 检查是否与上一个峰值相同
                        If i <> lastPeakIndex Then
                            ' 获取 X 轴的值
                            Dim xValue As Variant
                            xValue = series.XValues(i)
                            
                            ' 发现峰值点，添加数据标签
                            series.Points(i).HasDataLabel = True
                            series.Points(i).DataLabel.Text = xValue  ' 根据需要自定义标签内容
                            
                            ' 更新上一个峰值的索引
                            lastPeakIndex = i
                            
                            ' 在 D 列的当前行写入 X 轴的值
                            dColumn.Offset(Row, 0).Value = xValue
                            
                            ' 增加 row 变量以逐行写入
                            Row = Row + 1
                        End If
                    End If
                End If
            End If
        Next i
    Next series
End Sub

