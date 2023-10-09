Attribute VB_Name = "ģ��1"
Sub CreateScatterChartWithLine()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chartRange As Range
    
    Dim series As series
    Dim i As Integer
    Dim differenceThreshold As Double ' ��ֵ��ֵ
    Dim lastPeakIndex As Integer ' ��һ����ֵ������
    Dim xColumn As Range ' X �з�Χ
    Dim dColumn As Range ' D �з�Χ

    
    ' ���ù�����
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' ����ͼ��Χ
    Set chartRange = ws.Range("A1:B" & ws.Cells(Rows.Count, "A").End(xlUp).Row)
    
    ' ����ɢ��ͼ
    Set chartObj = ws.ChartObjects.Add(Left:=200, Width:=1000, Top:=20, Height:=500)
    chartObj.Chart.ChartType = xlXYScatterLinesNoMarkers ' ����ͼ������Ϊɢ��ͼ��ֱ��
    
    ' ����ͼ������Դ
    chartObj.Chart.SetSourceData Source:=chartRange
    
    ' �����Զ���ͼ����������ԣ�����⡢�������
    
    
    ' ���ò�ֵ��ֵ��������Ҫ�Զ�����ֵ��С
    differenceThreshold = 200 ' ��ֵ��ֵ
    lastPeakIndex = -1 ' ��ʼ����һ����ֵ����Ϊ-1
    Row = 1 ' ��ʼ�� row ����Ϊ 1
    
    ' ��ȡ���������
    Set ws = ActiveSheet
    
    ' ָ�� X �к� D �еķ�Χ
    Set xColumn = ws.Range("A2:A" & ws.Cells(Rows.Count, "A").End(xlUp).Row)
    Set dColumn = ws.Range("D1")
    
    ' ����ÿ������ϵ��
    For Each series In chartObj.Chart.SeriesCollection
        ' �������ݵ�
        For i = 2 To series.Points.Count - 1 ' �ӵڶ������ݵ㿪ʼ������Խ��
            ' �ж�X���ֵ�Ƿ����250
            If series.XValues(i) > 250 Then
                ' �ж��Ƿ������ֵ����
                If series.Values(i) >= series.Values(i - 1) And series.Values(i) >= series.Values(i + 1) Then
                    ' ���㵱ǰ���ݵ���ǰ�����ݵ�Ĳ�ֵ
                    Dim diff1 As Double
                    Dim diff2 As Double
                    diff1 = Abs(series.Values(i) - series.Values(i - 1))
                    diff2 = Abs(series.Values(i) - series.Values(i + 1))
                    
                    ' �����ֵ������ֵ��������Ϊ�Ƿ�ֵ
                    If diff1 > differenceThreshold Or diff2 > differenceThreshold Then
                        ' ����Ƿ�����һ����ֵ��ͬ
                        If i <> lastPeakIndex Then
                            ' ��ȡ X ���ֵ
                            Dim xValue As Variant
                            xValue = series.XValues(i)
                            
                            ' ���ַ�ֵ�㣬������ݱ�ǩ
                            series.Points(i).HasDataLabel = True
                            series.Points(i).DataLabel.Text = xValue  ' ������Ҫ�Զ����ǩ����
                            
                            ' ������һ����ֵ������
                            lastPeakIndex = i
                            
                            ' �� D �еĵ�ǰ��д�� X ���ֵ
                            dColumn.Offset(Row, 0).Value = xValue
                            
                            ' ���� row ����������д��
                            Row = Row + 1
                        End If
                    End If
                End If
            End If
        Next i
    Next series
End Sub

