# VBA
EXCEL VBA
Sub 组内平均值()
'
' 组内平均值 Macro
' 宏由 zzh_my@163.com 录制，时间: 2015/11/18
'
'点名    初始化时间(s)   测量时刻    次序号  平面x(m)    平面y(m)    大地高(m)   解类型  备注    观测类型    观测者  x组内离均值 y组内离均值 H组内离均值 abs(x组内离均值)    abs(y组内离均值)    abs(H组内离均值)        x   x百分比 y   y百分比 H   H百分比
      '需按照点名先排列好

            n = 0
            x = 0
            y = 0
            H = 0
            sheetname = "xzl-gnss"
           ' rowNumber = 933  'x = range("A65536").end(xlup).row  i = cells(rows.count,1).end(xlup).row+1
            rowNumber = Worksheets(sheetname).Cells(Rows.Count, 1).End(xlUp).Row
      For i = 2 To rowNumber Step 1
            
           ' MsgBox (Name)
           '假如点名和下一条相等，且观测类型相等
           If (Worksheets(sheetname).Cells(i, 1).Value = Worksheets(sheetname).Cells(i + 1, 1)) And (Worksheets(sheetname).Cells(i, 10).Value = Worksheets(sheetname).Cells(i + 1, 10)) Then
               n = n + 1
               
                x = x + Worksheets(sheetname).Cells(i, 5).Value
                y = y + Worksheets(sheetname).Cells(i, 6).Value
                H = H + Worksheets(sheetname).Cells(i, 7).Value
           Else
                For j = 1 To n + 1 Step 1
                    x_mean = x / n
                    y_mean = y / n
                    H_mean = H / n
                    h1 = Worksheets(sheetname).Cells(i - j + 1, 7).Value
                   ' MsgBox (x_mean)
                    Worksheets(sheetname).Cells(i - j + 1, 12).Value = Worksheets(sheetname).Cells(i - j + 1, 5).Value - x_mean
                    Worksheets(sheetname).Cells(i - j + 1, 13).Value = Worksheets(sheetname).Cells(i - j + 1, 6).Value - y_mean
                    Worksheets(sheetname).Cells(i - j + 1, 14).Value = Worksheets(sheetname).Cells(i - j + 1, 7).Value - H_mean
                 Next j
                    x = 0
                    y = 0
                    H = 0
                    x_mean = 0
                    y_mean = 0
                    H_mean = 0
                    n = 0
                    
           End If

     
     Next i
    
End Sub

Sub 统计单模式频率区间()
'
' 组内平均值 Macro
' 宏由 zzh_my@163.com 录制，时间: 2015/11/18
'
'点名    初始化时间(s)   测量时刻    次序号  平面x(m)    平面y(m)    大地高(m)   解类型  备注    观测类型    观测者  x组内离均值 y组内离均值 H组内离均值 abs(x组内离均值)    abs(y组内离均值)    abs(H组内离均值)        x   x百分比 y   y百分比 H   H百分比
                                                                                                                                                    
'
           
            x1 = 0
            x2 = 0
            x3 = 0
            x4 = 0
            
            y1 = 0
            y2 = 0
            y3 = 0
            y4 = 0
            
            h1 = 0
            h2 = 0
            h3 = 0
            h4 = 0
            

            '区间范围m key0~key1  key1~key2 key2~key3 key3以上
            key0 = 0
            key1 = 0.03
            key2 = 0.05
            key3 = 0.1

            sheetname = "xzl-gnss"
           ' rowNumber = 933  'x = range("A65536").end(xlup).row  i = cells(rows.count,1).end(xlup).row+1
            rowNumber = Worksheets(sheetname).Cells(Rows.Count, 1).End(xlUp).Row
            
            'x循环
      For i = 2 To rowNumber Step 1
            
           If (Worksheets(sheetname).Cells(i, 15).Value) <= key1 Then
               x1 = x1 + 1
           End If
           If key1 < (Worksheets(sheetname).Cells(i, 15).Value) And (Worksheets(sheetname).Cells(i, 15).Value) <= key2 Then
               x2 = x2 + 1
           End If

           If key2 < (Worksheets(sheetname).Cells(i, 15).Value) And (Worksheets(sheetname).Cells(i, 15).Value) <= key3 Then
               x3 = x3 + 1
           End If
           
           If key3 < (Worksheets(sheetname).Cells(i, 15).Value) Then
               x4 = x4 + 1
           End If
     
     Next i
     '个数
     Worksheets(sheetname).Cells(2, 19).Value = x1
     Worksheets(sheetname).Cells(3, 19).Value = x2
     Worksheets(sheetname).Cells(4, 19).Value = x3
     Worksheets(sheetname).Cells(5, 19).Value = x4
     '百分比
     Worksheets(sheetname).Cells(2, 20).Value = x1 / (x1 + x2 + x3 + x4)
     Worksheets(sheetname).Cells(3, 20).Value = x2 / (x1 + x2 + x3 + x4)
     Worksheets(sheetname).Cells(4, 20).Value = x3 / (x1 + x2 + x3 + x4)
     Worksheets(sheetname).Cells(5, 20).Value = x4 / (x1 + x2 + x3 + x4)
     
                 'y循环
      For i = 2 To rowNumber Step 1
            
           If (Worksheets(sheetname).Cells(i, 16).Value) <= key1 Then
               y1 = y1 + 1
           End If
           If key1 < (Worksheets(sheetname).Cells(i, 16).Value) And (Worksheets(sheetname).Cells(i, 16).Value) <= key2 Then
               y2 = y2 + 1
           End If

           If key2 < (Worksheets(sheetname).Cells(i, 16).Value) And (Worksheets(sheetname).Cells(i, 16).Value) <= key3 Then
               y3 = y3 + 1
           End If
           
           If key3 < (Worksheets(sheetname).Cells(i, 16).Value) Then
               y4 = y4 + 1
           End If
     
     Next i
     '个数
     Worksheets(sheetname).Cells(2, 21).Value = y1
     Worksheets(sheetname).Cells(3, 21).Value = y2
     Worksheets(sheetname).Cells(4, 21).Value = y3
     Worksheets(sheetname).Cells(5, 21).Value = y4
     '百分比
     Worksheets(sheetname).Cells(2, 22).Value = y1 / (y1 + y2 + y3 + y4)
     Worksheets(sheetname).Cells(3, 22).Value = y2 / (y1 + y2 + y3 + y4)
     Worksheets(sheetname).Cells(4, 22).Value = y3 / (y1 + y2 + y3 + y4)
     Worksheets(sheetname).Cells(5, 22).Value = y4 / (y1 + y2 + y3 + y4)
     
                      'H循环
      For i = 2 To rowNumber Step 1
            
           If (Worksheets(sheetname).Cells(i, 17).Value) <= key1 Then
               h1 = h1 + 1
           End If
           If key1 < (Worksheets(sheetname).Cells(i, 17).Value) And (Worksheets(sheetname).Cells(i, 17).Value) <= key2 Then
               h2 = h2 + 1
           End If

           If key2 < (Worksheets(sheetname).Cells(i, 17).Value) And (Worksheets(sheetname).Cells(i, 17).Value) <= key3 Then
               h3 = h3 + 1
           End If
           
           If key3 < (Worksheets(sheetname).Cells(i, 17).Value) Then
               h4 = h4 + 1
           End If
     
     Next i
     '个数
     Worksheets(sheetname).Cells(2, 23).Value = h1
     Worksheets(sheetname).Cells(3, 23).Value = h2
     Worksheets(sheetname).Cells(4, 23).Value = h3
     Worksheets(sheetname).Cells(5, 23).Value = h4
     '百分比
     Worksheets(sheetname).Cells(2, 24).Value = h1 / (h1 + h2 + h3 + h4)
     Worksheets(sheetname).Cells(3, 24).Value = h2 / (h1 + h2 + h3 + h4)
     Worksheets(sheetname).Cells(4, 24).Value = h3 / (h1 + h2 + h3 + h4)
     Worksheets(sheetname).Cells(5, 24).Value = h4 / (h1 + h2 + h3 + h4)
End Sub

