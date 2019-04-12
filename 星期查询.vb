Private Sub weekdate()
	'查询距离今天指定天数的日期是星期几,比如今天是星期五，那么往后数第3天应该是星期一。
	''''''''''''
	'算法介绍，
	'比如向后查找20天，20除以7的余数是6，
	'如果几天是星期5，那么本周还剩2天，本周已经历4天
	'因为6大于2，所以用6减去7等于-1，
	'所以查询结果为5-1=4，即为周四
	'如果今天是周1，则查询结果为1+6=7，即为周日
	''''''''''''''''
	
    Dim x As Integer    '今天星期几
    Dim n As Integer    '向后查询几天
    Dim m As Integer    'n除以7的余数
    Dim r As Integer    '本星期中，今天后面还剩几天
    Dim l As Integer    '本星期中，今天之前有几天
    Dim z As Integer    '查询结果
    Dim weekdate        '
    
    weekdate = Array("星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日")
    
    'x = Range("D8")    '可以改为其他单元格
    x = InputBox("今天是星期几？请输入1-7之间的整数", , 1)
    
    'n = Range("D9")    '可以改为其他单元格
    n = InputBox("向后查询几天？请输入大于等于0的整数", , 0)
    
    m = n Mod 7
    r = 7 - x
    l = x - 1
    
    If m <= r Then
        z = x + m
    End If
    
    If m > r Then
        z = x - (7 - m)
    End If
    
    'Range("D10") = weekdate(z - 1) '可以改为其他单元格
    
    MsgBox (weekdate(z - 1))
    
End Sub