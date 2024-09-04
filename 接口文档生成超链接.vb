Sub AutoSub()
    Dim ws As Worksheet
    Dim numSheets As Long
    Dim isFirstSheet As Boolean
    
    ' 获取工作簿中的工作表数量
    numSheets = ThisWorkbook.Worksheets.Count
    
    ' 遍历所有工作表
    For Each ws In ThisWorkbook.Worksheets
        isFirstSheet = (ws.Index = 1)
        
        ' 调用不同的处理函数
        If isFirstSheet Then
            Call MainSheetManager(ws)
        Else
            Call SubSheetManager(ws)
        End If
    Next ws
End Sub


Function MainSheetManager(rv_ws As Worksheet)
    Dim c_dir_start_cell As Range
    Dim c_dir_index_start_cell As Range
    Dim at_title As String
    Dim at_title_num As Long
    Dim numSheets As Long
    Dim i As Long

    Set c_dir_start_cell = rv_ws.Range("E13")
    Set c_dir_index_start_cell = rv_ws.Range("D13")

    at_title_num = 0

    ' 获取工作簿中的工作表数量
    numSheets = ThisWorkbook.Worksheets.Count

    ' 从第二个工作表开始遍历
    For i = 2 To numSheets
        Set ws = ThisWorkbook.Worksheets(i)
        at_title_num = at_title_num + 1
        
        ' 设置链接单元格
        Set linkCell = c_dir_start_cell.Offset(at_title_num)
        
        ' 创建超链接
        linkCell.Hyperlinks.Add Anchor:=linkCell, _
                                   Address:="", _
                                   SubAddress:="'" & ws.Name & "'!A1", _
                                   TextToDisplay:=ws.Name

        ' 设置编号
        c_dir_index_start_cell.Offset((at_title_num)).Value = at_title_num
    Next i

End Function

Function SubSheetManager(rv_ws As Worksheet)
    Dim isFirst As Boolean
    Dim isRunFirst As Boolean
    Dim isWork As Boolean    ' 判断正式的逻辑
    Dim isNoneLast As Boolean
    Dim isNoneCurr As Boolean

    Dim cell As Range
    Dim rowNumber As Long
    Dim c_dir_start_cell As Range
    Dim c_dir_index_start_cell As Range
    Dim c_data_col_offset As Long

    Dim at_count As Long
    Dim at_title As String
    Dim at_title_num As Long
    Dim at_data_start As Range
    Dim at_data_end As Range
    Dim at_data_area As Range
    
    Set c_dir_start_cell = rv_ws.Range("B2")
    Set c_dir_index_start_cell = rv_ws.Range("A2")
    c_data_col_offset = 8

    isFirst = True
    isRunFirst = True
    at_count = 0
    isWork = False
    isNoneLast = True
    at_title_num = 0
    rowNumber = 1

    ' 先创建一个返回封面的超链接
    call CreateLinkBackCover(sheet1, rv_ws.Range("D13"))

    ' 遍历 O 列,找到起始地址
    Do While isRunFirst
        Set cell = rv_ws.Cells(rowNumber, "O")
        'Set cell = rv_ws.Range("O" & rowNumber)
        
        If isFirst Then
            ' 检查首个有效内容，超时50个单元格
            If IsEmpty(cell.Value) Then
                rowNumber = rowNumber + 1
                at_count = at_count + 1
                If 50 < at_count Then
                    ' MsgBox "无有效数据，即将结束"
                    Exit Function
                End If
            Else
                isFirst = False
                isRunFirst = False
                isWork = True
            End If
        End If
    Loop

    ' 跑到这里就开始正式的处理逻辑
    Do While isWork
        Set cell = rv_ws.Cells(rowNumber, "O")
        isNoneCurr = IsEmpty(cell.Value)
        If isNoneLast And Not isNoneCurr Then '上无现有，数据块头部
            Set at_data_start = cell

            at_title = cell.Value '保存下标题
            at_title_num = at_title_num + 1

            ' 创建返回目录的标签
            Call CreateLinkBackDir(rv_ws, cell.Offset(, c_data_col_offset))
        ElseIf Not isNoneLast And isNoneCurr Then '上有现无，数据块末尾
            Set at_data_end = cell.Offset(-1, 0)
            
            ' 创建标题序号
            c_dir_index_start_cell.Offset(at_title_num).Value = at_title_num

            ' 获取数据区域
            Set at_data_area = GetExtendedRange(at_data_start, at_data_end, c_data_col_offset)

            ' 创建标题超链接,指向数据区域
            Call CreateLinkGoData(rv_ws, c_dir_start_cell.Offset(at_title_num), at_data_area, at_title)

        ElseIf isNoneLast And isNoneCurr Then '上无现无，结束
            isWork = False
            ' MsgBox "执行完成"
        End If
        rowNumber = rowNumber + 1
        isNoneLast = IsEmpty(cell.Value)
    Loop
 End Function

' 获取基址偏移的一片区域
Function GetAdjacentRegion(startCell As Range, rowsOffset As Long, colsOffset As Long) As Range
    ' 计算新的右下角单元格的位置
    Dim endCell As Range
    Set endCell = startCell.Offset(rowsOffset, colsOffset)
    
    ' 返回从起始单元格到计算出的右下角单元格的区域
    Set GetAdjacentRegion = startCell.Worksheet.Range(startCell.Address & ":" & endCell.Address)
End Function


Function CreateLinkGoData(rv_ws As Worksheet, rv_anchor As Range, rv_subadress As Range, rv_title As String)
    ' 创建一个指定文本的超链接到指定数据区域
    rv_ws.Hyperlinks.Add Anchor:=rv_anchor, _
                           Address:="", _
                           SubAddress:=rv_subadress.Address, _
                           ScreenTip:="", _
                           TextToDisplay:=rv_title
End Function

Function CreateLinkBackDir(rv_ws As Worksheet, rv_range As Range)
    ' 创建一个返回目录的超链接到指定单元格
    rv_ws.Hyperlinks.Add Anchor:=rv_range, _
                           Address:="", _
                           SubAddress:="'" & rv_ws.Name & "'!C9", _
                           ScreenTip:="", _
                           TextToDisplay:="返回目录"
End Function

Function CreateLinkBackCover(rv_ws As Worksheet, rv_range As Range)
    ' 创建一个返回封面的超链接到指定单元格
    rv_ws.Hyperlinks.Add Anchor:=rv_range, _
                           Address:="", _
                           SubAddress:="'" & rv_ws.Name & "'!D13", _
                           ScreenTip:="", _
                           TextToDisplay:="返回封面"
End Function

Function CreateLinkMainDir(rv_ws As Worksheet, rv_range As Range)
    ' 创建一个表名的超链接到指定单元格
    rv_ws.Hyperlinks.Add Anchor:=rv_range, _
                           Address:="", _
                           SubAddress:="'" & rv_ws.Name & "'!A1", _
                           ScreenTip:="", _
                           TextToDisplay:=rv_ws.Name
End Function

' 获得基于起始地址与结束地址加上列偏移的一片区域
Function GetExtendedRange(startCell As Range, endRowCell As Range, colOffset As Long) As Range
    ' 确定结束列的位置
    Dim endColCell As Range
    Dim ws As Worksheet
    
    ' 获取工作表对象
    Set ws = startCell.Worksheet
    
    ' 计算结束列的名称
    Dim endColName As String
    endColName = ws.Cells(1, startCell.Column + colOffset).Address(True, False, xlA1)
    
    ' 构建结束单元格的地址
    Dim endCellAddress As String
    ' 使用 endRowCell 的 Row 属性来获取行号
    endCellAddress = Left(endColName, InStr(endColName, "$") - 1) & endRowCell.Row
    
    ' 创建并返回新的区域
    Set GetExtendedRange = ws.Range(startCell.Address & ":" & endCellAddress)
End Function





