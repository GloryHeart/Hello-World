Sub 填充cluster()
    CurrentPath = "?"
    fileobjecta = Dir(CurrentPath) '获取指定目录下第一个文件完整名称
    Application.Visible = False
    Do While fileobjecta <> "" And fileobjecta <> "Excel 自定义.exportedUI" '循环体条件
        Debug.Print (fileobjecta) '测试打印
        filepath = CurrentPath & fileobjecta '还原路径
        Workbooks.Open FileName:= _
            filepath, _
            UpdateLinks:=3
        'Workbooks(fileobjecta).Windows(1).WindowState = xlMinimized
        Workbooks(fileobjecta).Worksheets("Counter").Activate '指定激活的工作表名称
        worksheet_name = ActiveWorkbook.Name '返回活动工作薄的名称
        result = Mid(worksheet_name, 5, Len(worksheet_name) - InStr(".", worksheet_name) - 9) '处理字符串还原工作簿名称
        'Sheets("Counter").Select
        Range("B2:B43").Select '选中区域
        Selection.FormulaR1C1 = result '写入区域名称
        ActiveWorkbook.RefreshAll '刷新所有链接
        ActiveWorkbook.Save '保存
        ActiveWorkbook.Close '关闭
        fileobjecta = Dir '获取下一个文件
        Loop '结束循环体
End Sub
