Sub 移除Excel外部链接()
    CurrentPath = "?"
    fileobjecta = Dir(CurrentPath) '获取指定目录下第一个文件完整名称
    Application.Visible = False
    Do While fileobjecta <> ""
        filepath = CurrentPath & fileobjecta
        Excel.Workbooks.Open FileName:= _
        filepath, _
        UpdateLinks:=3
        Set objWorkbook = Excel.ActiveWorkbook
        For Each LinkSource In objWorkbook.LinkSources
            objWorkbook.BreakLink LinkSource, 1 '此处 LinkSourses 为链接到外部excel的地址; 1即为要处理断开链接的类型 即为EXCEL 数据源
            Next
        objWorkbook.Save
        objWorkbook.Close
        fileobjecta = Dir
    Loop
End Sub
