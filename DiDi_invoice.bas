Attribute VB_Name = "DiDi_invoice"
Sub ProcessMailAndCallOCR()
    Dim olApp As Outlook.Application
    Dim olMail As Outlook.mailItem
    Dim olExplorer As Outlook.Explorer
    Dim olSelection As Outlook.Selection
    Dim olAttachments As Outlook.Attachments
    Dim PR_SEARCH_KEY As String
    Dim tempFolder As String
    Dim targetFolder As String
    Dim destinationFolder As String
    Dim logFile As String
    Dim pdfFilePath As String
    Dim base64Content As String
    Dim OCRResult As String
    Dim objProp As PropertyAccessor
    Dim logContent As String
    Dim startDate As String
    Dim totalAmount As String
    Dim regExp As Object
    Dim matches As Object
    Dim fileSystem As Object
    Dim iniFile As Object
    Dim configPath As String
    Dim oldFilePaths As Variant
    Dim oldFilePath As String
    Dim destinationFilePath As String
    Dim newFilePath As String
    Const PR_SEARCH_KEY_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x300B0102"
    
    On Error Resume Next
    
    ' 初始化 Outlook 对象
    Set olApp = Outlook.Application
    Set olExplorer = olApp.ActiveExplorer
    Set olSelection = olExplorer.Selection
    
    ' 检查是否选中邮件
    If olSelection.Count = 0 Then
        MsgBox "请选中一封邮件后再运行程序。", vbExclamation
        Exit Sub
    End If
    
    ' 获取选中的邮件
    Set olMail = olSelection.Item(1)
    
    ' 检查邮件标题是否包含“行程报销单”
    If InStr(1, olMail.Subject, "行程报销单", vbTextCompare) = 0 Then
        MsgBox "邮件标题不包含'行程报销单'字样，请检查后再试。", vbExclamation
        Exit Sub
    End If
    
    ' 获取 PR_SEARCH_KEY
    Set objProp = olMail.PropertyAccessor
    PR_SEARCH_KEY = objProp.BinaryToString(objProp.GetProperty(PR_SEARCH_KEY_ID))
    If PR_SEARCH_KEY = "" Then
        MsgBox "无法获取邮件的 PR_SEARCH_KEY 属性。", vbExclamation
        Exit Sub
    End If
    
    ' 获取 %temp% 文件夹路径
    tempFolder = Environ("TEMP")
    If tempFolder = "" Then
        MsgBox "无法获取系统临时文件夹路径。", vbExclamation
        Exit Sub
    End If
    
    ' 创建目标文件夹路径
    targetFolder = tempFolder & "\" & PR_SEARCH_KEY
    If Dir(targetFolder, vbDirectory) = "" Then
        MkDir targetFolder
    End If
    
    ' 获取附件集合并保存
    Set olAttachments = olMail.Attachments
    If olAttachments.Count = 0 Then
        MsgBox "当前邮件没有附件。", vbInformation
        Exit Sub
    End If
    
    For Each attachment In olAttachments
        attachment.SaveAsFile targetFolder & "\" & attachment.fileName
    Next attachment
    
    ' 检查指定 PDF 文件是否存在
    pdfFilePath = targetFolder & "\滴滴出行行程报销单.pdf"
    If Dir(pdfFilePath) = "" Then
        MsgBox "文件夹中未找到滴滴出行行程报销单.pdf。", vbExclamation
        Exit Sub
    End If
    
    ' 将 PDF 文件转换为 Base64
    base64Content = ConvertFileToBase64(pdfFilePath)
    If base64Content = "" Then
        MsgBox "PDF 文件转换为 Base64 失败。", vbExclamation
        Exit Sub
    End If
    
    ' 调用 Baidu OCR API
    OCRResult = CallBaiduOCR(base64Content)
    If OCRResult = "" Then
        MsgBox "调用 Baidu OCR API 失败。", vbExclamation
        Exit Sub
    End If
    
    ' 从配置文件中获取目标文件夹路径
    configPath = Environ("USERPROFILE") & "\OutlookPlugin\DiDiInvoice\config.ini"
    Set iniFile = CreateObject("Scripting.Dictionary")
    ParseINIFile configPath, iniFile
    
    If iniFile.Exists("destinationFolder") Then
        destinationFolder = iniFile("destinationFolder")
    Else
        MsgBox "配置文件中缺少 destinationFolder，请检查：" & vbCrLf & configPath, vbExclamation
        Exit Sub
    End If
    
    ' 检查目标文件夹路径是否存在
    If Dir(destinationFolder, vbDirectory) = "" Then
        MkDir destinationFolder
    End If
    
    ' 创建 log 文件并保存 OCR 结果
    logFile = targetFolder & "\" & PR_SEARCH_KEY & ".log"
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Dim logFileWriter As Object
    Set logFileWriter = fileSystem.CreateTextFile(logFile, True)
    logFileWriter.WriteLine "OCR API 返回结果:"
    logFileWriter.WriteLine OCRResult
    logFileWriter.Close
    
    ' 读取日志文件内容
    Dim logFileReader As Object
    Set logFileReader = fileSystem.OpenTextFile(logFile, 1)
    logContent = logFileReader.ReadAll
    logFileReader.Close
    
    ' 使用正则表达式提取关键信息
    Set regExp = CreateObject("VBScript.RegExp")
    regExp.Global = True
    regExp.IgnoreCase = True
    
    ' 提取日期
    regExp.Pattern = "行程起止日期：(\d{4}-\d{2}-\d{2})"
    Set matches = regExp.Execute(logContent)
    If matches.Count > 0 Then
        startDate = matches(0).SubMatches(0)
    Else
        startDate = "未知日期"
    End If
    
    ' 提取金额
    regExp.Pattern = "合计([\d\.]+)元"
    Set matches = regExp.Execute(logContent)
    If matches.Count > 0 Then
        totalAmount = matches(0).SubMatches(0)
    Else
        totalAmount = "未知金额"
    End If
    
    ' 重命名文件
    oldFilePaths = Array(targetFolder & "\滴滴出行行程报销单.pdf", targetFolder & "\滴滴电子发票.pdf")
    For i = LBound(oldFilePaths) To UBound(oldFilePaths)
        oldFilePath = oldFilePaths(i)
        If Dir(oldFilePath) <> "" Then
            newFilePath = targetFolder & "\" & startDate & "_" & totalAmount & "_" & Mid(oldFilePath, InStrRev(oldFilePath, "\") + 1)
            Name oldFilePath As newFilePath
            oldFilePaths(i) = newFilePath
        End If
    Next i
    
    ' 移动文件到目标文件夹
    For i = LBound(oldFilePaths) To UBound(oldFilePaths)
        oldFilePath = oldFilePaths(i)
        If Dir(oldFilePath) <> "" Then
            destinationFilePath = destinationFolder & "\" & Mid(oldFilePath, InStrRev(oldFilePath, "\") + 1)
            FileCopy oldFilePath, destinationFilePath
            Kill oldFilePath
        End If
    Next i
    
    ' 打开目标文件夹
    Shell "explorer.exe " & destinationFolder, vbNormalFocus
    
    ' 清理对象
    Set olAttachments = Nothing
    Set olMail = Nothing
    Set olExplorer = Nothing
    Set olSelection = Nothing
    Set objProp = Nothing
    Set fileSystem = Nothing
    Set regExp = Nothing
    Set matches = Nothing
End Sub


Sub ShowConfigForm()
    ' 显示窗体
    ConfigForm.Show
End Sub

