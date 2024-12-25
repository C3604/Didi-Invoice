Attribute VB_Name = "SubFunction"
Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Declare PtrSafe Function SetForegroundWindow Lib "user32" _
    (ByVal hWnd As Long) As Long
' 将文件转换为Base64
Function ConvertFileToBase64(filePath As String) As String
    Dim fileStream As Object
    Dim binaryData() As Byte
    Dim base64Data As String
    
    On Error Resume Next
    Set fileStream = CreateObject("ADODB.Stream")
    fileStream.Type = 1 ' 二进制
    fileStream.Open
    fileStream.LoadFromFile filePath
    binaryData = fileStream.Read
    fileStream.Close
    Set fileStream = Nothing
    
    If IsEmpty(binaryData) Then
        ConvertFileToBase64 = ""
        Exit Function
    End If
    
    ConvertFileToBase64 = EncodeBase64(binaryData)
End Function
' 调用Baidu OCR API（支持从配置文件加载凭据）
Function CallBaiduOCR(base64Content As String) As String
    Dim http As Object
    Dim clientId As String
    Dim clientSecret As String
    Dim tokenURL As String
    Dim accessToken As String
    Dim OCRURL As String
    Dim postData As String
    Dim JSON As Object
    Dim configPath As String
    Dim iniFile As Object
    Dim responseText As String
    
    ' 配置文件路径
    configPath = Environ("USERPROFILE") & "\OutlookPlugin\DiDiInvoice\config.ini"
    
    ' 检查配置文件是否存在
    If Dir(configPath) = "" Then
        MsgBox "配置文件不存在，请确保文件路径正确：" & vbCrLf & configPath, vbExclamation
        CallBaiduOCR = ""
        Exit Function
    End If
    
    ' 读取配置文件
    Set iniFile = CreateObject("Scripting.Dictionary")
    ParseINIFile configPath, iniFile
    
    ' 从配置文件中获取clientId和clientSecret
    If iniFile.Exists("clientId") And iniFile.Exists("clientSecret") Then
        clientId = iniFile("clientId")
        clientSecret = iniFile("clientSecret")
    Else
        MsgBox "配置文件中缺少clientId或clientSecret，请检查：" & vbCrLf & configPath, vbExclamation
        CallBaiduOCR = ""
        Exit Function
    End If
    
    ' 获取 Access Token
    tokenURL = "https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id=" & clientId & "&client_secret=" & clientSecret
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", tokenURL, False
    http.setRequestHeader "Content-Type", "application/json"
    http.Send
    
    If http.Status <> 200 Then
        MsgBox "获取 Access Token 失败: " & http.Status & " - " & http.statusText, vbExclamation
        CallBaiduOCR = ""
        Exit Function
    End If
    
    ' 解析 Access Token
    Set JSON = JsonConverter.ParseJSON(http.responseText)
    accessToken = JSON("access_token")
    
    ' OCR API URL
    OCRURL = "https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic?access_token=" & accessToken
    
    ' 请求参数
    postData = "pdf_file=" & URLEncode(base64Content) & _
               "&detect_direction=false" & _
               "&paragraph=false" & _
               "&probability=false" & _
               "&multidirectional_recognize=false"
    
    ' 创建 HTTP 请求
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", OCRURL, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.setRequestHeader "Accept", "application/json"
    http.Send postData
    
    ' 检查响应状态
    If http.Status = 200 Then
        CallBaiduOCR = http.responseText
    Else
        MsgBox "OCR 请求失败: " & http.Status & " - " & http.statusText, vbExclamation
        CallBaiduOCR = ""
    End If
End Function

' 解析INI文件为键值对
Sub ParseINIFile(filePath As String, ByRef iniDict As Object)
    Dim stream As Object
    Dim fileContent As String
    Dim lines() As String
    Dim line As Variant
    Dim keyValue() As String

    ' 使用 ADODB.Stream 读取配置文件
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' 文本类型
    stream.charSet = "utf-8" ' 指定编码为 UTF-8
    stream.Open
    stream.LoadFromFile filePath
    fileContent = stream.ReadText
    stream.Close
    Set stream = Nothing

    ' 按行解析内容
    lines = Split(fileContent, vbCrLf)
    For Each line In lines
        If InStr(line, "=") > 0 Then
            keyValue = Split(line, "=")
            iniDict(Trim(keyValue(0))) = Trim(keyValue(1))
        End If
    Next
End Sub




' URL 编码函数（保持不变）
Function URLEncode(text As String) As String
    Dim i As Long
    Dim char As String
    Dim encoded As String
    
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        Select Case Asc(char)
            Case 48 To 57, 65 To 90, 97 To 122 ' 字母和数字
                encoded = encoded & char
            Case Else
                encoded = encoded & "%" & Right("0" & Hex(Asc(char)), 2)
        End Select
    Next i
    
    URLEncode = encoded
End Function

' 二进制数据编码为Base64
Function EncodeBase64(binaryData() As Byte) As String
    Dim xmlDoc As Object
    Dim xmlNode As Object
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set xmlNode = xmlDoc.createElement("Base64")
    xmlNode.dataType = "bin.base64"
    xmlNode.nodeTypedValue = binaryData
    EncodeBase64 = xmlNode.text
End Function



Sub OpenOrActivateFolder(destinationFolder As String)
    Dim folderPath As String
    Dim hWnd As Long
    Dim folderName As String

    ' 获取文件夹名称
    folderPath = Replace(destinationFolder, "/", "\") ' 确保路径为反斜杠
    If Right(folderPath, 1) = "\" Then
        folderName = Mid(folderPath, InStrRev(folderPath, "\", Len(folderPath) - 1) + 1)
    Else
        folderName = Mid(folderPath, InStrRev(folderPath, "\") + 1)
    End If

    ' 检查文件夹窗口是否已经打开
    hWnd = FindWindow("CabinetWClass", folderName)

    If hWnd <> 0 Then
        ' 如果窗口已经打开，将其激活到前台
        SetForegroundWindow hWnd
    Else
        ' 如果窗口未打开，则新建窗口打开文件夹
        Shell "explorer.exe " & folderPath, vbNormalFocus
    End If
End Sub
