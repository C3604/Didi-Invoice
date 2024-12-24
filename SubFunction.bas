Attribute VB_Name = "SubFunction"
' ���ļ�ת��ΪBase64
Function ConvertFileToBase64(filePath As String) As String
    Dim fileStream As Object
    Dim binaryData() As Byte
    Dim base64Data As String
    
    On Error Resume Next
    Set fileStream = CreateObject("ADODB.Stream")
    fileStream.Type = 1 ' ������
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
' ����Baidu OCR API��֧�ִ������ļ�����ƾ�ݣ�
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
    
    ' �����ļ�·��
    configPath = Environ("USERPROFILE") & "\OutlookPlugin\DiDiInvoice\config.ini"
    
    ' ��������ļ��Ƿ����
    If Dir(configPath) = "" Then
        MsgBox "�����ļ������ڣ���ȷ���ļ�·����ȷ��" & vbCrLf & configPath, vbExclamation
        CallBaiduOCR = ""
        Exit Function
    End If
    
    ' ��ȡ�����ļ�
    Set iniFile = CreateObject("Scripting.Dictionary")
    ParseINIFile configPath, iniFile
    
    ' �������ļ��л�ȡclientId��clientSecret
    If iniFile.Exists("clientId") And iniFile.Exists("clientSecret") Then
        clientId = iniFile("clientId")
        clientSecret = iniFile("clientSecret")
    Else
        MsgBox "�����ļ���ȱ��clientId��clientSecret�����飺" & vbCrLf & configPath, vbExclamation
        CallBaiduOCR = ""
        Exit Function
    End If
    
    ' ��ȡ Access Token
    tokenURL = "https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id=" & clientId & "&client_secret=" & clientSecret
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", tokenURL, False
    http.setRequestHeader "Content-Type", "application/json"
    http.Send
    
    If http.Status <> 200 Then
        MsgBox "��ȡ Access Token ʧ��: " & http.Status & " - " & http.statusText, vbExclamation
        CallBaiduOCR = ""
        Exit Function
    End If
    
    ' ���� Access Token
    Set JSON = JsonConverter.ParseJSON(http.responseText)
    accessToken = JSON("access_token")
    
    ' OCR API URL
    OCRURL = "https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic?access_token=" & accessToken
    
    ' �������
    postData = "pdf_file=" & URLEncode(base64Content) & _
               "&detect_direction=false" & _
               "&paragraph=false" & _
               "&probability=false" & _
               "&multidirectional_recognize=false"
    
    ' ���� HTTP ����
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", OCRURL, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.setRequestHeader "Accept", "application/json"
    http.Send postData
    
    ' �����Ӧ״̬
    If http.Status = 200 Then
        CallBaiduOCR = http.responseText
    Else
        MsgBox "OCR ����ʧ��: " & http.Status & " - " & http.statusText, vbExclamation
        CallBaiduOCR = ""
    End If
End Function

' ����INI�ļ�Ϊ��ֵ��
Sub ParseINIFile(filePath As String, ByRef iniDict As Object)
    Dim stream As Object
    Dim fileContent As String
    Dim lines() As String
    Dim line As Variant
    Dim keyValue() As String

    ' ʹ�� ADODB.Stream ��ȡ�����ļ�
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' �ı�����
    stream.charSet = "utf-8" ' ָ������Ϊ UTF-8
    stream.Open
    stream.LoadFromFile filePath
    fileContent = stream.ReadText
    stream.Close
    Set stream = Nothing

    ' ���н�������
    lines = Split(fileContent, vbCrLf)
    For Each line In lines
        If InStr(line, "=") > 0 Then
            keyValue = Split(line, "=")
            iniDict(Trim(keyValue(0))) = Trim(keyValue(1))
        End If
    Next
End Sub




' URL ���뺯�������ֲ��䣩
Function URLEncode(text As String) As String
    Dim i As Long
    Dim char As String
    Dim encoded As String
    
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        Select Case Asc(char)
            Case 48 To 57, 65 To 90, 97 To 122 ' ��ĸ������
                encoded = encoded & char
            Case Else
                encoded = encoded & "%" & Right("0" & Hex(Asc(char)), 2)
        End Select
    Next i
    
    URLEncode = encoded
End Function

' ���������ݱ���ΪBase64
Function EncodeBase64(binaryData() As Byte) As String
    Dim xmlDoc As Object
    Dim xmlNode As Object
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set xmlNode = xmlDoc.createElement("Base64")
    xmlNode.dataType = "bin.base64"
    xmlNode.nodeTypedValue = binaryData
    EncodeBase64 = xmlNode.text
End Function

