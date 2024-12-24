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
    
    ' ��ʼ�� Outlook ����
    Set olApp = Outlook.Application
    Set olExplorer = olApp.ActiveExplorer
    Set olSelection = olExplorer.Selection
    
    ' ����Ƿ�ѡ���ʼ�
    If olSelection.Count = 0 Then
        MsgBox "��ѡ��һ���ʼ��������г���", vbExclamation
        Exit Sub
    End If
    
    ' ��ȡѡ�е��ʼ�
    Set olMail = olSelection.Item(1)
    
    ' ����ʼ������Ƿ�������г̱�������
    If InStr(1, olMail.Subject, "�г̱�����", vbTextCompare) = 0 Then
        MsgBox "�ʼ����ⲻ����'�г̱�����'��������������ԡ�", vbExclamation
        Exit Sub
    End If
    
    ' ��ȡ PR_SEARCH_KEY
    Set objProp = olMail.PropertyAccessor
    PR_SEARCH_KEY = objProp.BinaryToString(objProp.GetProperty(PR_SEARCH_KEY_ID))
    If PR_SEARCH_KEY = "" Then
        MsgBox "�޷���ȡ�ʼ��� PR_SEARCH_KEY ���ԡ�", vbExclamation
        Exit Sub
    End If
    
    ' ��ȡ %temp% �ļ���·��
    tempFolder = Environ("TEMP")
    If tempFolder = "" Then
        MsgBox "�޷���ȡϵͳ��ʱ�ļ���·����", vbExclamation
        Exit Sub
    End If
    
    ' ����Ŀ���ļ���·��
    targetFolder = tempFolder & "\" & PR_SEARCH_KEY
    If Dir(targetFolder, vbDirectory) = "" Then
        MkDir targetFolder
    End If
    
    ' ��ȡ�������ϲ�����
    Set olAttachments = olMail.Attachments
    If olAttachments.Count = 0 Then
        MsgBox "��ǰ�ʼ�û�и�����", vbInformation
        Exit Sub
    End If
    
    For Each attachment In olAttachments
        attachment.SaveAsFile targetFolder & "\" & attachment.fileName
    Next attachment
    
    ' ���ָ�� PDF �ļ��Ƿ����
    pdfFilePath = targetFolder & "\�εγ����г̱�����.pdf"
    If Dir(pdfFilePath) = "" Then
        MsgBox "�ļ�����δ�ҵ��εγ����г̱�����.pdf��", vbExclamation
        Exit Sub
    End If
    
    ' �� PDF �ļ�ת��Ϊ Base64
    base64Content = ConvertFileToBase64(pdfFilePath)
    If base64Content = "" Then
        MsgBox "PDF �ļ�ת��Ϊ Base64 ʧ�ܡ�", vbExclamation
        Exit Sub
    End If
    
    ' ���� Baidu OCR API
    OCRResult = CallBaiduOCR(base64Content)
    If OCRResult = "" Then
        MsgBox "���� Baidu OCR API ʧ�ܡ�", vbExclamation
        Exit Sub
    End If
    
    ' �������ļ��л�ȡĿ���ļ���·��
    configPath = Environ("USERPROFILE") & "\OutlookPlugin\DiDiInvoice\config.ini"
    Set iniFile = CreateObject("Scripting.Dictionary")
    ParseINIFile configPath, iniFile
    
    If iniFile.Exists("destinationFolder") Then
        destinationFolder = iniFile("destinationFolder")
    Else
        MsgBox "�����ļ���ȱ�� destinationFolder�����飺" & vbCrLf & configPath, vbExclamation
        Exit Sub
    End If
    
    ' ���Ŀ���ļ���·���Ƿ����
    If Dir(destinationFolder, vbDirectory) = "" Then
        MkDir destinationFolder
    End If
    
    ' ���� log �ļ������� OCR ���
    logFile = targetFolder & "\" & PR_SEARCH_KEY & ".log"
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Dim logFileWriter As Object
    Set logFileWriter = fileSystem.CreateTextFile(logFile, True)
    logFileWriter.WriteLine "OCR API ���ؽ��:"
    logFileWriter.WriteLine OCRResult
    logFileWriter.Close
    
    ' ��ȡ��־�ļ�����
    Dim logFileReader As Object
    Set logFileReader = fileSystem.OpenTextFile(logFile, 1)
    logContent = logFileReader.ReadAll
    logFileReader.Close
    
    ' ʹ��������ʽ��ȡ�ؼ���Ϣ
    Set regExp = CreateObject("VBScript.RegExp")
    regExp.Global = True
    regExp.IgnoreCase = True
    
    ' ��ȡ����
    regExp.Pattern = "�г���ֹ���ڣ�(\d{4}-\d{2}-\d{2})"
    Set matches = regExp.Execute(logContent)
    If matches.Count > 0 Then
        startDate = matches(0).SubMatches(0)
    Else
        startDate = "δ֪����"
    End If
    
    ' ��ȡ���
    regExp.Pattern = "�ϼ�([\d\.]+)Ԫ"
    Set matches = regExp.Execute(logContent)
    If matches.Count > 0 Then
        totalAmount = matches(0).SubMatches(0)
    Else
        totalAmount = "δ֪���"
    End If
    
    ' �������ļ�
    oldFilePaths = Array(targetFolder & "\�εγ����г̱�����.pdf", targetFolder & "\�εε��ӷ�Ʊ.pdf")
    For i = LBound(oldFilePaths) To UBound(oldFilePaths)
        oldFilePath = oldFilePaths(i)
        If Dir(oldFilePath) <> "" Then
            newFilePath = targetFolder & "\" & startDate & "_" & totalAmount & "_" & Mid(oldFilePath, InStrRev(oldFilePath, "\") + 1)
            Name oldFilePath As newFilePath
            oldFilePaths(i) = newFilePath
        End If
    Next i
    
    ' �ƶ��ļ���Ŀ���ļ���
    For i = LBound(oldFilePaths) To UBound(oldFilePaths)
        oldFilePath = oldFilePaths(i)
        If Dir(oldFilePath) <> "" Then
            destinationFilePath = destinationFolder & "\" & Mid(oldFilePath, InStrRev(oldFilePath, "\") + 1)
            FileCopy oldFilePath, destinationFilePath
            Kill oldFilePath
        End If
    Next i
    
    ' ��Ŀ���ļ���
    Shell "explorer.exe " & destinationFolder, vbNormalFocus
    
    ' �������
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
    ' ��ʾ����
    ConfigForm.Show
End Sub

