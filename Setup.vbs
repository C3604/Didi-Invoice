On Error Resume Next

' �������
Dim objOutlook, objVBE, filePath, importPath, objFSO, currentFolder

' ���� Outlook Ӧ�ó������
Set objOutlook = CreateObject("Outlook.Application")

' ���Ի�ȡ VBE ����
Set objVBE = Nothing
Set objVBE = objOutlook.VBE.ActiveVBProject

If objVBE Is Nothing Then
    WScript.Echo "�޷����� VBA ��Ŀ�������������ã�" & vbCrLf & _
                 "1. ȷ�������˶� VBA ��Ŀ����ģ�͵ķ��ʡ�" & vbCrLf & _
                 "2. ȷ���ű�Ȩ����ȷ��" & vbCrLf & _
                 "3. ���� Outlook �������д˽ű���"
    WScript.Quit
End If

' ��ȡ��ǰ�ű�����Ŀ¼
Set objFSO = CreateObject("Scripting.FileSystemObject")
currentFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)

' ����ģ��ʹ����ļ�����Ŀ¼Ϊ��ǰĿ¼�µ� "VBA Code"
importPath = currentFolder & "\VBA Code\"

' ���Ŀ¼�Ƿ����
If Not objFSO.FolderExists(importPath) Then
    WScript.Echo "Ŀ¼δ�ҵ�: " & importPath
    WScript.Quit
End If

' ����ģ��ʹ���
filePath = importPath & "DiDi_invoice.bas"
If objFSO.FileExists(filePath) Then
    objVBE.VBComponents.Import filePath
Else
    WScript.Echo "�ļ�δ�ҵ�: " & filePath
End If

filePath = importPath & "SubFunction.bas"
If objFSO.FileExists(filePath) Then
    objVBE.VBComponents.Import filePath
Else
    WScript.Echo "�ļ�δ�ҵ�: " & filePath
End If

filePath = importPath & "JsonConverter.bas"
If objFSO.FileExists(filePath) Then
    objVBE.VBComponents.Import filePath
Else
    WScript.Echo "�ļ�δ�ҵ�: " & filePath
End If

filePath = importPath & "ConfigForm.frm"
If objFSO.FileExists(filePath) Then
    objVBE.VBComponents.Import filePath
Else
    WScript.Echo "�ļ�δ�ҵ�: " & filePath
End If

' ��ʾ���
WScript.Echo "����ģ��ʹ����ѳɹ����뵽 Outlook VBA ��Ŀ�У�"