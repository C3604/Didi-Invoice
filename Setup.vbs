On Error Resume Next

' 定义变量
Dim objOutlook, objVBE, filePath, importPath, objFSO, currentFolder

' 创建 Outlook 应用程序对象
Set objOutlook = CreateObject("Outlook.Application")

' 尝试获取 VBE 对象
Set objVBE = Nothing
Set objVBE = objOutlook.VBE.ActiveVBProject

If objVBE Is Nothing Then
    WScript.Echo "无法访问 VBA 项目。请检查以下设置：" & vbCrLf & _
                 "1. 确保启用了对 VBA 项目对象模型的访问。" & vbCrLf & _
                 "2. 确保脚本权限正确。" & vbCrLf & _
                 "3. 启用 Outlook 后再运行此脚本。"
    WScript.Quit
End If

' 获取当前脚本所在目录
Set objFSO = CreateObject("Scripting.FileSystemObject")
currentFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)

' 设置模块和窗体文件所在目录为当前目录下的 "VBA Code"
importPath = currentFolder & "\VBA Code\"

' 检查目录是否存在
If Not objFSO.FolderExists(importPath) Then
    WScript.Echo "目录未找到: " & importPath
    WScript.Quit
End If

' 导入模块和窗体
filePath = importPath & "DiDi_invoice.bas"
If objFSO.FileExists(filePath) Then
    objVBE.VBComponents.Import filePath
Else
    WScript.Echo "文件未找到: " & filePath
End If

filePath = importPath & "SubFunction.bas"
If objFSO.FileExists(filePath) Then
    objVBE.VBComponents.Import filePath
Else
    WScript.Echo "文件未找到: " & filePath
End If

filePath = importPath & "JsonConverter.bas"
If objFSO.FileExists(filePath) Then
    objVBE.VBComponents.Import filePath
Else
    WScript.Echo "文件未找到: " & filePath
End If

filePath = importPath & "ConfigForm.frm"
If objFSO.FileExists(filePath) Then
    objVBE.VBComponents.Import filePath
Else
    WScript.Echo "文件未找到: " & filePath
End If

' 提示完成
WScript.Echo "所有模块和窗体已成功导入到 Outlook VBA 项目中！"