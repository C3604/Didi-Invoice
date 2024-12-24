VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConfigForm 
   Caption         =   "配置"
   ClientHeight    =   5304
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "ConfigForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "ConfigForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' 窗体加载时初始化
Private Sub UserForm_Initialize()
    Dim configFilePath As String
    Dim iniFile As Object

    ' 配置文件路径
    configFilePath = Environ("USERPROFILE") & "\\OutlookPlugin\\DiDiInvoice\\config.ini"

    ' 检查配置文件是否存在
    If Dir(configFilePath) = "" Then
        ' 创建空配置文件
        Dim fileSystem As Object
        Set fileSystem = CreateObject("Scripting.FileSystemObject")
        Dim newFile As Object
        Set newFile = fileSystem.CreateTextFile(configFilePath, True)
        newFile.WriteLine "destinationFolder="
        newFile.WriteLine "clientId="
        newFile.WriteLine "clientSecret="
        newFile.Close
        Set fileSystem = Nothing
    End If

    ' 读取配置文件内容
    Set iniFile = CreateObject("Scripting.Dictionary")
    Call ParseINIFile(configFilePath, iniFile)

    ' 将内容填入文本框
    If iniFile.Exists("destinationFolder") Then
        TextBox1.text = iniFile("destinationFolder")
    End If
    If iniFile.Exists("clientId") Then
        TextBox2.text = iniFile("clientId")
    End If
    If iniFile.Exists("clientSecret") Then
        TextBox3.text = iniFile("clientSecret")
    End If
End Sub

' 浏览按钮点击事件
Private Sub Buttonbrowser_Click()
    Dim folderPath As String
    folderPath = BrowseForFolder("请选择文件夹")
    If folderPath <> "" Then
        TextBox1.text = folderPath
    End If
End Sub

Private Sub Buttonsubmit_Click()
    Dim configFilePath As String
    Dim stream As Object

    ' 配置文件路径
    configFilePath = Environ("USERPROFILE") & "\\OutlookPlugin\\DiDiInvoice\\config.ini"

    ' 使用 ADODB.Stream 以 UTF-8 编码写入配置文件
    On Error GoTo ErrorHandler
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' 文本类型
    stream.charSet = "utf-8" ' 指定编码为 UTF-8
    stream.Open

    ' 写入内容
    stream.WriteText "destinationFolder=" & TextBox1.text & vbCrLf
    stream.WriteText "clientId=" & TextBox2.text & vbCrLf
    stream.WriteText "clientSecret=" & TextBox3.text & vbCrLf

    stream.SaveToFile configFilePath, 2 ' 覆盖写入
    stream.Close

    MsgBox "配置已保存！", vbInformation

    ' 关闭窗体
    Unload Me
    Exit Sub

ErrorHandler:
    MsgBox "配置保存失败，请检查文件路径或权限问题。", vbExclamation
End Sub


' 取消按钮点击事件
Private Sub Buttoncancel_Click()
    ' 销毁窗体
    Unload Me
End Sub

' 使用 Shell API 显示文件夹选择对话框
Function BrowseForFolder(prompt As String) As String
    Dim shellApp As Object
    Dim folder As Object

    ' 创建 Shell 对象
    Set shellApp = CreateObject("Shell.Application")

    ' 显示文件夹选择对话框
    Set folder = shellApp.BrowseForFolder(0, prompt, &H1)

    ' 检查用户是否选择了文件夹
    If Not folder Is Nothing Then
        BrowseForFolder = folder.Self.Path
    Else
        BrowseForFolder = ""
    End If

    Set folder = Nothing
    Set shellApp = Nothing
End Function

