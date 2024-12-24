VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConfigForm 
   Caption         =   "����"
   ClientHeight    =   5304
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "ConfigForm.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "ConfigForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' �������ʱ��ʼ��
Private Sub UserForm_Initialize()
    Dim configFilePath As String
    Dim iniFile As Object

    ' �����ļ�·��
    configFilePath = Environ("USERPROFILE") & "\\OutlookPlugin\\DiDiInvoice\\config.ini"

    ' ��������ļ��Ƿ����
    If Dir(configFilePath) = "" Then
        ' �����������ļ�
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

    ' ��ȡ�����ļ�����
    Set iniFile = CreateObject("Scripting.Dictionary")
    Call ParseINIFile(configFilePath, iniFile)

    ' �����������ı���
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

' �����ť����¼�
Private Sub Buttonbrowser_Click()
    Dim folderPath As String
    folderPath = BrowseForFolder("��ѡ���ļ���")
    If folderPath <> "" Then
        TextBox1.text = folderPath
    End If
End Sub

Private Sub Buttonsubmit_Click()
    Dim configFilePath As String
    Dim stream As Object

    ' �����ļ�·��
    configFilePath = Environ("USERPROFILE") & "\\OutlookPlugin\\DiDiInvoice\\config.ini"

    ' ʹ�� ADODB.Stream �� UTF-8 ����д�������ļ�
    On Error GoTo ErrorHandler
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' �ı�����
    stream.charSet = "utf-8" ' ָ������Ϊ UTF-8
    stream.Open

    ' д������
    stream.WriteText "destinationFolder=" & TextBox1.text & vbCrLf
    stream.WriteText "clientId=" & TextBox2.text & vbCrLf
    stream.WriteText "clientSecret=" & TextBox3.text & vbCrLf

    stream.SaveToFile configFilePath, 2 ' ����д��
    stream.Close

    MsgBox "�����ѱ��棡", vbInformation

    ' �رմ���
    Unload Me
    Exit Sub

ErrorHandler:
    MsgBox "���ñ���ʧ�ܣ������ļ�·����Ȩ�����⡣", vbExclamation
End Sub


' ȡ����ť����¼�
Private Sub Buttoncancel_Click()
    ' ���ٴ���
    Unload Me
End Sub

' ʹ�� Shell API ��ʾ�ļ���ѡ��Ի���
Function BrowseForFolder(prompt As String) As String
    Dim shellApp As Object
    Dim folder As Object

    ' ���� Shell ����
    Set shellApp = CreateObject("Shell.Application")

    ' ��ʾ�ļ���ѡ��Ի���
    Set folder = shellApp.BrowseForFolder(0, prompt, &H1)

    ' ����û��Ƿ�ѡ�����ļ���
    If Not folder Is Nothing Then
        BrowseForFolder = folder.Self.Path
    Else
        BrowseForFolder = ""
    End If

    Set folder = Nothing
    Set shellApp = Nothing
End Function

