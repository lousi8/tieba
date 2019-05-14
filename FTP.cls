VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, ByVal lpszRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExsiting As String, ByVal lpszNew As String) As Boolean
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession&, ByVal lpszDirectory$) As Boolean
Dim con_status  As Boolean
Dim con_point  As Long

Public Function ���ӷ�����(IP As String, �˿� As Integer, �ʺ� As String, ���� As String) As Boolean
    Dim ��ʼ��    As Long
    Const vbNullString As String = ""
    Const INTERNET_OPEN_TYPE_PRECONFIG As Long = 0
    ��ʼ�� = InternetOpen(vbNullString, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0&)
    con_point = InternetConnect(��ʼ��, IP, �˿�, �ʺ�, ����, 1, 0, 0)
    If con_point > 0 Then
        con_status = True
    Else
        con_status = False
    End If
    ���ӷ����� = con_status
End Function
                                                                    
Public Function �ϴ��ļ�(����·�� As String, Զ��·�� As String) As Boolean
    If con_status = True Then
        �ϴ��ļ� = FtpPutFile(con_point, ����·��, Զ��·��, 1, 0)
    Else
        MsgBox "����ȷ������"
    End If
End Function
                                                                    
Public Function �����ļ�(Զ��·�� As String, ����·�� As String) As Boolean
    If con_status = True Then
        �����ļ� = FtpGetFile(con_point, Զ��·��, ����·��, False, &H80000000, &H0, 0)
    Else
        MsgBox "����ȷ������"
    End If
End Function
                                                                    
Public Function ɾ���ļ�(ByVal �ļ�·��) As Boolean
    If con_status = True Then
        ɾ���ļ� = FtpDeleteFile(con_point, �ļ�·��)
    Else
        MsgBox "����ȷ������"
    End If
End Function
                                                                    
Public Function ����Ŀ¼(ByVal ·�� As String) As String
    If con_status = True Then
        ����Ŀ¼ = FtpCreateDirectory(con_point, ·��)
    Else
        MsgBox "����ȷ������"
    End If
End Function
                                                                    
Public Function �ļ�����(ByVal ԭ�ļ��� As String, ByVal ���ļ��� As String) As String
    If con_status = True Then
        �ļ����� = FtpRenameFile(con_point, ԭ�ļ���, ���ļ���)
    Else
        MsgBox "����ȷ������"
    End If
End Function
                                                                    

Public Function �ر�����()
Dim iStatus As Integer
iStatus = InternetCloseHandle(con_point)
End Function