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

Public Function 连接服务器(IP As String, 端口 As Integer, 帐号 As String, 密码 As String) As Boolean
    Dim 初始化    As Long
    Const vbNullString As String = ""
    Const INTERNET_OPEN_TYPE_PRECONFIG As Long = 0
    初始化 = InternetOpen(vbNullString, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0&)
    con_point = InternetConnect(初始化, IP, 端口, 帐号, 密码, 1, 0, 0)
    If con_point > 0 Then
        con_status = True
    Else
        con_status = False
    End If
    连接服务器 = con_status
End Function
                                                                    
Public Function 上传文件(本地路径 As String, 远程路径 As String) As Boolean
    If con_status = True Then
        上传文件 = FtpPutFile(con_point, 本地路径, 远程路径, 1, 0)
    Else
        MsgBox "请先确认连接"
    End If
End Function
                                                                    
Public Function 下载文件(远程路径 As String, 本地路径 As String) As Boolean
    If con_status = True Then
        下载文件 = FtpGetFile(con_point, 远程路径, 本地路径, False, &H80000000, &H0, 0)
    Else
        MsgBox "请先确认连接"
    End If
End Function
                                                                    
Public Function 删除文件(ByVal 文件路径) As Boolean
    If con_status = True Then
        删除文件 = FtpDeleteFile(con_point, 文件路径)
    Else
        MsgBox "请先确认连接"
    End If
End Function
                                                                    
Public Function 创建目录(ByVal 路径 As String) As String
    If con_status = True Then
        创建目录 = FtpCreateDirectory(con_point, 路径)
    Else
        MsgBox "请先确认连接"
    End If
End Function
                                                                    
Public Function 文件改名(ByVal 原文件名 As String, ByVal 新文件名 As String) As String
    If con_status = True Then
        文件改名 = FtpRenameFile(con_point, 原文件名, 新文件名)
    Else
        MsgBox "请先确认连接"
    End If
End Function
                                                                    

Public Function 关闭连接()
Dim iStatus As Integer
iStatus = InternetCloseHandle(con_point)
End Function
