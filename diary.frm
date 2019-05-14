VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "百度贴吧下载器"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7905
   Icon            =   "diary.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   7905
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txt_path 
      Height          =   375
      Left            =   1500
      TabIndex        =   16
      Top             =   1800
      Width           =   4695
   End
   Begin VB.TextBox txt_id1 
      Height          =   375
      Left            =   1500
      TabIndex        =   15
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txt_name 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1500
      TabIndex        =   14
      Top             =   300
      Width           =   4755
   End
   Begin VB.TextBox txt_filter 
      Height          =   375
      Left            =   1500
      TabIndex        =   13
      Top             =   1380
      Width           =   4695
   End
   Begin VB.TextBox txt_id 
      Height          =   375
      Left            =   1500
      TabIndex        =   12
      Top             =   600
      Width           =   4755
   End
   Begin VB.CheckBox chk_add 
      Caption         =   "补页"
      Height          =   180
      Left            =   6660
      TabIndex        =   11
      Top             =   660
      Width           =   1155
   End
   Begin VB.CommandButton cmd_simple 
      Caption         =   "网页精简"
      Enabled         =   0   'False
      Height          =   435
      Left            =   6600
      TabIndex        =   9
      Top             =   1260
      Width           =   1095
   End
   Begin VB.CheckBox chk_simple 
      Caption         =   "即时精简"
      Height          =   315
      Left            =   6660
      TabIndex        =   8
      Top             =   900
      Value           =   2  'Grayed
      Width           =   1155
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5820
      Top             =   -180
   End
   Begin VB.TextBox txt_id2 
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   960
      Width           =   2595
   End
   Begin VB.CommandButton cmd_go 
      Caption         =   "开始备份"
      Height          =   435
      Left            =   6600
      TabIndex        =   4
      Top             =   180
      Width           =   1095
   End
   Begin VB.CommandButton Cmd_Exit 
      Caption         =   "退出(&E)"
      Height          =   435
      Left            =   6600
      TabIndex        =   3
      Top             =   1740
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "指定帖子号"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   660
      Width           =   1155
   End
   Begin VB.Label Label4 
      Caption         =   "标题过滤"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label label_new 
      Caption         =   "百度贴吧名"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   1275
   End
   Begin VB.Label label3 
      Caption         =   "页码范围"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "保存路径"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1860
      Width           =   1035
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   300
      TabIndex        =   0
      Top             =   2280
      Width           =   7515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "Kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function internetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" (ByVal hinternet As Long, ByVal dwoption As Long, ByRef lpBuffer As Any, ByVal dwbufferlength As Long) As Long

Private Type INTERNET_PROXY_INFO
dwAccessType As Long
lpszProxy As String
lpszProxyBypass As String
End Type
 
Dim urlOK As Long
Dim ServerXmlHttp As MSXML2.ServerXmlHttp
Dim proxyAddress As String 'IP:PORT
Dim wrongURL As String

Sub setProxy(proxyAddr As String, Optional mode As String = "HTTP")

If proxyAddr = "" Then Exit Sub
Const INTERNET_OPEN_TYPE_PRECONFIG = 0 'use registry configuration
Const INTERNET_OPEN_TYPE_DIRECT = 1    'direct to net
Const INTERNET_OPEN_TYPE_PROXY = 3     'via named proxy
Const INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY = 4   'prevent using java/script/INS
Const INTERNET_OPTION_PROXY = 38
Const INTERNET_OPTION_SETTINGS_CHANGED = 39

Dim options As INTERNET_PROXY_INFO
options.dwAccessType = INTERNET_OPEN_TYPE_PROXY
options.lpszProxy = mode & "=" & proxyAddr '"HTTP=IP:PORT"
options.lpszProxyBypass = ""
internetSetOption 0, INTERNET_OPTION_PROXY, options, LenB(options)
internetSetOption 0, INTERNET_OPTION_SETTINGS_CHANGED, 0, 0

End Sub

Private Sub Form_Load()
  Call init_Form
End Sub

Private Sub Cmd_Exit_Click()
Unload Me
End Sub

Private Sub cmd_simple_Click()
Dim sPath As String
Dim i As Long
If InStr(txt_name.Text, ",") = 0 Then
sPath = txt_path.Text & txt_name.Text & "\"
i = simple_webFile(sPath, 2, 1)
Label1.Caption = "精简了" & i & "个网页"
End If
End Sub


Public Sub init_Form()
Dim iniTxt As String
iniTxt = readutf8("ini.txt")
If iniTxt <> "" Then
  txt_name.Text = Fetch(iniTxt, "name=[", "]")
  txt_id.Text = Fetch(iniTxt, "id=[", "]")
  txt_id1.Text = Fetch(iniTxt, "id1=[", "]")
  txt_id2.Text = Fetch(iniTxt, "id2=[", "]")
  txt_filter.Text = Fetch(iniTxt, "filter=[", "]")
  txt_path.Text = Fetch(iniTxt, "path=[", "]")
  proxyAddress = Fetch(iniTxt, "proxy=[", "]")
Else
  txt_name.Text = "太阳的后裔"
  txt_id.Text = ""
  txt_id1.Text = ""
  txt_id2.Text = ""
  txt_filter.Text = ""
  proxyAddress = ""
End If

Label1.Caption = Format$(Now, "yyyy-mm-dd")
If txt_path.Text = "" Then txt_path.Text = App.Path & "\"
Dim fso, folder, f, fc

Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FolderExists(txt_path.Text) = False Then
Call fso.createfolder(txt_path.Text)
End If
Set fso = Nothing

Me.Caption = Me.Caption & " 请输入想要下载的贴吧名或者帖子ID"
End Sub
Public Function getHtmlStr(strURL As String, Optional timeout As Long) As String
Dim sTime, ntime
Dim XmlHttp As Object
If strURL = "" Or strURL = vbCrLf Or Len(strURL) < 2 Then Exit Function
Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
XmlHttp.open "GET", strURL, False
On Error GoTo Err_net
sTime = Now '获取当前时间
XmlHttp.send
While XmlHttp.ReadyState <> 4
  DoEvents
  ntime = Now '获取循环时间
  If timeout <> 0 Then
    If DateDiff("s", sTime, ntime) > timeout Then
      getHtmlStr = ""
      Debug.Print "timeout:" & strURL & vbCrLf
      Exit Function '判断超出3秒即超时退出过程
    End If
  End If
Wend
getHtmlStr = BytesToBstr(XmlHttp.responseBody, "UTF-8")
Set XmlHttp = Nothing
Err_net:
End Function

Public Function sendHtmlStr(strURL As String) As String
Dim XmlHttp As MSXML2.XmlHttp
If strURL = "" Or strURL = vbCrLf Then Exit Function
Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
XmlHttp.open "GET", strURL, False
On Error GoTo Err_net
XmlHttp.send
While XmlHttp.ReadyState <> 4
DoEvents
Wend
sendHtmlStr = XmlHttp.StatusText

Set XmlHttp = Nothing
Err_net:
End Function

Public Function getHtmlStr_Async(strURL As String, Optional Sync As Boolean = False, Optional getHtmBody As Boolean = True) As String
'Dim XmlHttp As MSXML2.ServerXmlHttp
If strURL = "" Or strURL = vbCrLf Then Exit Function
Set ServerXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
ServerXmlHttp.open "GET", strURL, Sync
On Error GoTo Err_net
ServerXmlHttp.send
If Sync Then '异步 用timer来实现
Timer1.Enabled = Sync
getHtmlStr_Async = ""
Else '同步,直接获得结果网页
  While ServerXmlHttp.ReadyState <> 4
    DoEvents
  Wend
  If getHtmBody Then
    getHtmlStr_Async = BytesToBstr(ServerXmlHttp.responseBody, "UTF-8")
  Else
    getHtmlStr_Async = ServerXmlHttp.StatusText
  End If
End If
Set ServerXmlHttp = Nothing
Err_net:
End Function

Private Sub Timer1_Timer()
    If ServerXmlHttp.ReadyState = 4 Then
        Timer1.Enabled = False
        If ServerXmlHttp.Status = 200 Then
          urlOK = urlOK + 1
          Label1.Caption = Label1.Caption & "已经成功访问" & urlOK & "个网页 "
        End If
    End If
End Sub


Private Function BytesToBstr(strBody, codeBase) As String
Dim objStream As Object
Set objStream = CreateObject("Adodb.Stream")
objStream.Charset = codeBase
objStream.Type = 1
objStream.mode = 3
objStream.open
objStream.write strBody
objStream.position = 0
objStream.Type = 2

BytesToBstr = objStream.ReadText
objStream.Close
Set objStream = Nothing

End Function

Private Function readutf8(fileName As String, Optional codeBase As String = "UTF-8") As String
'Object.Open(Source,[Mode],[Options],[UserName],[Password])
'Mode 指定打开模式，可不指定，可选参数如下：
 ' adModeRead = 1
  'adModeReadWrite = 3
  'adModeRecursive = 4194304
  'adModeShareDenyNone = 16
  'adModeShareDenyRead = 4
  'adModeShareDenyWrite = 8
  'adModeShareExclusive = 12
  'adModeUnknown = 0
  'adModeWrite = 2
' Options 指定打开的选项，可不指定，可选参数如下：
 ' adOpenStreamAsync = 1
  'adOpenStreamFromRecord = 4
  'adOpenStreamUnspecified = -1
 'UserName 指定用户名，可不指定。
 'Password 指定用户名的密码

Const adTypeText = 2
Dim objStream As Object
Set objStream = CreateObject("Adodb.Stream")
objStream.Charset = codeBase
If Dir(fileName) = "" Then Exit Function
objStream.open
objStream.position = 0
objStream.Type = adTypeText
On Error GoTo line_error
objStream.LoadFromFile fileName
readutf8 = objStream.ReadText

line_error:
objStream.Close
Set objStream = Nothing
End Function

Function writeutf8(filePath As String, str As String, Optional codeBase As String = "UTF-8", Optional compare As Integer = 0)

Const adSaveCreateOverWrite = 2
Const adSaveCreateNotExist = 1
On Error Resume Next
Dim objStream As Object
Dim filefoler As String
Dim fso As Object
Dim newFolder As Object
Dim sfolder As String
Dim oldFile As String
Set fso = CreateObject("Scripting.FileSystemObject")
' 如果目录不存在,就创建该目录
'If fso.FolderExists(filepath) = "" Then fso.createfolder ("filepath")

Set objStream = CreateObject("ADODB.Stream")
objStream.Charset = codeBase
objStream.open
objStream.WriteText str
If compare = 0 Then
  If Dir(filePath) <> "" Then fso.DeleteFile filePath
  objStream.SaveToFile filePath, adSaveCreateOverWrite
ElseIf compare = 1 Then
  oldFile = readutf8(filePath)
  If str <> oldFile And str <> "" Then
    objStream.SaveToFile filePath, adSaveCreateOverWrite
  End If
End If

objStream.Close
Set objStream = Nothing
End Function
    
Sub fileMove(strFilename As String, newFilename As String)
Dim str As String
Dim fso As Object
Dim f As Object
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(strFilename) Then Exit Sub
If fso.FileExists(newFilename) Then
fso.DeleteFile newFilename
End If

'fso.MoveFile strFilename, newFilename
Set fso = Nothing
'修改last modified time 必须使用新建一个file的方式
str = readutf8(strFilename)
Call writeutf8(newFilename, str)
End Sub

Sub fileRename(source As String, target As String)
Dim str As String
str = readutf8(source)
If str <> "" Then
Call writeutf8(target, str)
Call fileDelete(source)
End If
End Sub

Sub fileDelete(strFilename)

Dim fso As Object

Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(strFilename) Then
fso.DeleteFile strFilename
End If

Set fso = Nothing

End Sub
Sub fileCopy(strFilename, newpath)

Dim fso As Object
Dim f As Object
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(strFilename) Then Exit Sub
If fso.FileExists(newpath) Then
fso.DeleteFile newpath
End If
Set f = fso.GetFile(strFilename)
fso.copyFile strFilename, newpath
Set fso = Nothing

End Sub

Private Function Fetch(LinkText As String, s1 As String, s2 As String) As String
On Error GoTo err2
Dim LinkStart As Double '从一个字符串的两个标记中间截取文字
Dim LinkEnd As Double
Dim TempVar As String
If InStr(1, LinkText, s1) > 0 And InStr(1, LinkText, s2) > 0 Then
LinkStart = InStr(1, LinkText, s1)
LinkText = Mid$(LinkText, LinkStart + Len(s1))
LinkEnd = InStr(1, LinkText, s2)
If LinkEnd <= 1 Then GoTo err2
    TempVar = Mid$(LinkText, 1, LinkEnd - 1)
    Fetch = Trim$(TempVar)
    TempVar = ""
    'LinkText = Mid$(LinkText, LinkEnd + Len(s2))
err2:
Else:
    Fetch = ""
End If
End Function


Private Function strSub(sourceText As String, s1 As String, i As Integer) As String
On Error GoTo err2
Dim sourceStart As Integer
Dim sourceEnd As Integer
Dim TempVar As String
If InStr(1, sourceText, s1) > 0 Then
sourceStart = InStr(1, sourceText, s1)
strSub = Mid$(sourceText, sourceStart + Len(s1), i)
err2:   Else: strSub = ""
End If
End Function


Public Sub wait1000(HaoMiao As Long)
Dim t1 As Long
t1 = timeGetTime
While (timeGetTime - t1) < HaoMiao
DoEvents
Wend
End Sub

Function UTF8Encode_ForJs(ByVal szInput As String) As String
       Dim wch  As String
       Dim uch As String
       Dim szRet As String
       Dim x As Long
       Dim inputLen As Long
       Dim nAsc  As Long
       Dim nAsc2 As Long
       Dim nAsc3 As Long
          
       If szInput = "" Then
           UTF8Encode_ForJs = szInput
           Exit Function
       End If
       inputLen = Len(szInput)
       For x = 1 To inputLen
       '得到每个字符
           wch = Mid(szInput, x, 1)
           '得到相应的UNICODE编码
           nAsc = AscW(wch)
       '对于<0的编码　其需要加上65536
           If nAsc < 0 Then nAsc = nAsc + 65536
       '对于<128位的ASCII的编码则无需更改
           If (nAsc And &HFF80) = 0 Then
               szRet = szRet & wch
           Else
               If (nAsc And &HF000) = 0 Then
               '真正的第二层编码范围为000080 - 0007FF
               'Unicode在范围D800-DFFF中不存在任何字符，基本多文种平面中约定了这个范围用于UTF-16扩展标识辅助平面（两个UTF-16表示一个辅助平面字符）.
               '当然，任何编码都是可以被转换到这个范围，但在unicode中他们并不代表任何合法的值。
          
                   uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
                   szRet = szRet & uch
                      
     
     
     
               Else
               '第三层编码00000800 – 0000FFFF
               '首先取其前四位与11100000进行或去处得到UTF-8编码的前8位
               '其次取其前10位与111111进行并运算，这样就能得到其前10中最后6位的真正的编码　再与10000000进行或运算来得到UTF-8编码中间的8位
               '最后将其与111111进行并运算，这样就能得到其最后6位的真正的编码　再与10000000进行或运算来得到UTF-8编码最后8位编码
                   uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
                   Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
                   Hex(nAsc And &H3F Or &H80)
                   szRet = szRet & uch
               End If
           End If
       Next
          
       UTF8Encode_ForJs = szRet
End Function

Sub removeMark(sInput As String)
sInput = Replace(sInput, """", "-")
sInput = Replace(sInput, "'", "")
sInput = Replace(sInput, ":", "-")
sInput = Replace(sInput, "?", "-")
sInput = Replace(sInput, vbCrLf, "")
End Sub


Public Function drill_tieba(txtURL As String, Optional urlNum As Long = 0) As Long
Dim sHtm As String
Dim str1 As String
Dim str2 As String
Dim lastNum As String
Dim pageURL As String
Dim linkAll As String
Dim htmTemp As String
Dim pageTitle As String
Dim sPath As String
Dim shtmPage As String
Dim i As Long
Dim j As Long
Dim k As Long
Dim n As Long
Dim deepDrill As String
Dim id1 As Long
Dim id2 As Long

If txtURL = "" Then Exit Function
sHtm = getHtmlStr(txtURL)
If sHtm = "" Then Exit Function

htmTemp = sHtm 'the first page of this tieba

pageTitle = Fetch(htmTemp, "<title>", "</title>")
pageTitle = Replace(pageTitle, " ", "")
pageTitle = Replace(pageTitle, vbCrLf, "")
pageTitle = Replace(pageTitle, "吧_百度贴吧", "")
pageTitle = Trim(pageTitle)
removeMark (pageTitle)
sPath = txt_path.Text & pageTitle & "\"
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FolderExists(sPath) = "False" Then fso.createfolder (sPath)
If fso.FolderExists(sPath & "p\") = "False" Then fso.createfolder (sPath & "p\")
If fso.FolderExists(sPath & "a\") = "False" Then fso.createfolder (sPath & "a\")
Set fso = Nothing

lastNum = Fetch(htmTemp, txtURL, """")
If lastNum = "" Then n = 0 '只有一页的贴吧不处理~

While lastNum <> ""
  n = CLng(lastNum) / 50
  lastNum = Fetch(htmTemp, txtURL, """")
  lastNum = Trim(lastNum)
Wend

If txt_id1.Text = "" Or txt_id1.Text = "0" Then
    id1 = 1
Else
   id1 = CLng(txt_id1.Text)
End If

If txt_id2.Text = "" Or txt_id2.Text = "0" Then
 id2 = CLng(Trim(str(n + 1)))
Else
 id2 = CLng(txt_id2.Text)
End If
 

For i = 0 To n
  If i >= (id1 - 1) And i <= (id2 - 1) Then
  j = i * 50
  pageURL = txtURL & Trim(str(j))
  If i > 0 Then sHtm = getHtmlStr(pageURL)
  deepDrill = simpleHTM_tieba(sHtm, n, txtURL, pageTitle, i)
  If chk_simple.Value = 1 Then Call shorten("", sHtm, "<div id=""content_leftList""", "txt1.txt", "txt2.txt") '缩水
  Call writeutf8(sPath & "\pn" & Trim(str(i + 1)) & ".htm", sHtm)
  linkAll = linkAll & vbCrLf & pageURL
  End If
Next i

Call fileRename(sPath & "pn1.htm", sPath & "index.htm")

urlNum = n + 1
drill:
For i = 1 To urlNum
deepDrill = readutf8(sPath & "link_pn" & Trim(str(i)) & ".txt")
k = k + FetchDetailAll(deepDrill, , pageTitle)
Call fileMove(sPath & "link_pn" & Trim(str(i)) & ".txt", sPath & "a\link_pn" & Trim(str(i)) & ".txt") '完成了就挪到a
Next i


Call writeutf8(sPath & "a\nav.txt", linkAll)
Call fileCopy("1.css", sPath)
Call fileCopy("2.css", sPath)
Call fileCopy("1.js", sPath)
drill_tieba = k
End Function

Public Function simpleHTM(htmTemp As String, lastPage As Long, txtURL As String, title As String, n As Long, Optional pageID As String = "") As String
Dim i As Integer
Dim j As Integer
Dim str1 As String
Dim str2 As String
Dim pageURL As String
Dim temp As String
Dim id As String
Dim urlAll As String
Dim strURL As String
Dim sPath As String
Dim m As Long
Dim str3 As String
Dim str4 As String

If pageID = "" Then
  htmTemp = Replace(htmTemp, txtURL & "0", "index.htm") '贴吧首页
Else
 htmTemp = Replace(htmTemp, txtURL & "1", pageID & ".htm") '帖子首页
End If

If pageID = "" Then
For i = 1 To lastPage
  If pageID = "" Then
     j = i * 50
    m = i + 1
Else
     j = i
     m = i
End If
  pageURL = txtURL & Trim(str(j))
  If pageID <> "" And m <> 1 Then
  htmTemp = Replace(htmTemp, pageURL, pageID & "pn" & Trim(str(m)) & ".htm")
  End If
Next i
End If

temp = htmTemp
' <a href="/p/4526609445" title=" 变为 <a href="p/4526609445.htm" title="
id = Fetch(temp, "<a href=""/p/", """ title=""")
While id <> ""
  str1 = "<a href=""/p/" & id & """ title="""
  str2 = "<a href=""p/" & id & ".htm"" title="""
  htmTemp = Replace(htmTemp, str1, str2)
  strURL = "http://tieba.baidu.com/p/" & id
  urlAll = urlAll & strURL & vbCrLf
  id = Fetch(temp, "<a href=""/p/", """ title=""")
Wend

If pageID <> "" Then
'sPath = EXEPATH & title & "\p\link_" & pageID & ".txt"
simpleHTM = ""
'详情页<a href="/p/3522726221?pn=2">变为<a href="3522726221pn2.htm">
temp = htmTemp
str1 = "<a href=""/p/" & pageID & "?pn="
str2 = """>"
id = Fetch(temp, str1, str2)
While id <> ""
  str3 = "<a href=""/p/" & pageID & "?pn=" & Trim(id) & """>"
  str4 = "<a href=""" & pageID & "pn" & Trim(id) & ".htm"">"
  htmTemp = Replace(htmTemp, str3, str4)
  id = Fetch(temp, str1, str2)
Wend

  str3 = "<a href=""p/" & pageID & "?pn=1"">"
  str4 = "<a href=""" & pageID & ".htm"">"
  htmTemp = Replace(htmTemp, str3, str4)
  str3 = "<a href=""p/" & pageID & "pn1.htm"">"
  htmTemp = Replace(htmTemp, str3, str4)
  str3 = "<a href=""" & pageID & "pn1.htm"">"
  htmTemp = Replace(htmTemp, str3, str4)
id = 2
  str3 = "<a href=""p/" & pageID & "?pn=" & Trim(id) & """>"
  str4 = "<a href=""" & pageID & "pn" & Trim(id) & ".htm"">"
  htmTemp = Replace(htmTemp, str3, str4)
Else
   sPath = txt_path.Text & title & "\link_pn" & Trim(str(n + 1)) & ".txt"
   Call writeutf8(sPath, urlAll)
  simpleHTM = urlAll
End If

str1 = "href=""/f?kw=" & UTF8Encode_ForJs(title) & "&ie=utf-8&tp=0"""
str2 = "href=""../index.htm"""  '首页
htmTemp = Replace(htmTemp, str1, str2)

str1 = "href=""/f?kw=" & title & "&ie=utf-8&tp=0"""
str2 = "href=""../index.htm"""  '首页
htmTemp = Replace(htmTemp, str1, str2)

str1 = "href=""/f?kw=" & UTF8Encode_ForJs(title) & "&ie=utf-8"""
str2 = "href=""../index.htm"""  '首页
htmTemp = Replace(htmTemp, str1, str2)

str1 = "href=""/f?kw=" & title & "&ie=utf-8"""
str2 = "href=""../index.htm"""  '首页
htmTemp = Replace(htmTemp, str1, str2)

str1 = "</head>"
str2 = "<link rel=""stylesheet"" href=""../2.css"" /><script src=""../1.js""></script></head>"
htmTemp = Replace(htmTemp, str1, str2)

str1 = "href=""/"
str2 = "href=""http://tieba.baidu.com/"
htmTemp = Replace(htmTemp, str1, str2)
End Function

Public Function simpleHTM_tieba(htmTemp As String, lastPage As Long, txtURL As String, title As String, pageNum As Long) As String
Dim i As Integer
Dim j As Integer
Dim str1 As String
Dim str2 As String
Dim pageURL As String
Dim temp As String
Dim id As String
Dim urlAll As String
Dim strURL As String
Dim sPath As String
Dim m As Long
Dim str3 As String
Dim str4 As String
'i->j->m
' 0->0->index.htm  1->50->pn2.htm 2->100->3.htm 154->7700->155.htm

For i = 0 To lastPage
     j = i * 50
    m = i + 1
  pageURL = txtURL & Trim(str(j))
  If i = 0 Then
     htmTemp = Replace(htmTemp, pageURL, "index.htm") '贴吧首页
  Else
     htmTemp = Replace(htmTemp, pageURL, "pn" & Trim(str(m)) & ".htm")
  End If
Next i

temp = htmTemp
' <a href="/p/4526609445" title=" 变为 <a href="p/4526609445.htm" title="
id = Fetch(temp, "<a href=""/p/", """ title=""")
While id <> ""
  str1 = "<a href=""/p/" & id & """ title="""
  str2 = "<a href=""p/" & id & ".htm"" title="""
  htmTemp = Replace(htmTemp, str1, str2)
  strURL = "http://tieba.baidu.com/p/" & id
  urlAll = urlAll & strURL & vbCrLf
  id = Fetch(temp, "<a href=""/p/", """ title=""")
Wend

str1 = "href=""/f?kw=" & UTF8Encode_ForJs(title) & "&ie=utf-8"""
str2 = "href=""index.htm"""  '首页
htmTemp = Replace(htmTemp, str1, str2)

str1 = "href=""/f?kw=" & title & "&ie=utf-8"""
str2 = "href=""index.htm"""  '首页
htmTemp = Replace(htmTemp, str1, str2)

str1 = "</head>"
str2 = "<link rel=""stylesheet"" href=""1.css"" /><script src=""1.js""></script></head>"
htmTemp = Replace(htmTemp, str1, str2)

str1 = "href=""/"
str2 = "href=""http://tieba.baidu.com/"
htmTemp = Replace(htmTemp, str1, str2)

str1 = "'is_login': 0,"
str2 = "'is_login': 1,"
htmTemp = Replace(htmTemp, str1, str2)

str1 = "''is_new_user': 1,"
str2 = "''is_new_user': 0,"
htmTemp = Replace(htmTemp, str1, str2)

sPath = txt_path.Text & title & "\link_pn" & Trim(str(pageNum + 1)) & ".txt"
Call writeutf8(sPath, urlAll)
simpleHTM_tieba = urlAll
End Function

Public Function FetchDetailAll(source As String, Optional filePath As String, Optional title As String) As Long

Dim tURLItem() As String
Dim i As Long
Dim j As Long
Dim k As Long
Dim id As String
Dim pageURL As String

If source = "" And filePath <> "" Then
source = readutf8(filePath)
End If

If source = "" Then Exit Function
tURLItem = Split(source, vbCrLf)

For i = 0 To UBound(tURLItem)
 If tURLItem(i) <> vbCrLf And Trim(tURLItem(i)) <> "" Then
        pageURL = tURLItem(i) & "?pn="
        id = Replace(tURLItem(i), "http://tieba.baidu.com/p/", "")
        k = k + drill_detail(tURLItem(i), j, title)
  End If
Next i

FetchDetailAll = k
End Function


Private Sub cmd_go_Click()
Dim i As Long
Dim j As Long
Dim n As Integer
Dim k As Long
Dim m As Long
Dim sTemp As String
Dim txtNames() As String
Dim txtIDs() As String
Dim sTime As String
wrongURL = ""

'当没有网络的时候停止检查
'Call setProxy(proxyAddress)
sTemp = getHtmlStr("http://www.tieba.com")
If sTemp = "" Then Exit Sub
sTime = Now()
Label1.Caption = sTime & "开始 "
If chk_add.Value = 1 Then
   m = drill_tieba_pn(txt_name.Text)
   sTime = Now()
   Label1.Caption = Label1.Caption & sTime & "结束:获取帖子" & m & "页"
   GoTo done:
End If

If txt_id.Text <> "" Then '批量下载多个帖子
         i = 0
         If InStr(txt_id.Text, ",") = 0 Then '下载单个帖子
            ReDim txtIDs(0)
            txtIDs(0) = txt_id.Text
         Else
            txtIDs = Split(txt_id.Text, ",") '逗号分隔的多个帖子ID
         End If
        For n = 0 To UBound(txtIDs)
            sTemp = "http://tieba.baidu.com/p/" & txtIDs(n)
            j = j + drill_detail(sTemp)
        Next n
        m = n
        sTime = Now()
       Label1.Caption = Label1.Caption & sTime & "结束:获取帖子" & m & "个 包括分页共" & j & "页"
 ElseIf txt_name.Text <> "" Then
         If InStr(txt_name.Text, ",") = 0 Then '下载某一个帖吧全吧的帖子
            ReDim txtNames(0)
            txtNames(0) = txt_name.Text
         Else
            txtNames = Split(txt_name.Text, ",") '逗号分隔的多个贴吧名,批量下载多个贴吧
         End If
         For n = 0 To UBound(txtNames)
            sTemp = "http://tieba.baidu.com/f?kw=" & UCase(UTF8Encode_ForJs(txtNames(n))) & "&ie=utf-8&pn="
            j = j + drill_tieba(sTemp, k)
            i = i + k
        Next n
        m = n
        sTime = Now()
        Label1.Caption = Label1.Caption & sTime & "结束:获取贴吧" & m & "个 导航页" & i & "页  帖子" & j & "页"
End If
done:
If wrongURL <> "" Then
    Call writeutf8("err.txt", wrongURL)
End If
End Sub

Public Function fetchElement(sHtm As String, startStr As String, endStr As String, count As Integer) As Integer
Dim tempHtm As String
Dim pos1 As Long
Dim pos2 As Long
Dim pos3 As Long
tempHtm = sHtm
pos1 = InStr(sHtm, startStr)
If pos1 = 0 Then Exit Function
pos2 = InStr(pos1, sHtm, endStr)
If pos2 = 0 Then Exit Function
If count = 1 Then
 sHtm = Mid(sHtm, pos1, (pos2 + Len(endStr) - pos1))
 Exit Function
End If
 
     pos3 = pos1 + Len(startStr)
     Do
     pos2 = InStr(pos3, tempHtm, endStr)
     If pos2 = 0 Then Exit Do '没搜到?
     pos3 = pos2 + Len(endStr)
     count = count - 1
     Loop Until count = 0
 
 sHtm = Mid(sHtm, pos1, (pos3 - pos1))
End Function

Public Function FetchPair(source As String, startStr As String, Optional endStr As String = "</div>") As String
Dim pairStr As String
Dim pos1 As Long
Dim pos2 As Long
Dim pos3 As Long
Dim count As Integer
Dim tempHtm As String
Dim sHtm As String
tempHtm = source
sHtm = source

pos1 = InStr(sHtm, startStr)
If pos1 = 0 Then Exit Function
pos2 = InStr(pos1, sHtm, endStr)
If pos2 = 0 Then Exit Function
'</div> -> <div
pairStr = Replace(endStr, "/", "")
pairStr = Replace(pairStr, ">", "")

     Do
      pos2 = pos1 + Len(pairStr)
      count = count + 1
      pos1 = InStr(pos2, sHtm, pairStr)
     Loop Until pos1 = 0
 
Call fetchElement(tempHtm, startStr, endStr, count)
FetchPair = tempHtm
End Function

Public Function shorten(sFilename As String, sHtm As String, strStart As String, fileStart As String, fileEnd As String) As String
'网页脱水-导航页
Dim htmTemp As String
Dim htmFile As String
Dim txt1 As String
Dim txt2 As String

If sHtm = "" Then
sHtm = readutf8(sFilename)
End If
htmTemp = FetchPair(sHtm, strStart)
txt1 = readutf8(fileStart)
txt2 = readutf8(fileEnd)
htmFile = txt1 & htmTemp & txt2
If sFilename <> "" And htmTemp <> "" Then
Call writeutf8(sFilename, htmFile)
End If
shorten = htmTemp
End Function

Public Function shorten2(sFilename As String, sHtm As String, strStart As String, fileStart As String, fileEnd As String) As String
'网页脱水-详情页
Dim htmTemp1 As String
Dim htmTemp2 As String
Dim htmTemp3 As String
Dim htmFile As String
Dim txt1 As String
Dim txt2 As String

If sHtm = "" Then
sHtm = readutf8(sFilename)
End If
htmTemp1 = FetchPair(sHtm, "<div class=""p_thread thread_theme_5"" id=""thread_theme_5"">")
htmTemp2 = FetchPair(sHtm, "<div class=""left_section"">")
htmTemp3 = FetchPair(sHtm, "<div class=""p_thread thread_theme_7"" id=""thread_theme_7"">")
txt1 = readutf8(fileStart)
txt2 = readutf8(fileEnd)
txt1 = Replace(txt1, "<!-- A --!>", htmTemp1)
txt2 = Replace(txt1, "<!-- B --!>", htmTemp3)
htmFile = txt1 & htmTemp2 & txt2
If sFilename <> "" And htmTemp2 <> "" Then
Call writeutf8(sFilename, htmFile)
End If
shorten2 = htmTemp2
End Function

Public Function simple_webFile(sPath As String, Optional level1_num As Long = 0, Optional level2_num As Long = 0) As Long

Dim sname As String
Dim temp As String
Dim iOk As Long
Dim iOk_p As Long
Dim fso, folder, f, fc

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(sPath) = False Then Exit Function
    Set folder = fso.GetFolder(sPath)
     If level1_num = 0 Then level1_num = folder.Files.count
    Set fc = folder.Files
    For Each f In fc
        sname = f.Name
        If Right(sname, 4) = ".htm" And iOk < level1_num Then
            temp = shorten(sPath & sname, "", "<div id=""content_leftList""", "txt1.txt", "txt2.txt")
            If temp <> "" Then iOk = iOk + 1
       End If
    Next
'详情页
    Set folder = fso.GetFolder(sPath & "\p\")
    If fso.FolderExists(sPath & "\p\") = False Then Exit Function
    If level2_num = 0 Then level2_num = folder.Files.count
    Set fc = folder.Files
    For Each f In fc
        sname = f.Name
        If Right(sname, 4) = ".htm" And iOk_p < level2_num Then
            temp = shorten2(sPath & "\p\" & sname, "", "", "txt3.txt", "txt4.txt")
            If temp <> "" Then iOk_p = iOk_p + 1
       End If
    Next

Set fc = Nothing
Set f = Nothing
Set folder = Nothing
Set fso = Nothing
simple_webFile = iOk + iOk_p
End Function

Public Function filter(source As String, Optional sHtm As String = "", Optional id As String = "", Optional targetHtm As String) As Boolean
Dim tempHtm As String
Dim keyNames() As String
Dim i As Long
Dim n As Long
filter = False
If source = "" Then Exit Function
If sHtm = "" And id = "" Then Exit Function
If InStr(source, ",") = 0 Then '单个
            ReDim keyNames(0)
            keyNames(0) = source
Else
            keyNames = Split(source, ",") '逗号分隔的多个帖子ID
End If
For n = 0 To UBound(keyNames)
           If sHtm <> "" And InStr(sHtm, keyNames(n)) > 0 Then
              targetHtm = Replace(sHtm, keyNames(n), "")
              filter = True
              Exit Function
           End If
           If id <> "" And keyNames(n) = "#" & id & "#" Then
              filter = True
              Exit Function
           End If
Next n
End Function

Public Function getTitle(sHtm As String, Optional detailTitle As String) As String
Dim htmTemp As String
Dim pageTitle As String
Dim j As Integer
htmTemp = sHtm
pageTitle = Fetch(htmTemp, "<title>", "</title>")
Call removeMark(pageTitle)
pageTitle = Replace(pageTitle, " ", "")
pageTitle = Replace(pageTitle, "吧_百度贴吧", "")
pageTitle = Trim(pageTitle)
    j = InStr(pageTitle, "_")
    If j > 0 Then
     detailTitle = Mid(pageTitle, 1, (j - 1))
     getTitle = Mid(pageTitle, (j + 1))
    Else
     detailTitle = ""
     getTitle = pageTitle
    End If
End Function

Public Function drill_detail(txtURL As String, Optional urlNum As Long = 0, Optional title As String = "") As Long
Dim sHtm As String
Dim lastNum As String
Dim pageURL As String
Dim htmTemp As String
Dim sPath As String
Dim indexHtm As String
Dim i As Long
Dim j As Long
Dim k As Long
Dim n As Long
Dim tempUrl As String
Dim id As String
Dim detailTitle As String
Dim fso As Object
Dim id1 As Long
Dim id2 As Long

If txtURL = "" Then Exit Function
id = Replace(txtURL, "http://tieba.baidu.com/p/", "")
If id = "" Then Exit Function
sHtm = getHtmlStr(txtURL)
If sHtm = "" Then
    wrongURL = wrongURL & txtURL & vbCrLf
    Exit Function
End If
Set fso = CreateObject("Scripting.FileSystemObject")

If id <> "" And title = "" Then '单独下载某个ID
    title = getTitle(sHtm, detailTitle)
    sPath = txt_path.Text & detailTitle & "\"
    If fso.FolderExists(sPath) = "False" Then
        Call fso.createfolder(sPath)
        Call fileCopy("2.css", sPath)
        Call fileCopy("1.js", sPath)
    End If
    sPath = sPath & "\p\"
    If fso.FolderExists(sPath) = "False" Then fso.createfolder (sPath)
    If txt_id1.Text <> "" Then id1 = CLng(txt_id1.Text)
    If txt_id2.Text <> "" Then id2 = CLng(txt_id2.Text)
Else
    sPath = txt_path.Text & title & "\p\"
End If
    

Set fso = Nothing


If filter(txt_filter.Text, detailTitle, id) = True Then Exit Function '过滤
htmTemp = sHtm

If Right(txtURL, 4) <> "?pn=" Then txtURL = txtURL & "?pn="
tempUrl = "/p/" & id & "?pn="

lastNum = Trim(Fetch(htmTemp, tempUrl, """"))
If lastNum = "" Then n = 1 '只有一页的不处理挖掘下一页~

While lastNum <> ""
  n = CInt(lastNum)
  lastNum = Trim(Fetch(htmTemp, tempUrl, """"))
Wend

If id1 = 0 Then id1 = 1
If id2 = 0 Then id2 = n
For i = 1 To n
  If i >= id1 And i <= id2 Then
        pageURL = txtURL & Trim(str(i))
        If i > 1 Then sHtm = getHtmlStr(pageURL)
        If sHtm <> "" Then '感觉此处应该有一个此贴已经被删除的判断
            Call simpleHTM(sHtm, n, txtURL, title, i, id)
            If chk_simple.Value = 1 Then Call shorten2("", sHtm, "", "txt3.txt", "txt4.txt")
            Call writeutf8(sPath & id & "pn" & Trim(str(i)) & ".htm", sHtm)
            If i = 1 Then indexHtm = sHtm
            k = k + 1
        End If
  End If
Next i

urlNum = n
If chk_simple.Value = 1 Then Call shorten2("", indexHtm, "", "txt3.txt", "txt4.txt")
Call writeutf8(sPath & id & ".htm", indexHtm)
fileDelete (sPath & id & "pn1.htm")

drill_detail = k
End Function


Public Function drill_tieba_pn(title As String) As Long

Dim pageTitle As String
Dim sPath As String
Dim i As Long
Dim j As Long
Dim k As Long
Dim deepDrill As String
Dim id1 As Long
Dim id2 As Long

If title = "" Then Exit Function

pageTitle = title
removeMark (pageTitle)
sPath = txt_path.Text & pageTitle & "\"
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FolderExists(sPath) = "False" Then fso.createfolder (sPath)
If fso.FolderExists(sPath & "p\") = "False" Then fso.createfolder (sPath & "p\")
If fso.FolderExists(sPath & "a\") = "False" Then fso.createfolder (sPath & "a\")
Set fso = Nothing

If txt_id1.Text = "" Or txt_id1.Text = "0" Then
Else
   id1 = CLng(txt_id1.Text)
End If

If txt_id2.Text = "" Or txt_id2.Text = "0" Then
Else
 id2 = CLng(txt_id2.Text)
End If
 
If id1 = 0 Or id2 = 0 Then Exit Function

For i = id1 To id2
deepDrill = readutf8(sPath & "link_pn" & Trim(str(i)) & ".txt")
k = k + FetchDetailAll(deepDrill, , pageTitle)
Call fileMove(sPath & "link_pn" & Trim(str(i)) & ".txt", sPath & "a\link_pn" & Trim(str(i)) & ".txt")  '完成了就挪到a
Next i

drill_tieba_pn = k
End Function
