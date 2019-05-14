VERSION 5.00
Begin VB.UserControl uXMLHTTP 
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3180
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   3180
   Begin VB.Timer tmrReader 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   960
      Top             =   120
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   7  'Diagonal Cross
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   420
   End
End
Attribute VB_Name = "uXMLHTTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Inet  As Object
Private m_Sync  As Boolean
Public Event OnDateArrived()
Public Event OnReadyStateChange()

Private Sub tmrReader_Timer()
    RaiseEvent OnReadyStateChange

    If ReadyState = 4 Then
        If Status = 200 Then
            tmrReader.Enabled = False
            RaiseEvent OnDateArrived
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    Set m_Inet = CreateObject("MSXML2.ServerXMLHTTP")
End Sub

Private Sub UserControl_Terminate()
    tmrReader.Enabled = False
    Set m_Inet = Nothing
End Sub

Private Sub UserControl_Resize()
    UserControl.Size 420, 420
End Sub

Public Sub Open_(AMethod As String, AUrl As String, Optional Sync As Boolean = False)
    m_Inet.open UCase(AMethod), AUrl, Sync
    m_Sync = Sync
End Sub

Public Sub SetRequestHeader(AHeader As String, AValue As String)
    m_Inet.SetRequestHeader AHeader, AValue
End Sub

Public Sub Send_(Optional ABody As String)
    m_Inet.Send ABody
    tmrReader.Enabled = m_Sync
End Sub

Public Sub Abort_()
    tmrReader.Enabled = False
    m_Inet.Abort
End Sub

Public Property Get ReadyState() As Long
    ReadyState = m_Inet.ReadyState
End Property

Public Property Get Status() As Long
    On Error Resume Next
    Status = m_Inet.Status
End Property

Public Property Get StatusText() As String
    On Error Resume Next
    StatusText = m_Inet.StatusText
End Property

Public Property Get GetResponseHeader(AHeader As String) As String
    GetResponseHeader = m_Inet.GetResponseHeader(AHeader)
End Property

Public Property Get GetAllResponseHeaders() As String
    GetAllResponseHeaders = m_Inet.GetAllResponseHeaders
End Property

Public Property Get ResponseBody() As Byte()
    ResponseBody = m_Inet.ResponseBody
End Property

Public Property Get ResponseText() As String
    ResponseText = m_Inet.ResponseText
End Property

