VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QClass_QApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Version 0.1
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000

Public Property Get Name() As String
    Name = QApp_Name
End Property

Public Property Get Author() As String
    Author = QApp_Author
End Property

Public Property Get Author_Website() As String
    Author_Website = QApp_Author_Website
End Property

Public Property Get Version() As String
    Version = QApp_Version
End Property

Public Property Get MajorVersion() As Integer
    MajorVersion = QApp_MajorVersion
End Property

Public Property Get MinorVersion() As Integer
    MinorVersion = QApp_MinorVersion
End Property

Public Property Get ReleaseVersion() As Integer
    ReleaseVersion = QApp_ReleaseVersion
End Property

Public Property Get Comments() As String
    Comments = QApp_Comments
End Property

Public Property Get FileDescription() As String
    FileDescription = QApp_FileDescription
End Property

Public Property Get Website() As String
    Website = QApp_Website
End Property

Public Property Get LegalCopyright() As String
    LegalCopyright = QApp_LegalCopyright
End Property

Public Property Get LegalTrademarks() As String
    LegalTrademarks = QApp_LegalTrademarks
End Property

Public Property Get SubTitle() As String
    SubTitle = QApp_SubTitle
End Property

Public Property Get Title() As String
    Title = QApp_Title
End Property

Public Property Get ThreadID()
    ThreadID = App.ThreadID
End Property

Public Property Get hInstance()
    hInstance = App.hInstance
End Property

Public Property Get Path() As String
    Path = App.Path
End Property

Public Sub ExitQApp()
    QMod_Main.QMsg "exitqapp"
End Sub

Public Sub SetFormTransparency(ByVal limpid As Long)    ' 设置窗体透明度
    Dim tmpForm As Form
    For Each tmpForm In Forms
        Call SetWindowLong(tmpForm.hwnd, GWL_EXSTYLE, GetWindowLong(tmpForm.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
        Call SetLayeredWindowAttributes(tmpForm.hwnd, 0, limpid, LWA_ALPHA)     'limpid在0--255之间
    Next
End Sub


