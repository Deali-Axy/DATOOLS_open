VERSION 5.00
Begin VB.Form QFrm_About 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8430
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   15.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   8430
   StartUpPosition =   2  '屏幕中心
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "代码协会  Copyright @ CodeInstitute 2014"
      ForeColor       =   &H00800000&
      Height          =   540
      Left            =   0
      TabIndex        =   5
      Top             =   4680
      Width           =   8325
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      X1              =   120
      X2              =   8520
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C000&
      Height          =   2655
      Left            =   1080
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   6015
   End
   Begin VB.Label lbl_App_Website 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://sjnt.taobao.com"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3120
      TabIndex        =   4
      Top             =   3840
      Width           =   3405
   End
   Begin VB.Label lbl_Web 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "相关网站："
      Height          =   420
      Left            =   1440
      TabIndex        =   3
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label lbl_App_Version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "程序版本："
      Height          =   420
      Left            =   1440
      TabIndex        =   2
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lbl_App_Author 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "程序作者："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      TabIndex        =   1
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lbl_App_Name 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "程序名称："
      Height          =   420
      Left            =   1440
      TabIndex        =   0
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Image Img_CI 
      Appearance      =   0  'Flat
      Height          =   840
      Left            =   1560
      Picture         =   "QFrm_About.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   7020
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      X1              =   0
      X2              =   8400
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Image Img_AppIcon 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   0
      Picture         =   "QFrm_About.frx":0F6F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "QFrm_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Activate()
    QDB.Log Me.Name & " Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 27
        Unload Me
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo load_err
    mLoadConfig
    Me.Caption = "关于  " & QApp_Title & " (CQAF框架 Ver " & CQAF_Version & ")"
    lbl_App_Name.Caption = lbl_App_Name & QApp_Name
    lbl_App_Author.Caption = lbl_App_Author & QApp_Author
    lbl_App_Version.Caption = lbl_App_Version & QApp_Version
    lbl_App_Website.Caption = QApp_Website
    QDB.Log Me.Name & " Load hWnd=" & Me.hwnd
    Exit Sub
load_err:
    QDB.Runtime_Error Me.Name & "->Load", Err.Description, Err.Number
    Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    QDB.Log Me.Name & " Unload"
End Sub

Private Sub lbl_App_Author_Click()
    ShellExecute Me.hwnd, "open", QApp_Author_Website, "", "", 5
End Sub

Private Sub lbl_App_Website_Click()
    ShellExecute Me.hwnd, "open", lbl_App_Website, "", "", 5
End Sub

Public Sub mLoadConfig()
    On Error GoTo Err
    If UIConfig.useBackPicture Then
        Set Me.Picture = UIConfig.AppBackPicture
    Else
        Set Me.Picture = Nothing
    End If
    Dim c As Control
    For Each c In Me.Controls
        c.foreColor = QApp_ForeColor
        c.backColor = QApp_BackColor
    Next
    Exit Sub
Err:
    QDB.Runtime_Error Me.Name & "->mLoadConfig", Err.Description, Err.Number
    Resume Next
End Sub

