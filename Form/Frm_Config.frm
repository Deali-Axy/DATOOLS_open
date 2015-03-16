VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frm_Config 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "³ÌÐòÅäÖÃ"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7440
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_Config.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   7440
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   6840
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbl_UseBackPic 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "±³¾°Í¼Æ¬(ÒÑÆôÓÃ)"
      Height          =   315
      Left            =   360
      TabIndex        =   12
      Top             =   2160
      Width           =   1830
   End
   Begin MSForms.CheckBox Check_UseQUISkin 
      Height          =   465
      Left            =   4920
      TabIndex        =   11
      Top             =   1440
      Width           =   1815
      VariousPropertyBits=   1015023633
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "3201;820"
      Value           =   "0"
      Caption         =   "ÆôÓÃQUIÆ¤·ô"
      FontName        =   "Î¢ÈíÑÅºÚ"
      FontEffects     =   1073750016
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.CheckBox Check_ShowGreet 
      Height          =   465
      Left            =   3240
      TabIndex        =   10
      Top             =   1440
      Width           =   1605
      VariousPropertyBits=   1015023633
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2831;820"
      Value           =   "0"
      Caption         =   "ÏÔÊ¾ÎÊºòÓï"
      SpecialEffect   =   0
      FontName        =   "Î¢ÈíÑÅºÚ"
      FontEffects     =   1073750016
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.CheckBox Check_ShowTime 
      Height          =   465
      Left            =   1800
      TabIndex        =   9
      Top             =   1440
      Width           =   1365
      VariousPropertyBits=   1015023633
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2408;820"
      Value           =   "0"
      Caption         =   "ÏÔÊ¾Ê±¼ä"
      SpecialEffect   =   0
      FontName        =   "Î¢ÈíÑÅºÚ"
      FontEffects     =   1073750016
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.CheckBox Check_ShowBorder 
      Height          =   465
      Left            =   360
      TabIndex        =   8
      Top             =   1440
      Width           =   1365
      VariousPropertyBits=   1015023633
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "2408;820"
      Value           =   "0"
      Caption         =   "ÏÔÊ¾±ß¿ò"
      SpecialEffect   =   0
      FontName        =   "Î¢ÈíÑÅºÚ"
      FontEffects     =   1073750016
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.CommandButton Btn_BackupConfig 
      Height          =   1020
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   2655
      VariousPropertyBits=   268435475
      Caption         =   "±¸·ÝÅäÖÃ"
      Size            =   "4683;1799"
      TakeFocusOnClick=   0   'False
      FontName        =   "Î¢ÈíÑÅºÚ"
      FontHeight      =   600
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Btn_Transparency_OK 
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   2760
      Width           =   1095
      VariousPropertyBits=   19
      Caption         =   "È·¶¨"
      Size            =   "1931;873"
      TakeFocusOnClick=   0   'False
      FontName        =   "Î¢ÈíÑÅºÚ"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox Txt_Transparency 
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      ToolTipText     =   "ÊýÖµ·¶Î§£º0-255 ¡¾0ÎªÈ«Í¸Ã÷£¬²»½¨ÒéÉèÖÃ£¬255Îª²»Í¸Ã÷¡¿"
      Top             =   2760
      Width           =   1335
      VariousPropertyBits=   746604563
      Size            =   "2355;873"
      SpecialEffect   =   3
      FontName        =   "Î¢ÈíÑÅºÚ"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Txt_BackPic 
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   2040
      Width           =   3615
      VariousPropertyBits=   746604563
      Size            =   "6376;873"
      FontName        =   "Î¢ÈíÑÅºÚ"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.CommandButton Btn_UseBackPic_Open 
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   2040
      Width           =   735
      VariousPropertyBits=   19
      Caption         =   "ä¯ÀÀ"
      Size            =   "1296;873"
      TakeFocusOnClick=   0   'False
      FontName        =   "Î¢ÈíÑÅºÚ"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Btn_Custom 
      Height          =   555
      Left            =   5520
      TabIndex        =   2
      Top             =   720
      Width           =   1455
      VariousPropertyBits=   268435475
      Caption         =   "×Ô¶¨ÒåÑÕÉ«"
      Size            =   "2566;979"
      TakeFocusOnClick=   0   'False
      FontName        =   "Î¢ÈíÑÅºÚ"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin VB.Image Image_Theme 
      Height          =   600
      Index           =   3
      Left            =   2520
      Picture         =   "Frm_Config.frx":164A
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image_Theme 
      Height          =   600
      Index           =   0
      Left            =   360
      Picture         =   "Frm_Config.frx":1E17
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image_Theme 
      Height          =   600
      Index           =   1
      Left            =   1080
      Picture         =   "Frm_Config.frx":25DF
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image_Theme 
      Height          =   600
      Index           =   2
      Left            =   1800
      Picture         =   "Frm_Config.frx":2DA7
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image_Theme 
      Height          =   600
      Index           =   4
      Left            =   3240
      Picture         =   "Frm_Config.frx":3574
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image_Theme 
      Height          =   600
      Index           =   5
      Left            =   3960
      Picture         =   "Frm_Config.frx":3D41
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image_Theme 
      Height          =   600
      Index           =   6
      Left            =   4680
      Picture         =   "Frm_Config.frx":450E
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÉèÖÃÍ¸Ã÷¶È"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2880
      Width           =   1200
   End
   Begin MSForms.Label Label2 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1695
      VariousPropertyBits=   276824083
      Caption         =   "Ö÷Ìâ¡¢½çÃæÅäÖÃ"
      Size            =   "2990;556"
      FontName        =   "Î¢ÈíÑÅºÚ"
      FontHeight      =   240
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin VB.Shape Shape_Theme 
      BorderWidth     =   2
      Height          =   3375
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "Frm_Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const Color_Black As Long = &H0&
Private Const Color_Black2 As Long = &H7E7E7E
Private Const Color_Blue As Long = &HE8A300
Private Const Color_Blue2 As Long = &HEAD99A
Private Const Color_Green As Long = &H1CE5B5
Private Const Color_Yellow As Long = &HCC9FF
Private Const Color_Yellow2 As Long = &HAFE3EF

Private Type ThemeStruct
    backColor As Long
    foreColor As Long
End Type

Dim Themes() As ThemeStruct

Private Sub Btn_BackupConfig_Click()
    Dim tmpName As String, tmpDate As String
    On Error GoTo Err
    tmpDate = Trim(Str(Year(Now))) & "-" & Trim(Str(Month(Now))) & "-" & Trim(Str(Day(Now))) & ".ConfigPkg"
    tmpName = InputBox("ÇëÊäÈëÅäÖÃÃû³Æ", "±¸·ÝÅäÖÃÎÄ¼þ", tmpDate, Me.Left, Me.Top)
    If Len(tmpName) > 0 Then
        ConfigBackup tmpName
    End If
    Exit Sub
Err:
    QDB.Runtime_Error Me.Name & "_Btn_BackupConfig_Click", Err.Description, Err.Number
    Resume Next
End Sub

Private Sub Btn_Custom_Click()
    Dim tmpBackColor As Long, tmpForeColor As Long
    With CommonDialog
        .CancelError = True
        On Error GoTo CancelError
        MsgBox "ÇëÑ¡Ôñ±³¾°É«"
        .DialogTitle = "ÇëÑ¡Ôñ±³¾°É«"
        .ShowColor
        tmpBackColor = .Color
        MsgBox "ÇëÑ¡ÔñÇ°¾°É«"
        .DialogTitle = "ÇëÑ¡ÔñÇ°¾°É«"
        .ShowColor
        tmpForeColor = .Color
    End With

    QApp_BackColor = tmpBackColor
    QApp_ForeColor = tmpForeColor
    SaveConfig
    mLoadConfig
    Frm_Main.mLoadConfig
    Exit Sub
CancelError:

End Sub

Private Sub Btn_Transparency_OK_Click()
    Dim i As Long
    i = Val(Txt_Transparency.Text)
    If i >= 0 And i <= 255 Then
        QApp.SetFormTransparency i
        AppTransparency = i
        SaveConfig
    Else
        MsgBox "ÊýÖµ´íÎó£¡"
    End If
End Sub

Private Sub Btn_UseBackPic_Open_Click()
    Dim tmpStrBackPic As String
    With CommonDialog
        .CancelError = True
        On Error GoTo CancelError
        .Filter = "JPEG|*.jpg|GIF|*.gif|BMP|*.bmp"
        .ShowOpen
        tmpStrBackPic = .FileName
    End With
    Txt_BackPic.Text = tmpStrBackPic
    UIConfig.AppBackPicturePath = tmpStrBackPic
    Set UIConfig.AppBackPicture = LoadPicture(tmpStrBackPic)
    SaveConfig
    mLoadConfig
    Frm_Main.mLoadConfig
    Exit Sub
CancelError:
End Sub

Private Sub Check_ShowBorder_Click()
    If UIConfig.Border Then
        UIConfig.Border = False
    Else
        UIConfig.Border = True
    End If
    SaveConfig
    mLoadConfig
    Frm_Main.mLoadConfig
End Sub

Private Sub Check_ShowGreet_Click()
    If UIConfig.Greet Then
        UIConfig.Greet = False
    Else
        UIConfig.Greet = True
    End If
    SaveConfig
    mLoadConfig
    Frm_Main.mLoadConfig
End Sub

Private Sub Check_ShowTime_Click()
    If UIConfig.Clock Then
        UIConfig.Clock = False
    Else
        UIConfig.Clock = True
    End If
    SaveConfig
    mLoadConfig
    Frm_Main.mLoadConfig
End Sub

Private Sub Check_UseQUISkin_Click()
    If UIConfig.QUISkin Then
        UIConfig.QUISkin = False
    Else
        UIConfig.QUISkin = True
    End If
    SaveConfig
    mLoadConfig
    Frm_Main.mLoadConfig
End Sub

Private Sub Form_Load()
    On Error GoTo Err
    mLoadConfig
    ReDim Themes(6)
    Themes(0).backColor = Color_Black
    Themes(0).foreColor = vbWhite
    Themes(1).backColor = Color_Black2
    Themes(1).foreColor = Color_Yellow2
    Themes(2).backColor = Color_Blue
    Themes(2).foreColor = vbWhite
    Themes(3).backColor = Color_Blue2
    Themes(3).foreColor = Color_Black
    Themes(4).backColor = Color_Green
    Themes(4).foreColor = Color_Blue
    Themes(5).backColor = Color_Yellow
    Themes(5).foreColor = Color_Black2
    Themes(6).backColor = Color_Yellow2
    Themes(6).foreColor = Color_Blue



    'Dim tmpImg As Image
    'For Each tmpImg In Me.Image_Theme
    '    tmpImg.Appearance = 0
    '    tmpImg.BorderStyle = 1
    'Next
    Exit Sub
Err:
    QDB.Runtime_Error Me.Name & "_Form_Load", Err.Description, Err.Number
    Resume Next
End Sub

Public Sub mLoadConfig()
    On Error GoTo Err
    '    Dim c As Control
    '    For Each c In Me.Controls
    '        c.backColor = QApp_BackColor
    '        c.foreColor = QApp_ForeColor
    '    Next
    Me.backColor = QApp_BackColor
    Shape_Theme.BorderColor = QApp_ForeColor
    Me.Check_ShowBorder.foreColor = QApp_ForeColor
    Me.Check_ShowGreet.foreColor = QApp_ForeColor
    Me.Check_ShowTime.foreColor = QApp_ForeColor
    Me.Check_UseQUISkin.foreColor = QApp_ForeColor
    Me.Txt_BackPic.foreColor = QApp_ForeColor
    Me.Txt_Transparency.foreColor = QApp_ForeColor
    Me.Btn_BackupConfig.foreColor = QApp_ForeColor
    Me.Btn_Custom.foreColor = QApp_ForeColor
    Me.Btn_Transparency_OK.foreColor = QApp_ForeColor
    Me.Btn_UseBackPic_Open.foreColor = QApp_ForeColor
    Me.Label1.foreColor = QApp_ForeColor
    Me.Label2.foreColor = QApp_ForeColor
    lbl_UseBackPic.foreColor = QApp_ForeColor


    'If UIConfig.Border Then
    '    Me.Check_ShowBorder.Value = 1
    '    For Each c In Me.Controls
    '        c.BorderStyle = 1
    '    Next
    'Else
    '    Me.Check_ShowBorder.Value = 0
    '    For Each c In Me.Controls
    '        c.BorderStyle = 0
    '    Next
    'End If

    'If UIConfig.Clock Then
    '    Me.Check_ShowTime.Value = True
    'Else
    '    Me.Check_ShowTime.Value = False
    'End If

    'If UIConfig.Greet Then
    '    Me.Check_ShowGreet.Value = True
    'Else
    '    Me.Check_ShowGreet.Value = False
    'End If

    'If UIConfig.QUISkin Then
    '    Me.Check_UseQUISkin.Value = 1
    '    Mod_QHookSkin.Attach Me.hwnd
    'Else
    '    Me.Check_UseQUISkin.Value = 0
    '    Mod_QHookSkin.Detach Me.hwnd
    'End If
    Me.Txt_Transparency.Text = Str(AppTransparency)

    Me.Txt_BackPic.Text = UIConfig.AppBackPicturePath

    If UIConfig.useBackPicture Then
        Set Me.Picture = UIConfig.AppBackPicture
        lbl_UseBackPic.Caption = "±³¾°Í¼Æ¬(ÒÑÆôÓÃ)"
    Else
        Set Me.Picture = Nothing
        lbl_UseBackPic.Caption = "±³¾°Í¼Æ¬(Î´ÆôÓÃ)"
    End If

    Exit Sub
Err:
    QDB.Runtime_Error Me.Name & "_mLoadConfig", Err.Description, Err.Number
    Resume Next
End Sub

Private Sub Image_Theme_Click(index As Integer)
    QApp_BackColor = Themes(index).backColor
    QApp_ForeColor = Themes(index).foreColor
    SaveConfig
    mLoadConfig
    Frm_Main.mLoadConfig
End Sub

Private Sub lbl_UseBackPic_Click()
On Error GoTo Err
    If Len(Txt_BackPic.Text) > 0 Then
        If UIConfig.useBackPicture Then
            UIConfig.useBackPicture = False
        Else
            UIConfig.useBackPicture = True
        End If
        SaveConfig
        mLoadConfig
        Frm_Main.mLoadConfig
    Else
        MsgBox "ÇëÏÈÉèÖÃ±³¾°Í¼Æ¬"
        lbl_UseBackPic.Caption = "±³¾°Í¼Æ¬(Î´ÆôÓÃ)"
    End If
    Exit Sub
Err:
    QDB.Runtime_Error Me.Name & "->lbl_UseBackPic_Click", Err.Description, Err.Number
    Resume Next
End Sub
