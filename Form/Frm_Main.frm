VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frm_Main 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DATOOLS 可以自定义的自由工具箱！"
   ClientHeight    =   8145
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   13530
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Frm_Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   13530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Btn_AddItem 
      Caption         =   "添加"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   7020
      Width           =   615
   End
   Begin VB.CommandButton Btn_RemoveItem 
      Caption         =   "移除"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   7020
      Width           =   615
   End
   Begin VB.CommandButton Btn_EditItem 
      Caption         =   "编辑"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   7020
      Width           =   615
   End
   Begin VB.CommandButton Btn_RefreshItem 
      Caption         =   "刷新"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      Top             =   7020
      Width           =   615
   End
   Begin VB.CommandButton Btn_RefreshGroup 
      Caption         =   "刷新"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   7020
      Width           =   615
   End
   Begin VB.CommandButton Btn_EditGroup 
      Caption         =   "编辑"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   7020
      Width           =   615
   End
   Begin VB.CommandButton Btn_RemoveGroup 
      Caption         =   "移除"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   7020
      Width           =   615
   End
   Begin VB.CommandButton Btn_AddGroup 
      Caption         =   "添加"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   7020
      Width           =   615
   End
   Begin VB.ListBox Lst_Item 
      BackColor       =   &H00404040&
      ForeColor       =   &H00E0E0E0&
      Height          =   5685
      Left            =   2880
      TabIndex        =   3
      Top             =   1260
      Width           =   3975
   End
   Begin VB.Timer Tmr_ShowTime 
      Interval        =   1
      Left            =   240
      Top             =   6000
   End
   Begin VB.ListBox Lst_Group 
      BackColor       =   &H00404040&
      ForeColor       =   &H00E0E0E0&
      Height          =   5685
      Left            =   120
      TabIndex        =   0
      Top             =   1260
      Width           =   2535
   End
   Begin MSForms.CommandButton Btn_CMD 
      Height          =   615
      Index           =   10
      Left            =   7000
      TabIndex        =   27
      Top             =   6000
      Width           =   1695
      VariousPropertyBits=   19
      Caption         =   "关机"
      Size            =   "2990;1085"
      TakeFocusOnClick=   0   'False
      FontName        =   "微软雅黑"
      FontHeight      =   360
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Btn_CMD 
      Height          =   615
      Index           =   9
      Left            =   7000
      TabIndex        =   26
      Top             =   5400
      Width           =   1695
      VariousPropertyBits=   19
      Caption         =   "关机"
      Size            =   "2990;1085"
      TakeFocusOnClick=   0   'False
      FontName        =   "微软雅黑"
      FontHeight      =   360
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Btn_CMD 
      Height          =   615
      Index           =   8
      Left            =   7000
      TabIndex        =   25
      Top             =   4800
      Width           =   1695
      VariousPropertyBits=   19
      Caption         =   "关机"
      Size            =   "2990;1085"
      TakeFocusOnClick=   0   'False
      FontName        =   "微软雅黑"
      FontHeight      =   360
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Btn_CMD 
      Height          =   615
      Index           =   7
      Left            =   7000
      TabIndex        =   24
      Top             =   4200
      Width           =   1695
      VariousPropertyBits=   19
      Caption         =   "关机"
      Size            =   "2990;1085"
      TakeFocusOnClick=   0   'False
      FontName        =   "微软雅黑"
      FontHeight      =   360
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Btn_CMD 
      Height          =   615
      Index           =   6
      Left            =   7000
      TabIndex        =   23
      Top             =   3600
      Width           =   1695
      VariousPropertyBits=   19
      Caption         =   "关机"
      Size            =   "2990;1085"
      TakeFocusOnClick=   0   'False
      FontName        =   "微软雅黑"
      FontHeight      =   360
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Btn_CMD 
      Height          =   615
      Index           =   5
      Left            =   7000
      TabIndex        =   22
      Top             =   3000
      Width           =   1695
      VariousPropertyBits=   19
      Caption         =   "关机"
      Size            =   "2990;1085"
      TakeFocusOnClick=   0   'False
      FontName        =   "微软雅黑"
      FontHeight      =   360
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Btn_CMD 
      Height          =   615
      Index           =   4
      Left            =   7000
      TabIndex        =   21
      Top             =   2400
      Width           =   1695
      VariousPropertyBits=   19
      Caption         =   "关机"
      Size            =   "2990;1085"
      TakeFocusOnClick=   0   'False
      FontName        =   "微软雅黑"
      FontHeight      =   360
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Btn_CMD 
      Height          =   615
      Index           =   3
      Left            =   7000
      TabIndex        =   20
      Top             =   1800
      Width           =   1695
      VariousPropertyBits=   19
      Caption         =   "关机"
      Size            =   "2990;1085"
      TakeFocusOnClick=   0   'False
      FontName        =   "微软雅黑"
      FontHeight      =   360
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Btn_CMD 
      Height          =   615
      Index           =   2
      Left            =   7000
      TabIndex        =   19
      Top             =   1200
      Width           =   1695
      VariousPropertyBits=   19
      Caption         =   "关机"
      Size            =   "2990;1085"
      TakeFocusOnClick=   0   'False
      FontName        =   "微软雅黑"
      FontHeight      =   360
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Btn_CMD 
      Height          =   615
      Index           =   1
      Left            =   7000
      TabIndex        =   18
      Top             =   600
      Width           =   1695
      VariousPropertyBits=   19
      Caption         =   "重启桌面"
      Size            =   "2990;1085"
      TakeFocusOnClick=   0   'False
      FontName        =   "微软雅黑"
      FontHeight      =   360
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Btn_CMD 
      Height          =   615
      Index           =   0
      Left            =   7000
      TabIndex        =   17
      Top             =   6600
      Width           =   1695
      VariousPropertyBits=   19
      Caption         =   "关机"
      Size            =   "2990;1085"
      TakeFocusOnClick=   0   'False
      FontName        =   "微软雅黑"
      FontHeight      =   360
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox Txt_CMD 
      Height          =   480
      Left            =   120
      TabIndex        =   16
      Top             =   650
      Width           =   4335
      VariousPropertyBits=   746604563
      Size            =   "7646;847"
      SpecialEffect   =   3
      FontName        =   "微软雅黑"
      FontHeight      =   285
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.CommandButton Btn_Config 
      Height          =   615
      Left            =   7000
      TabIndex        =   15
      Top             =   0
      Width           =   1695
      VariousPropertyBits=   19
      Caption         =   "配置"
      Size            =   "2990;1085"
      TakeFocusOnClick=   0   'False
      FontName        =   "微软雅黑"
      FontHeight      =   360
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Btn_Exit 
      Height          =   630
      Left            =   7000
      TabIndex        =   14
      Top             =   7500
      Width           =   1695
      VariousPropertyBits=   19
      Caption         =   "退出"
      Size            =   "2990;1111"
      TakeFocusOnClick=   0   'False
      FontName        =   "微软雅黑"
      FontHeight      =   360
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "DATools快捷方式"
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   4500
      TabIndex        =   13
      Top             =   720
      Width           =   2340
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00E0E0E0&
      X1              =   0
      X2              =   6960
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Line Line_V2 
      BorderColor     =   &H00E0E0E0&
      X1              =   8720
      X2              =   8720
      Y1              =   0
      Y2              =   9000
   End
   Begin VB.Line Line_V1 
      BorderColor     =   &H00E0E0E0&
      X1              =   6960
      X2              =   6960
      Y1              =   0
      Y2              =   9000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      X1              =   2760
      X2              =   2760
      Y1              =   1140
      Y2              =   7440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   60
      X2              =   11820
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   0
      X2              =   11760
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lbl_Greet 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hello,Deali-Axy!"
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2190
   End
   Begin VB.Label lbl_Time 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Width           =   1140
   End
   Begin VB.Label lbl_Trip 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   7740
      Width           =   6975
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim WithEvents cSHK As clsSysHotKey
Attribute cSHK.VB_VarHelpID = -1
Dim fSideBar As Boolean
Dim barCmd(2 To 10) As String
Const MAINFORMKEY = vbKeyF6

Private Sub Btn_AddGroup_Click()    '添加分类
    Dim tmpFrm_EditItem As New Frm_EditItem
    tmpFrm_EditItem.Msg_Proc "addgroup"
    tmpFrm_EditItem.Show
End Sub

Private Sub Btn_AddItem_Click()    '添加项目
    Dim tmpFrm_EditItem As New Frm_EditItem
    tmpFrm_EditItem.Msg_Proc "additem"
    tmpFrm_EditItem.Show
End Sub

Private Sub Btn_Cmd_Click(index As Integer)
    Select Case index
    Case 0
        Shell "shutdown -s -t 1"
    Case 1
        Shell "tskill explorer"
    Case Else
        On Error GoTo Err
        ShellExecute Me.hwnd, "open", barCmd(index), "", "", 5
    End Select
Err:
    QDB.Runtime_Error Me.Name & "->Btn_Cmd_Click", Err.Description, Err.Number
    Resume Next
End Sub

Private Sub Btn_Cmd_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Err
    Dim tmpBarName As String, tmpBarPath As String
    If index > 1 And index < 10 Then
        If Button = 2 Then
            tmpBarName = InputBox("请输入侧边栏按钮名称", "隐藏功能(现处于测试阶段)，请慎用", Btn_CMD(index).Caption)
            tmpBarPath = InputBox("请输入侧边栏按钮命令", "隐藏功能(现处于测试阶段)，请慎用", barCmd(index))
            If Len(tmpBarName) > 0 And Len(tmpBarPath) > 0 Then
                Open App.Path & "\Config\sidebar" & index & ".config" For Output As #33
                Print #33, tmpBarName & "|"; tmpBarPath
                Close #33
                Btn_CMD(index).Caption = tmpBarName
                barCmd(index) = tmpBarPath
                Btn_CMD(index).Refresh
                'Else
            End If
        End If
    End If
    Exit Sub
Err:
    QDB.Runtime_Error Me.Name & "->Btn_Cmd_MouseDown", Err.Description, Err.Number
End Sub

Private Sub Btn_Config_Click()
    Frm_Config.Show
End Sub

Private Sub Btn_EditGroup_Click()    '编辑分类
    If Lst_Group.SelCount Then
        Dim tmpFrm_EditItem As New Frm_EditItem
        tmpFrm_EditItem.Msg_Proc "editgroup"
        tmpFrm_EditItem.Show
    Else
        lbl_Trip = "还没有选择分类呢！"
    End If
End Sub

Private Sub Btn_EditItem_Click()    '编辑项目
    If Lst_Group.SelCount Then
        If Lst_Item.SelCount Then
            Dim tmpFrm_EditItem As New Frm_EditItem
            tmpFrm_EditItem.Msg_Proc "edititem"
            tmpFrm_EditItem.Show
        Else
            lbl_Trip = "还没有选择项目呢！"
        End If
    Else
        lbl_Trip = "还没有选择分类呢！"
    End If
End Sub

Private Sub Btn_Exit_Click()
    fExit = True
    CommandProc "exit"
    Unload Me
End Sub

Private Sub Btn_RefreshGroup_Click()    '刷新分类
    Call mLoadGroup
End Sub

Private Sub Btn_RefreshItem_Click()    '刷新项目
    If Lst_Group.SelCount Then
        Call mLoadItem(GroupIndex)
    Else
        lbl_Trip = "还没有选择分类呢！"
    End If
End Sub

Private Sub Btn_RemoveGroup_Click()    '移除分类
    Debug.Print "[" & Me.Name & ".Btn_RemoveItem.Click]"
    If Lst_Group.SelCount Then
        RemoveGroup GroupIndex    '执行移除
        Btn_RefreshGroup_Click    '刷新分类
    Else
        lbl_Trip = "还没有选择分类呢！"
    End If
End Sub

Private Sub Btn_RemoveItem_Click()    '移除项目
    Debug.Print "[" & Me.Name & ".Btn_RemoveItem.Click]"
    If Lst_Group.SelCount Then
        If Lst_Item.SelCount Then
            RemoveItem GroupIndex, ItemIndex(GroupIndex)
            'Btn_RefreshItem_Click
        Else
            lbl_Trip = "还没有选择项目呢！"
        End If
    Else
        lbl_Trip = "还没有选择分类呢！"
    End If
End Sub

Private Sub cSHK_SysKeyPressed()
    If Me.Visible Then
        Me.Hide
    Else
        Me.Show
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF1
        QFrm_About.Show
    Case vbKeyF5
        LoadConfig
        mLoadConfig
        lbl_Trip.Caption = "刷新配置成功！"
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo Err
    mLoadConfig    '加载配置
    Set cSHK = New clsSysHotKey    '创建SysHookKey对象
    cSHK.SetASysHotKey Me.hwnd, MAINFORMKEY, 0, False    '设置显示主窗口热键
    fSideBar = False
    fExit = False
    Me.Width = Line_V1.X1 + 100
    Exit Sub
Err:
    QDB.Runtime_Error Me.Name & "->Form_Load", Err.Description, Err.Number
    Resume Next
End Sub

Private Sub Form_Terminate()    '程序结束
    cSHK.UnSetSysHotKey    '取消热键
    Set cSHK = Nothing    '释放对象内存
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveConfig    '保存配置
    If fExit Then
    Else
        Me.Hide
        Cancel = 1
    End If
End Sub

Private Sub lbl_Greet_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Dim StrGreetTemp As String
        StrGreetTemp = InputBox("请输入问候内容：", "设置问候语", StrGreet)
        If Len(StrGreetTemp) > 0 Then
            StrGreet = StrGreetTemp
            lbl_Greet = StrGreet
        End If
    Case 2    '右键
        If fSideBar Then
            Me.Width = Line_V1.X1 - 100
            fSideBar = False
        Else
            Me.Width = Line_V2.X1 + 50
            fSideBar = True
        End If
    End Select
End Sub

Private Sub Lst_Group_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "[" & Me.Name & ".Lst_Group.MouseDown]Lst_Group.ListIndex=" & Lst_Group.ListIndex
    Lst_Item.Clear
    lbl_Trip = ""
    If Lst_Group.SelCount Then
        GroupIndex = Lst_Group.ListIndex + 1
        Debug.Print "[" & Me.Name & ".Lst_Group.MouseDown]GroupIndex=" & GroupIndex
        mLoadItem (GroupIndex)
    End If
End Sub

Private Sub Lst_Item_DblClick()
    If Lst_Item.SelCount Then
        Shell Items(GroupIndex).Path(ItemIndex(GroupIndex)), vbNormalFocus
    End If
End Sub

Private Sub Lst_Item_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "[" & Me.Name & ".Lst_Item.MouseDown]Lst_Item.ListIndex=" & Lst_Item.ListIndex
    lbl_Trip = ""
    If Lst_Item.SelCount Then
        ItemIndex(GroupIndex) = Lst_Item.ListIndex + 1
    End If
End Sub

Private Sub Tmr_ShowTime_Timer()    '显示时间
    lbl_Time = Str(Time)
End Sub

Private Sub mLoadGroup()
    Dim i As Integer
    If GroupsCount > 0 Then
        With Lst_Group
            .Clear
            For i = 1 To GroupsCount
                .AddItem Groups(i).Name
            Next
        End With
    End If
End Sub

Private Sub mLoadItem(paramGroupIndex As Integer)
    Dim i As Integer
    If paramGroupIndex > 0 And paramGroupIndex <= GroupsCount Then
        With Lst_Item
            .Clear
            For i = 1 To ItemsCount(paramGroupIndex)
                .AddItem Items(paramGroupIndex).ItemName(i)
            Next
        End With
    End If
End Sub

Public Sub mLoadConfig()
    mLoadGroup
    On Error GoTo Err
    If Len(StrGreet) > 0 Then
        lbl_Greet = StrGreet
    End If
    Dim c As Control
    Me.backColor = QApp_BackColor
    For Each c In Me.Controls
        c.backColor = QApp_BackColor
        c.foreColor = QApp_ForeColor
        c.BorderColor = QApp_ForeColor
    Next

    Dim tmpi As Integer
    Dim tmpName As String, tmpBarPath As String, tmpLine As String
    For tmpi = 2 To 10
        If Len(Dir(App.Path & "\Config\sidebar" & tmpi & ".config")) > 0 Then
            Open App.Path & "\Config\sidebar" & tmpi & ".config" For Input As #35
            Line Input #35, tmpLine
            If Len(tmpLine) > 0 Then
                tmpName = Mid(tmpLine, 1, InStr(1, tmpLine, "|") - 1)
                tmpBarPath = Mid(tmpLine, InStr(1, tmpLine, "|") + 1)
                Btn_CMD(tmpi).Caption = tmpName
                barCmd(tmpi) = tmpBarPath
            End If
            Close #35
        End If
    Next


    'If UIConfig.Border Then
    '    For Each c In Me.Controls
    '        c.BorderStyle = 0
    '    Next
    'Else
    '    For Each c In Me.Controls
    '        c.BorderStyle = 0
    '    Next
    'End If

    'If UIConfig.Clock Then
    '    Me.Tmr_ShowTime.Enabled = True
    '    lbl_Time.Visible = True
    'Else
    '    Me.Tmr_ShowTime.Enabled = False
    '    lbl_Time.Visible = False
    'End If

    'If UIConfig.Greet Then
    '    lbl_Greet.Caption = StrGreet
    'Else
    '    lbl_Greet.Caption = QApp.SubTitle
    'End If
    '
    '    If UIConfig.QUISkin Then
    '        Mod_QHookSkin.Attach Me.hwnd
    '    Else
    '        Mod_QHookSkin.Detach Me.hwnd
    '    End If
    '

    If UIConfig.useBackPicture Then    '设置背景图片
        Set Me.Picture = UIConfig.AppBackPicture
    Else
        Set Me.Picture = Nothing
    End If

    QApp.SetFormTransparency AppTransparency
    Exit Sub
Err:
    QDB.Runtime_Error Me.Name & "->mLoadConfig", Err.Description, Err.Number
    Resume Next
End Sub

Public Sub Msg_Proc(Msg As String, Optional Source As String)
    Select Case Msg
    Case "refresh"
        LoadConfig
        mLoadConfig
    End Select
End Sub

Private Sub Txt_CMD_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode.Value
    Case 13
        Dim RetVal As Boolean
        If Len(Txt_CMD) > 0 Then RetVal = CommandProc(Txt_CMD)
        If IsChinese(Txt_CMD) Then    '输入中文搜索
            ShellExecute Me.hwnd, "open", "http://www.baidu.com/s?wd=" & Txt_CMD, "", "", 5
            Txt_CMD = ""
            lbl_Trip = "命令执行完成 ^ ^"
        End If
        If RetVal Then
            Txt_CMD = ""
            lbl_Trip = "命令执行完成 ^ ^"
        Else
            lbl_Trip = "命令执行失败 T T.."
        End If
    End Select
End Sub
