Attribute VB_Name = "Mod_Main"
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'----------------------------QApp Public Object-------------------------------
Public QDB As New QClass_QDebug    '全局的QDebug调试器
Public QApp As New QClass_QApp
'----------------------------QApp Config-------------------------------
Public Const QApp_Name As String = "DATOOLS"    '程序名称
Public Const QApp_Author As String = "Deali-Axy"    '程序作者
Public Const QApp_Author_Website As String = "http://weibo.com/dealiaxy"    '作者网站
Public Const QApp_Version As String = "1.5.32 Beta 62"    '程序版本(字符串类型)
Public Const QApp_MajorVersion As Integer = 1    '程序主版本
Public Const QApp_MinorVersion As Integer = 5    '程序次版本
Public Const QApp_ReleaseVersion As Integer = 32    '程序修正版本
Public Const QApp_Comments As String = ""    '程序注释
Public Const QApp_FileDescription As String = ""    '文件说明
Public Const QApp_Website As String = "http://weibo.com/dealiaxy"    '相关网页
Public Const QApp_LegalCopyright As String = "Copyright @ Deali-Axy"    '法律版权
Public Const QApp_LegalTrademarks As String = "Deali-Axy"    '法律商标
Public Const QApp_SubTitle = "DATOOLS " & QApp_Version    '程序子标题
Public Const QApp_Title = "DATOOLS"    '程序主标题
'----------------------------CQAF Config-------------------------------
Public Const CQAF_Version = "0.2.1"    'CQAF版本
'----------------------------QApp Standard Error Config-------------------------------
Public Const ErrNum_SubMain = 1
Public Const ErrNum_FormLoad = 2
Public Const ErrNum_Form = 3
Public Const ErrNum_Other = 1024
'----------------------------QApp Pretreatment-------------------------------
#Const App_Load_Interface = False    'QFrm_Load 配置
#Const MLC_HookSkin = False    '使用QHookSkin皮肤配置

Private Type QApp_Info
    App_Name As String
    App_Authuor As String
    App_Version As String
    App_MajorVersion As Integer
    App_MinorVersion As Integer
    App_ReleaseVersion As Integer
    App_Comments As String
    App_FileDescription As String
    App_Website As String
    App_LegalCopyright As String
    App_LegalTrademarks As String
End Type

Public QApp_BackColor As Long      '程序背景色
Public QApp_ForeColor As Long    '程序前景色

Const MaxGroupsCount As Integer = 256
Const MaxItemsCount As Integer = 512

Public Type GroupConfig
    Name As String
    Path As String
    Description As String
End Type

Public Type ItemConfig
    ItemName(1 To MaxItemsCount) As String
    Path(1 To MaxItemsCount) As String
    Description(1 To MaxItemsCount) As String
End Type

Public Type UIConfigStruct
    Border As Boolean
    Clock As Boolean
    Greet As Boolean
    QUISkin As Boolean
    useBackPicture As Boolean
    AppBackPicture As StdPicture    '背景图片
    AppBackPicturePath As String
End Type

Public MainINI As String    '主配置文件路径
Public GroupsCount As Integer    '分类数
Public ItemsCount(1 To MaxGroupsCount) As Integer    '项目数
Public ItemIndex(1 To MaxGroupsCount) As Integer    '项目索引
Public GroupIndex As Integer  '分类索引
Public StrGreet As String  '问候语（软件标题）
Public AppTransparency As Long  '软件透明度
Public fExit As Boolean
Public Groups(1 To MaxGroupsCount) As GroupConfig, Items(1 To MaxGroupsCount) As ItemConfig
Public UIConfig As UIConfigStruct


Sub Main()
    On Error GoTo Err
    Select Case Trim(Command)
    Case "boot"
        '加载配置
        LoadConfig
        '加载主窗口
        Load Frm_Main
        Mod_QHookSkin.Attach Frm_Main.hwnd
    Case ""
        '加载配置
        LoadConfig
        '加载主窗口
        Load Frm_Main
        Mod_QHookSkin.Attach Frm_Main.hwnd
        Frm_Main.Show
    Case "shutdown"
        Shell "shutdown -s -t 0"
    Case "reboot"
        Shell "shutdown -r -t 0"
    Case Else
        Exit Sub
    End Select
    Exit Sub
Err:
    QDB.Runtime_Error "Mod_Main->Sub Main", Err.Description, Err.Number
    Resume Next
End Sub

Public Function LoadConfig() As String
    On Error GoTo Err
    MainINI = App.Path & "\Config\main.config"
    If Dir(MainINI) = "" Then
        GroupsCount = 0
        '没有main.config则创建
        Open MainINI For Output As #1
        Close #1
        Mod_QINI.WriteText MainINI, "DATOOLS", "Config.Count", "0"
    Else
        GroupsCount = Val(Mod_QINI.GetText(MainINI, "DATOOLS", "Config.Count"))    '读取分类配置文件数量
        If GroupsCount > 0 Then
            '在有分类配置文件的情况下读取配置
            LoadGroup
            Dim i As Integer, k As Integer, strItemsCount As String, Path As String
            For i = 1 To GroupsCount
                Path = Mod_QINI.GetText(MainINI, "Config." & i, "Config.Path")
                If Dir(Path) = "" Then
                    Path = App.Path & "\Config\group." & i & ".config"
                    Mod_QINI.WriteText MainINI, "Config." & i, "Config.Path", Path
                    Open Path For Output As #1
                    Close #1
                    Mod_QINI.WriteText MainINI, "Config." & i, "Config.Item.Count", "0"
                    Mod_QINI.WriteText Path, "DATOOLS", "Config.Item.Count", "0"
                End If
                Path = Trim(Mod_QINI.GetText(MainINI, "Config." & i, "Config.Path"))
                ItemsCount(i) = Val(Trim(Mod_QINI.GetText(MainINI, "Config." & i, "Config.Item.Count")))
                strItemsCount = Trim(Mod_QINI.GetText(MainINI, "Config." & i, "Config.Item.Count"))
                If strItemsCount = "" Then
                    Debug.Print "[Mod_Main.LoadConfig]strItemsCount(" & i & ")=“”"
                    Debug.Print "[Mod_Main.LoadConfig]ItemsCount(" & i & ")=0"
                    ItemsCount(i) = 0
                End If
                If ItemsCount(i) > 0 Then
                    Debug.Print "[Mod_Main.LoadConfig]ItemsCount(" & i & ")>0"
                    LoadItem (i)
                End If
            Next
            GroupIndex = 0
            For i = 1 To GroupsCount
                ItemIndex(i) = 0
            Next
        End If
    End If
    Debug.Print "[Mod_Main.LoadConfig]GroupsCount=" & GroupsCount
    For i = 1 To GroupsCount
        Debug.Print "[Mod_Main.LoadConfig]ItemsCount(" & i & ")=" & ItemsCount(i)
    Next
    StrGreet = Mod_QINI.GetText(MainINI, "DATOOLS", "Config.Greet")    '读取问候语
    QApp_BackColor = Val(Trim(Mod_QINI.GetText(MainINI, "UI", "UI.Config.BackColor")))    '读取颜色配置
    QApp_ForeColor = Val(Trim(Mod_QINI.GetText(MainINI, "UI", "UI.Config.ForeColor")))

    If Len(Trim(Mod_QINI.GetText(MainINI, "UI", "UI.Config.Transparency"))) > 0 Then
        AppTransparency = Val(Trim(Mod_QINI.GetText(MainINI, "UI", "UI.Config.Transparency")))    '读取透明度
    Else
        AppTransparency = 200
    End If
    'QApp.SetFormTransparency AppTransparency
    '读取界面配置
    If Trim(Mod_QINI.GetText(MainINI, "UI", "UI.Config.Border")) = "True" Then
        UIConfig.Border = True
    Else
        UIConfig.Border = False
    End If

    If Trim(Mod_QINI.GetText(MainINI, "UI", "UI.Config.Clock")) = "True" Then
        UIConfig.Clock = True
    Else
        UIConfig.Clock = False
    End If

    If Trim(Mod_QINI.GetText(MainINI, "UI", "UI.Config.Greet")) = "True" Then
        UIConfig.Greet = True
    Else
        UIConfig.Greet = False
    End If

    If Trim(Mod_QINI.GetText(MainINI, "UI", "UI.Config.QUISkin")) = "True" Then
        UIConfig.QUISkin = True
    Else
        UIConfig.QUISkin = False
    End If

    If Trim(Mod_QINI.GetText(MainINI, "UI", "UI.Config.useBackPicture")) = "True" Then
        UIConfig.useBackPicture = True
    Else
        UIConfig.useBackPicture = False
    End If
    '读取背景图片设置
    UIConfig.AppBackPicturePath = Mod_QINI.GetText(MainINI, "UI", "UI.Config.AppBackPicturePath")
    If Len(UIConfig.AppBackPicturePath) > 0 Then
        If Len(Dir(UIConfig.AppBackPicturePath)) > 0 Then
            Set UIConfig.AppBackPicture = LoadPicture(UIConfig.AppBackPicturePath)
        Else
            Mod_QINI.WriteText MainINI, "UI", "UI.Config.useBackPicture", "False"
            Mod_QINI.WriteText MainINI, "UI", "UI.Config.AppBackPicturePath", ""
        End If
    End If

    Exit Function
Err:
    QDB.Runtime_Error "Mod_Main->LoadConfig", Err.Description, Err.Number
    Resume Next
End Function

Public Function SaveConfig() As String
    On Error GoTo Err
    SaveGroup
    Dim i As Integer
    For i = 1 To GroupsCount
        SaveItem i
    Next
    Mod_QINI.WriteText MainINI, "DATOOLS", "Config.Greet", StrGreet
    Mod_QINI.WriteText MainINI, "UI", "UI.Config.BackColor", Trim(Str(QApp_BackColor))    '保存颜色配置
    Mod_QINI.WriteText MainINI, "UI", "UI.Config.ForeColor", Trim(Str(QApp_ForeColor))
    Mod_QINI.WriteText MainINI, "UI", "UI.Config.Transparency", Trim(Str(AppTransparency))    '保存透明度
    '保存界面配置
    If UIConfig.Border Then
        Mod_QINI.WriteText MainINI, "UI", "UI.Config.Border", "True"
    Else
        Mod_QINI.WriteText MainINI, "UI", "UI.Config.Border", "False"
    End If

    If UIConfig.Clock Then
        Mod_QINI.WriteText MainINI, "UI", "UI.Config.Clock", "True"
    Else
        Mod_QINI.WriteText MainINI, "UI", "UI.Config.Clock", "False"
    End If

    If UIConfig.Greet Then
        Mod_QINI.WriteText MainINI, "UI", "UI.Config.Greet", "True"
    Else
        Mod_QINI.WriteText MainINI, "UI", "UI.Config.Greet", "False"
    End If

    If UIConfig.QUISkin Then
        Mod_QINI.WriteText MainINI, "UI", "UI.Config.QUISkin", "True"
    Else
        Mod_QINI.WriteText MainINI, "UI", "UI.Config.QUISkin", "False"
    End If

    If UIConfig.useBackPicture Then
        Mod_QINI.WriteText MainINI, "UI", "UI.Config.useBackPicture", "True"
    Else
        Mod_QINI.WriteText MainINI, "UI", "UI.Config.useBackPicture", "False"
    End If

    Mod_QINI.WriteText MainINI, "UI", "UI.Config.AppBackPicturePath", UIConfig.AppBackPicturePath
    Exit Function
Err:
    QDB.Runtime_Error "Mod_Main->SaveConfig", Err.Description, Err.Number
    Resume Next
End Function

Public Function LoadGroup() As String
    On Error GoTo Err
    Dim i As Integer
    For i = 1 To GroupsCount
        With Groups(i)
            .Name = Mod_QINI.GetText(MainINI, "Config." & i, "Config.Name")
            .Path = Mod_QINI.GetText(MainINI, "Config." & i, "Config.Path")
            .Description = Mod_QINI.GetText(MainINI, "Config." & i, "Config.Description")
        End With
    Next
    Exit Function
Err:
    QDB.Runtime_Error "Mod_Main->LoadGroup", Err.Description, Err.Number
    Resume Next
End Function

Public Function LoadItem(paramGroupIndex As Integer) As String
    On Error GoTo Err
    Dim i As Integer, Path As String
    Path = Mod_QINI.GetText(MainINI, "Config." & paramGroupIndex, "Config.Path")
    For i = 1 To ItemsCount(paramGroupIndex)
        With Items(paramGroupIndex)
            .ItemName(i) = Mod_QINI.GetText(Path, "Config.Item." & i, "Config.Item.Name")
            .Path(i) = Mod_QINI.GetText(Path, "Config.Item." & i, "Config.Item.Path")
            .Description(i) = Mod_QINI.GetText(Path, "Config.Item." & i, "Config.Item.Description")
        End With
    Next
    Exit Function
Err:
    QDB.Runtime_Error "Mod_Main->LoadItem", Err.Description, Err.Number
    Resume Next
End Function

Public Function SaveGroup() As String
    On Error GoTo Err
    Dim i As Integer
    For i = 1 To GroupsCount
        With Groups(i)
            Mod_QINI.WriteText MainINI, "Config." & i, "Config.Name", .Name
            Mod_QINI.WriteText MainINI, "Config." & i, "Config.Path", .Path
            Mod_QINI.WriteText MainINI, "Config." & i, "Config.Description", .Description

            Mod_QINI.WriteText .Path, "DATOOLS", "Config.Name", .Name
            Mod_QINI.WriteText .Path, "DATOOLS", "Config.Path", .Path
            Mod_QINI.WriteText .Path, "DATOOLS", "Config.Description", .Description
        End With
    Next
    Mod_QINI.WriteText MainINI, "DATOOLS", "Config.Count", Str(GroupsCount)    '保存分类数目
    Exit Function
Err:
    QDB.Runtime_Error "Mod_Main->SaveGroup", Err.Description, Err.Number
    Resume Next
End Function

Public Function SaveItem(paramGroupIndex As Integer) As String
    On Error GoTo Err
    Dim i As Integer, Path As String
    Path = Mod_QINI.GetText(MainINI, "Config." & paramGroupIndex, "Config.Path")
    For i = 1 To ItemsCount(paramGroupIndex)
        With Items(paramGroupIndex)
            Mod_QINI.WriteText Path, "Config.Item." & i, "Config.Item.Name", .ItemName(i)
            Mod_QINI.WriteText Path, "Config.Item." & i, "Config.Item.Path", .Path(i)
            Mod_QINI.WriteText Path, "Config.Item." & i, "Config.Item.Description", .Description(i)
        End With
    Next
    Mod_QINI.WriteText MainINI, "Config." & paramGroupIndex, "Config.Item.Count", Str(ItemsCount(paramGroupIndex))
    Mod_QINI.WriteText Path, "DATOOLS", "Config.Item.Count", Str(ItemsCount(paramGroupIndex))
    Exit Function
Err:
    QDB.Runtime_Error "Mod_Main->SaveItem", Err.Description, Err.Number
    Resume Next
End Function

Public Function CreateGroup(GroupName As String, GroupConfigPath As String, _
                            GroupDescription As String) As String
    On Error GoTo Err
    GroupsCount = GroupsCount + 1
    With Groups(GroupsCount)
        .Name = GroupName
        .Path = GroupConfigPath
        .Description = GroupDescription
        '写入MainConfig文件
        Mod_QINI.WriteText MainINI, "DATOOLS", "Config.Count", Str(GroupsCount)
        Mod_QINI.WriteText MainINI, "Config." & GroupsCount, "Config.Name", .Name
        Mod_QINI.WriteText MainINI, "Config." & GroupsCount, "Config.Path", .Path
        Mod_QINI.WriteText MainINI, "Config." & GroupsCount, "Config.Description", .Description

        '创建相应的GroupConfig文件
        Open .Path For Output As #121
        Close #121
        Mod_QINI.WriteText .Path, "DATOOLS", "Config.Name", .Name
        Mod_QINI.WriteText .Path, "DATOOLS", "Config.Path", .Path
        Mod_QINI.WriteText .Path, "DATOOLS", "Config.Description", .Description
    End With
    Exit Function
Err:
    QDB.Runtime_Error "Mod_Main->CreateGroup", Err.Description, Err.Number
    Resume Next
End Function

Public Function CreateItem(paramGroupIndex As Integer, ItemName As String, _
                           ItemPath As String, ItemDescription As String) As String
    On Error GoTo Err
    If paramGroupIndex > GroupsCount Then
        CreateItem = "[Error]"
    End If
    ItemsCount(paramGroupIndex) = ItemsCount(paramGroupIndex) + 1
    With Items(ItemsCount(paramGroupIndex))
        .ItemName(ItemsCount(paramGroupIndex)) = ItemName
        .Path(ItemsCount(paramGroupIndex)) = ItemPath
        .Description(ItemsCount(paramGroupIndex)) = ItemDescription
        Mod_QINI.WriteText Groups(paramGroupIndex).Path, "Config.Item." & ItemsCount(paramGroupIndex), "Config.Item.Name", .ItemName(ItemsCount(paramGroupIndex))
        Mod_QINI.WriteText Groups(paramGroupIndex).Path, "Config.Item." & ItemsCount(paramGroupIndex), "Config.Item.Path", .Path(ItemsCount(paramGroupIndex))
        Mod_QINI.WriteText Groups(paramGroupIndex).Path, "Config.Item." & ItemsCount(paramGroupIndex), "Config.Item.Description", .Description(ItemsCount(paramGroupIndex))
    End With
    Mod_QINI.WriteText Groups(paramGroupIndex).Path, "DATOOLS", "Config.Item.Count", Str(ItemsCount(paramGroupIndex))
    Mod_QINI.WriteText MainINI, "Config." & paramGroupIndex, "Config.Item.Count", Str(ItemsCount(paramGroupIndex))
    Exit Function
Err:
    QDB.Runtime_Error "Mod_Main->CreateItem", Err.Description, Err.Number
    Resume Next
End Function

Public Function RemoveGroup(paramGroupIndex As Integer) As String    '移除分类
    On Error GoTo Err
    Dim i As Integer
    If GroupsCount = 1 Then
        With Groups(1): .Name = "": .Path = "": .Description = "": End With
        Exit Function
    End If
    For i = paramGroupIndex + 1 To GroupsCount
        With Groups(i - 1)
            .Name = Groups(i).Name
            .Path = Groups(i).Path
            .Description = Groups(i).Description
        End With
    Next
    GroupsCount = GroupsCount - 1
    Exit Function
Err:
    QDB.Runtime_Error "Mod_Main->RemoveGroup", Err.Description, Err.Number
    Resume Next
End Function

Public Function RemoveItem(paramGroupIndex As Integer, paramItemIndex As Integer) As String    '移除项目
    On Error GoTo Err
    Dim i As Integer
    If ItemsCount(paramGroupIndex) = 1 Then
        With Items(paramGroupIndex): .ItemName(1) = "": .Path(1) = "": .Description(1) = "": End With
        Exit Function
    End If
    For i = paramItemIndex + 1 To ItemsCount(paramGroupIndex)
        With Items(paramGroupIndex)
            .ItemName(i - 1) = Items(paramGroupIndex).ItemName(i)
            .Path(i - 1) = Items(paramGroupIndex).Path(i)
            .Description(i - 1) = Items(paramGroupIndex).Description(i)
        End With
    Next
    ItemsCount(paramGroupIndex) = ItemsCount(paramGroupIndex) - 1
    Exit Function
Err:
    QDB.Runtime_Error "Mod_Main->RemoveItem", Err.Description, Err.Number
    Resume Next
End Function

Public Function CommandProc(ParamString As String) As Boolean
    On Error GoTo Err
    Dim CommandString As String, CmdParam As String
    If InStr(1, ParamString, " ") <> 0 Then
        CommandString = Mid(ParamString, 1, InStr(1, ParamString, " ") - 1)
    Else
        CommandString = ParamString
    End If
    Select Case CommandString
    Case "project"
        Shell "explorer H:\_code\vb\Tools\DATOOLS", vbNormalFocus
        Shell "explorer H:\_code\vb\Tools\DATOOLS\DATOOLSproj.vbp", vbNormalFocus
        fExit = True
        Unload Frm_Main
        CommandProc = True
    Case "curve"
        Frm_Curve.Show
        CommandProc = True
    Case "exit"
        fExit = True
        Dim tmpForm As Form
        For Each tmpForm In Forms
            Unload tmpForm
        Next
        'Unload Frm_Main
    Case Else
        CommandProc = False
    End Select
    Exit Function
Err:
    QDB.Runtime_Error "Mod_Main->CommandProc", Err.Description, Err.Number
    CommandProc = False
    Resume Next
End Function

Function ConfigBackup(paramBackupName As String)
    On Error GoTo Err
    If Len(Dir(App.Path & "\Backup", vbDirectory)) = 0 Then
        MkDir App.Path & "\Backup"
    End If
    'FileCopy App.Path & "\Config", App.Path & "\Backup\" & paramBackupName
    Shell "cmd /c title DATOOLS & xcopy /e /c /v " & App.Path & "\Config " & App.Path & "\Backup\" & paramBackupName, vbNormalFocus
    Exit Function
Err:
    QDB.Runtime_Error "Mod_Main->ConfigBackup", Err.Description, Err.Number
    Resume Next
End Function

Function IsChinese(paramStr As String) As Boolean
    If Asc(paramStr) > 128 Or Asc(paramStr) < 0 Then
        IsChinese = True
    Else
        IsChinese = False
    End If
End Function

