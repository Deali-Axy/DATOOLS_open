Attribute VB_Name = "QMod_AppConfig"
'[Mod Name]Code Institute Common QApp Framework Config Module
'[Mod Author]Deali-Axy


Public Sub QApp_Main()
    On Error GoTo Main_Err
    QDB.Log "QApp Run! Name=" & QApp.Name
    QDB.Log "QApp ThreadID=" & QApp.ThreadID
    QDB.Log "QApp hInstance=" & QApp.hInstance
    App_Icon_Gif = QApp.Path & "\CI_Icon.gif"

    Load QFrm_Main
    QDB.Log "Load QFrm_main"
    With QFrm_Main
        .Caption = QApp_Title
        #If MLC_HookSkin Then
            Mod_HookSkinner.Attach .hwnd
            QDB.Log "Load QHookSkin"
        #End If
    End With

    #If App_Load_Interface Then
        Load QFrm_Load
        QDB.Log "Load QFrm_Load"
        With QFrm_Load
            .Caption = QApp.Title & "  ÕýÔÚ¼ÓÔØ..."
            .Show
            QDB.Log "QFrm_Load.Show"
        End With
    #Else
        QFrm_Main.Show
        QDB.Log "QFrm_Main.Show"
    #End If

    Exit Sub
Main_Err:
    QDB.Runtime_Error "Sub Main", Err.Description, Err.Number
End Sub

