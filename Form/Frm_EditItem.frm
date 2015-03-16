VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frm_EditItem 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ìí¼ÓÏîÄ¿"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5250
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   15.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   Icon            =   "Frm_EditItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   5250
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin MSForms.CommandButton Btn_Save 
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   975
      VariousPropertyBits=   19
      Caption         =   "±£´æ"
      Size            =   "1720;1085"
      FontName        =   "Î¢ÈíÑÅºÚ"
      FontHeight      =   315
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Btn_Cancel 
      Height          =   615
      Left            =   4200
      TabIndex        =   8
      Top             =   4320
      Width           =   975
      VariousPropertyBits=   19
      Caption         =   "È¡Ïû"
      Size            =   "1720;1085"
      FontName        =   "Î¢ÈíÑÅºÚ"
      FontHeight      =   315
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Btn_Open 
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      Top             =   1200
      Width           =   735
      VariousPropertyBits=   19
      Caption         =   "open"
      Size            =   "1296;873"
      TakeFocusOnClick=   0   'False
      FontName        =   "Î¢ÈíÑÅºÚ"
      FontHeight      =   195
      FontCharSet     =   134
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox Txt_Path 
      Height          =   495
      Left            =   1200
      TabIndex        =   6
      Top             =   1200
      Width           =   2895
      VariousPropertyBits=   746604563
      Size            =   "5106;873"
      FontName        =   "Î¢ÈíÑÅºÚ"
      FontHeight      =   315
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Txt_Name 
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   480
      Width           =   3735
      VariousPropertyBits=   746604563
      Size            =   "6588;873"
      FontName        =   "Î¢ÈíÑÅºÚ"
      FontHeight      =   315
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Txt_Description 
      Height          =   1455
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   4575
      VariousPropertyBits=   746604563
      Size            =   "8070;2566"
      FontName        =   "Î¢ÈíÑÅºÚ"
      FontHeight      =   315
      FontCharSet     =   134
      FontPitchAndFamily=   34
   End
   Begin VB.Label lbl_Trip 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Shape Shape_Main 
      BorderColor     =   &H00E0E0E0&
      Height          =   3975
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label lbl_Description 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÃèÊö"
      ForeColor       =   &H00E0E0E0&
      Height          =   420
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   630
   End
   Begin VB.Label lbl_Path 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Â·¾¶"
      ForeColor       =   &H00E0E0E0&
      Height          =   420
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   630
   End
   Begin VB.Label lbl_Name 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ãû³Æ"
      ForeColor       =   &H00E0E0E0&
      Height          =   420
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   630
   End
End
Attribute VB_Name = "Frm_EditItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const fADDGROUP As Integer = 1
Const fADDITEM As Integer = 2
Const fEDITGROUP As Integer = 3
Const fEDITITEM As Integer = 4
Dim QD As New Cls_QDialog
Dim Flag As Integer, IsChange As Boolean
Attribute IsChange.VB_VarUserMemId = 1073938433

Private Sub Btn_Cancel_Click()
    Unload Me
End Sub

Private Sub Btn_Open_Click()
    With QD
        .ShowOpen
        Txt_Path = .File
    End With
End Sub

Private Sub Btn_Save_Click()
    Debug.Print "[" & Me.Name & ".Btn_Save_Click] Flag=" & Flag
    On Error GoTo Err
    If Txt_Name = "" Then
        lbl_Trip = "Ãû³Æ²»ÄÜÎª¿Õ£¡"
        Exit Sub
    End If
    If Txt_Path = "" Then
        lbl_Trip = "Â·¾¶²»ÄÜÎª¿Õ£¡"
        Exit Sub
    End If
    Select Case Flag
    Case fADDGROUP
        CreateGroup Txt_Name, Txt_Path, Txt_Description
    Case fADDITEM
        CreateItem GroupIndex, Txt_Name, Txt_Path, Txt_Description
    Case fEDITGROUP
        With Groups(GroupIndex)
            .Name = Txt_Name
            .Path = Txt_Path
            .Description = Txt_Description
        End With
        SaveGroup
        IsChange = True
    Case fEDITITEM
        With Items(GroupIndex)
            .ItemName(ItemIndex(GroupIndex)) = Txt_Name
            .Path(ItemIndex(GroupIndex)) = Txt_Path
            .Description(ItemIndex(GroupIndex)) = Txt_Description
        End With
        SaveItem GroupIndex
        IsChange = True
    End Select
    Frm_Main.Msg_Proc "refresh"
    Unload Me
    Exit Sub
Err:
    QDB.Runtime_Error Me.Name & "->Btn_Save_Click", Err.Description, Err.Number
    Resume Next
End Sub

Public Sub Msg_Proc(Msg As String, Optional Source As String)
    Select Case Msg
    Case "addgroup"
        Me.Caption = "Ìí¼Ó·ÖÀà"
        Flag = fADDGROUP
        Txt_Path = App.Path & "\Config\group." & GroupsCount + 1 & ".config"
    Case "additem"
        Me.Caption = "Ìí¼ÓÏîÄ¿"
        Flag = fADDITEM
    Case "editgroup"
        Me.Caption = "±à¼­·ÖÀà"
        Flag = fEDITGROUP
        Txt_Name = Groups(GroupIndex).Name
        Txt_Path = Groups(GroupIndex).Path
    Case "edititem"
        Me.Caption = "±à¼­ÏîÄ¿"
        Flag = fEDITITEM
        Txt_Name = Items(GroupIndex).ItemName(ItemIndex(GroupIndex))
        Txt_Path = Items(GroupIndex).Path(ItemIndex(GroupIndex))
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13
        Btn_Save_Click
    Case 27
        Btn_Cancel_Click
    End Select
End Sub

Private Sub Form_Load()
    mLoadConfig
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If IsChange Then Frm_Main.Msg_Proc "refresh"
End Sub

Private Sub mLoadConfig()
    On Error GoTo Err
    Me.backColor = QApp_BackColor
    Me.foreColor = QApp_ForeColor
    Dim c As Control
    For Each c In Me.Controls
        c.foreColor = QApp_ForeColor
    Next
    If UIConfig.useBackPicture Then
        Set Me.Picture = UIConfig.AppBackPicture
    End If
    Exit Sub
Err:
    QDB.Runtime_Error Me.Name & "->mLoadConfig", Err.Description, Err.Number
    Resume Next
End Sub

