VERSION 5.00
Begin VB.Form QFrm_Main 
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8715
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   18
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "QFrm_Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   8715
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1500
      TabIndex        =   0
      Top             =   1440
      Width           =   2715
   End
End
Attribute VB_Name = "QFrm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Btn_About_Click()
    QFrm_About.Show
End Sub

Private Sub Command1_Click()
    QApp.ExitQApp
End Sub

Private Sub Form_Activate()
    On Error GoTo Err
    QDB.Log Me.Name & " Activate"
    Exit Sub
Err:
    QDB.Runtime_Error Me.Name & "_Activate", Err.Description, Err.Number
    Resume Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF1
        QFrm_About.Show
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo load_err
    Me.Caption = QApp_Title
    QDB.Log Me.Name & " Load hWnd=" & Me.hwnd
    Exit Sub
load_err:
    QDB.Runtime_Error Me.Name & "_Load", Err.Description, Err.Number
    Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    QDB.Log Me.Name & " Unload"
End Sub
