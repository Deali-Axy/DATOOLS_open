VERSION 5.00
Begin VB.UserControl Ctl_QCurve 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2115
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "Ctl_QCurve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event Click()
Private mvarPictureBox As PictureBox
Private mvarDataCollection As Collection
Private mvarPicBackground As OLE_COLOR
Private mvarPicForeground As OLE_COLOR
Private mvarBorderSize As Integer
Private mvarBorderColor As Long
Private mvarGridVisible As Boolean
Private mvarGridColor As Long
Dim DownloadSpeedTop As Double, UploadSpeedTop As Double, DownloadSpeedAverage As Double, UploadSpeedAverage As Double
Attribute UploadSpeedTop.VB_VarUserMemId = 1073938440
Attribute DownloadSpeedAverage.VB_VarUserMemId = 1073938440
Attribute UploadSpeedAverage.VB_VarUserMemId = 1073938440
Dim 变量 As New Collection
Attribute 变量.VB_VarUserMemId = 1073938444
Dim Hig As Long
Attribute Hig.VB_VarUserMemId = 1073938445
Dim con As Integer
Attribute con.VB_VarUserMemId = 1073938446

Public Property Let GridColor(ByVal vData As Long)
    mvarGridColor = vData
End Property

Public Property Get GridColor() As Long
    GridColor = mvarGridColor
End Property

Public Property Let GridVisible(ByVal vData As Boolean)
    mvarGridVisible = vData
End Property

Public Property Get GridVisible() As Boolean
    GridVisible = mvarGridVisible
End Property

Public Property Let BorderColor(ByVal vData As Long)
    mvarBorderColor = vData
End Property

Public Property Get BorderColor() As Long
    BorderColor = mvarBorderColor
End Property

Public Property Let BorderSize(ByVal vData As Integer)
    mvarBorderSize = vData
End Property

Public Property Get BorderSize() As Integer
    BorderStyle = mvarBorderSize
End Property

Public Property Let PicForeground(ByVal vData As Long)
    Let mvarPicForeground = vData
End Property

Public Property Get PicForeground() As Long
    PicForeground = mvarPicForeground
End Property

Public Property Let PicBackground(ByVal vData As Long)
    mvarPicBackground = vData
End Property

Public Property Get PicBackground() As Long
    PicBackground = mvarPicBackground
End Property

Public Property Set DataCollection(ByVal vData As Collection)
    Set mvarDataCollection = vData
End Property

Public Property Get DataCollection() As Collection
    Set DataCollection = mvarDataCollection
End Property

Public Property Set PictureBox(ByVal vData As PictureBox)
    Set mvarPictureBox = vData
End Property

Public Property Get PictureBox() As PictureBox
    Set PictureBox = mvarPictureBox
End Property

Public Sub Draw()
    Dim BDR As Integer, X As Integer
    Dim NewX As Double, NewY As Double
    Dim OldX As Double, OldY As Double
    Dim GridHeight As Double, GridWidth As Double
    On Error GoTo NoPicBox      ' In case the PicBox isn't set yet
    mvarPictureBox.DrawWidth = 2
    If mvarPictureBox.AutoRedraw = False Then mvarPictureBox.AutoRedraw = True
    mvarPictureBox.Cls
    BDR = mvarPictureBox.BorderStyle
    If mvarPictureBox.ScaleMode <> 3 Then mvarPictureBox.ScaleMode = 3
    If mvarPictureBox.backColor <> mvarPicBackground Then mvarPictureBox.backColor = mvarPicBackground
    If mvarBorderSize > 0 Then
        For X = 0 To mvarBorderSize
            mvarPictureBox.Line (X, X)-(mvarPictureBox.ScaleWidth - (BDR + X), mvarPictureBox.ScaleHeight - (BDR + X)), mvarBorderColor, B
        Next X
    End If
    If mvarGridVisible = True Then
        For X = 1 To 20
            mvarPictureBox.Line (mvarBorderSize, mvarBorderSize)-((((mvarPictureBox.ScaleWidth - (mvarBorderSize * 2)) / 20) * X), (mvarPictureBox.ScaleHeight - (mvarBorderSize * 2))), mvarGridColor, B
        Next X
        For X = 1 To 10
            mvarPictureBox.Line (mvarBorderSize, mvarBorderSize)-((mvarPictureBox.ScaleWidth - (mvarBorderSize * 2)), (((mvarPictureBox.ScaleHeight - (mvarBorderSize * 2)) / 10) * X)), mvarGridColor, B
        Next X
    End If

    If mvarDataCollection.Count > 0 Then
        GridHeight = ((mvarPictureBox.ScaleHeight - (mvarBorderSize * 2)) / 100) + 0    ' 0-100%
        GridWidth = ((mvarPictureBox.ScaleWidth - (mvarBorderSize * 2)) / 100) + 0      ' 1-100 Items
        Do
            If mvarDataCollection.Count > 100 Then _
               mvarDataCollection.Remove 1
        Loop While mvarDataCollection.Count > 100
        mvarPictureBox.DrawWidth = 3
        OldX = mvarBorderSize + 2
        OldY = ((mvarPictureBox.ScaleHeight - (mvarBorderSize * 2)) - (mvarDataCollection(1) * GridHeight))
        For X = 1 To 100
            NewX = (mvarPictureBox.ScaleWidth - (mvarBorderSize * 2)) - ((100 - (X - 1)) * GridWidth)
            NewY = ((mvarPictureBox.ScaleHeight - (mvarBorderSize * 2)) - (mvarDataCollection(X) * GridHeight))
            NewX = NewX + 2
            If NewX < mvarBorderSize Then NewX = mvarBorderSize
            If NewY < mvarBorderSize Then NewY = mvarBorderSize

            mvarPictureBox.Line (OldX, OldY)-(NewX, NewY), mvarPicForeground
            OldX = NewX: OldY = NewY
            If OldX < mvarBorderSize Then OldX = mvarBorderSize
            If OldY < mvarBorderSize Then OldY = mvarBorderSize
        Next X
    End If
NoPicBox:
    mvarPictureBox.DrawWidth = 2
End Sub

Public Property Let Value(ByVal vData As Long)
    变量.Add Int(Format(vData, "###,###,###,###,#0.#0")) + 5
    Call Draw
End Property

Public Property Get Value() As Long
    Value = 100
End Property

Private Sub Picture1_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
    Do: 变量.Add 0: Loop While 变量.Count < 100
    Set mvarPictureBox = Picture1
    Set DataCollection = 变量
    mvarPicForeground = vbGreen
    mvarGridColor = &H808000
    mvarGridVisible = True
End Sub

Private Sub UserControl_Resize()
    Picture1.Left = 0
    Picture1.Top = 0
    Picture1.Width = UserControl.Width
    Picture1.Height = UserControl.Height
    变量.Add Int(Format(0, "###,###,###,###,#0.#0")) + 5
    Call Draw
End Sub
