VERSION 5.00
Begin VB.Form Frm_Curve 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QCurve"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13665
   Icon            =   "Frm_Curve.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   13665
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Tmr_Curve 
      Interval        =   500
      Left            =   1560
      Top             =   4680
   End
   Begin DATOOLSproj.Ctl_QCurve QCurve_Main 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   13150
   End
End
Attribute VB_Name = "Frm_Curve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    With QCurve_Main
        .Height = Me.Height
        .Width = Me.Width
    End With
End Sub

Private Sub Tmr_Curve_Timer()
    Randomize
    QCurve_Main.Value = Rnd * 100
    Dim i As Integer
    For i = 1 To 20
        QCurve_Main.Value = 50
    Next
End Sub
