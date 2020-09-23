VERSION 5.00
Begin VB.Form frmRay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ray Tracing"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   343
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStep 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Text            =   "1"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdRender 
      Caption         =   "Render"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4680
      Width           =   975
   End
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      Height          =   4560
      Left            =   0
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   0
      Width           =   4560
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Step"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   4680
      Width           =   735
   End
End
Attribute VB_Name = "frmRay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRender_Click()
    If Running = False Then
        Running = True
        cmdRender.Caption = "Stop"
        Render pic1, txtStep
        cmdRender.Caption = "Render"
    Else
        Running = False
        cmdRender.Caption = "Render"
    End If
End Sub

Private Sub Form_Load()
    Dim Sphere1 As New Sphere
    Dim Sphere2 As New Sphere
    Dim Sphere3 As New Sphere
    Dim Light1 As New LightSource
    Me.Show
    DoEvents
    AmbIr = 128
    AmbIg = 128
    AmbIb = 128
    EyeR = 1000
    EyePhi = -0.3
    EyeTheta = 0
    Sphere1.SetValues 75, 0, 0, 50, 0.6, 0.1, 0.1, 0.6, 0.1, 0.1, 0.35, 20
    Sphere2.SetValues -30, 0, -60, 50, 0.1, 0.6, 0.1, 0.1, 0.6, 0.1, 0.35, 20
    Sphere3.SetValues -30, 0, 60, 50, 0.1, 0.1, 0.6, 0.1, 0.1, 0.6, 0.35, 20
    Light1.SetParameters 1000, -750, 1000, 255, 255, 255
    Objects.Add Sphere1
    Objects.Add Sphere2
    Objects.Add Sphere3
    LightSources.Add Light1
End Sub
