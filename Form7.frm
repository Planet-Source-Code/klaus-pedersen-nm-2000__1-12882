VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Advanced"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3390
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Top             =   1875
      Width           =   990
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Hide ""introduction"" text"
      Height          =   315
      Left            =   150
      TabIndex        =   3
      Top             =   1350
      Width           =   3090
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Hide all images"
      Height          =   315
      Left            =   150
      TabIndex        =   2
      Top             =   1050
      Width           =   3090
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Hide Basic Options"
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   750
      Width           =   3090
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "These options can hide or show some of the items in the main program window."
      Height          =   495
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   3195
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 ' check for changes and apply if any
 If Check1.Value = 1 Then
  Form1.Frame10.Visible = False
 Else
  Form1.Frame10.Visible = True
 End If
 If Check2.Value = 1 Then
  Form1.Image1.Visible = False
  Form1.Image2.Visible = False
  Form1.Image3.Visible = False
  Form1.Image4.Visible = False
 Else
  Form1.Image1.Visible = True
  Form1.Image2.Visible = True
  Form1.Image3.Visible = True
  Form1.Image4.Visible = True
 End If
 If Check3.Value = 1 Then
  Form1.Picture1.Visible = False
 Else
  Form1.Picture1.Visible = True
 End If
  Form7.Hide
End Sub

Private Sub Form_Load()

End Sub
