VERSION 5.00
Begin VB.Form Form5 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Your system"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   1425
      TabIndex        =   4
      Top             =   1500
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Windows Ver"
      Height          =   315
      Left            =   75
      TabIndex        =   3
      Top             =   1500
      Width           =   1290
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   225
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "C"
      Top             =   2175
      Width           =   2865
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   240
      Left            =   1725
      TabIndex        =   1
      Top             =   675
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Available on drive C:"
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   675
      Width           =   1455
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' show diskspace (uses a declaration from the module)

Option Explicit
Public Function DiskSpace(DrivePath As String) As Double
  Dim Drive As String
  Dim SectorsPerCluster As Long, BytesPerSector As Long
  Dim NumberOfFreeClusters As Long, TotalClusters As Long, Sts As Long
  Dim DS
  Drive = Left(Trim(DrivePath), 1) & ":\"
  Sts = GetDiskFreeSpace(Drive, SectorsPerCluster, BytesPerSector, NumberOfFreeClusters, TotalClusters)
  If Sts <> 0 Then
    DiskSpace = SectorsPerCluster * BytesPerSector * NumberOfFreeClusters
    DS = Format$(DiskSpace, "###,###")
    Label2 = DS & " bytes"
  Else
    DiskSpace = -1
  End If
End Function

Private Sub Command1_Click()
 Form6.Show
 Unload Me
End Sub

Private Sub Command2_Click()
 Unload Me
End Sub

Private Sub Form_Load()
' display the space available...
Dim x
If Text1 = "" Then
    MsgBox "Failed to display free disk space!", vbExclamation, "Error"
Else
    x = DiskSpace(Text1.Text)
End If

End Sub
