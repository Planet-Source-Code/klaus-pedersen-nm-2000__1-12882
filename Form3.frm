VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Your system"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   1275
      TabIndex        =   5
      Top             =   1500
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CPU Type"
      Height          =   315
      Left            =   75
      TabIndex        =   4
      Top             =   1500
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   75
      TabIndex        =   1
      Top             =   300
      Width           =   4515
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   195
         Left            =   1500
         TabIndex        =   3
         Top             =   225
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Memory available:"
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   225
         Width           =   1275
      End
   End
   Begin VB.Label Label1 
      Caption         =   "This box gives you information on your CPU type, memory, etc.."
      Height          =   240
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   4515
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' declare procedure and TYPE the options for it
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
    End Type

Private Sub Command1_Click()
 Form4.Show
 Unload Me
End Sub

Private Sub Command2_Click()
 Unload Me
End Sub

Private Sub Form_Load()
' display available memory in label3 at startup
Dim ms As MEMORYSTATUS

ms.dwLength = Len(ms)
GlobalMemoryStatus ms
Label3.Caption = "Total: " & ms.dwTotalPhys & Chr(13) & "Available: " & ms.dwAvailPhys & Chr(13) & "Memory load: " & dwMemoryLoad


End Sub
