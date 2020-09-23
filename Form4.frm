VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Your system"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Disk Space"
      Height          =   315
      Left            =   2850
      TabIndex        =   2
      Top             =   1500
      Width           =   1440
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   1500
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show CPU Type"
      Height          =   315
      Left            =   75
      TabIndex        =   0
      Top             =   1500
      Width           =   1665
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this is how to show the cpu type and no. of processors installed.
' first, get the Private Type and declare the function...
Option Explicit

Private Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOfProcessores As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type

Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)


Private Sub Command1_Click()
' then display it when the button is clicked...
Dim sys As SYSTEM_INFO
GetSystemInfo sys
Print "Processor type: "; sys.dwProcessorType
Print "No. Processors: "; sys.dwNumberOfProcessores
End Sub

Private Sub Command2_Click()
 Unload Me
End Sub

Private Sub Command3_Click()
 Form5.Show
 Unload Me
End Sub

Private Sub Form_Load()

End Sub
