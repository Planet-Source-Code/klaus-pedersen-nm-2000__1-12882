VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form6"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4650
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   315
      Left            =   75
      TabIndex        =   2
      Top             =   1425
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   1275
      TabIndex        =   1
      Top             =   675
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "You are using:"
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   675
      Width           =   1020
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' make the functions explicit, show Types and declare the functions
Option Explicit

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    End Type
    '
    Private Declare Function GetVersion Lib "kernel32" () As Long
    Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Function GetWindowsVersion() As String
' get the version number. if version = 5.00 then you run 2000 :)
Dim OSName As String
Dim APIName As String
Dim OSVI As OSVERSIONINFO
Dim WinVer As Long

OSVI.dwOSVersionInfoSize = 148
GetVersionEx OSVI
WinVer = GetVersion()
If WinVer & &H8000000 Then
    If OSVI.dwMajorVersion >= 4 Then
    OSName = "Windows 95"
    APIName = "Win 32"
  ElseIf OSVI.dwMinorVersion = 95 Then
    OSName = "Windows 95"
    APIName = "Win 16"
  Else
    OSName = "Windows"
    APIName = "Win32s"
  End If
Else
  OSName = "Windows NT"
  APIName = "Win32"
End If
GetWindowsVersion = OSName & " v" & OSVI.dwMajorVersion & "." & Format(OSVI.dwMinorVersion, "00") & " running " & APIName
End Function


Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Form_Load()
' start it all
 Label2.Caption = GetWindowsVersion()
End Sub
