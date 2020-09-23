VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NM 2000"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   6675
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "NM 2000"
            TextSave        =   "NM 2000"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   176
            MinWidth        =   176
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "00:08"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "19-11-2000"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CDlg 
      Left            =   1125
      Top             =   6225
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5190
      Left            =   75
      TabIndex        =   4
      Top             =   1350
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   9155
      _Version        =   327680
      Style           =   1
      Tabs            =   6
      Tab             =   4
      TabsPerRow      =   6
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Welcome"
      TabPicture(0)   =   "Form1.frx":030A
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "Messages"
      TabPicture(1)   =   "Form1.frx":0326
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "Hints"
      TabPicture(2)   =   "Form1.frx":0342
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "Misc"
      TabPicture(3)   =   "Form1.frx":035E
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(0).Enabled=   0   'False
      TabCaption(4)   =   "Help / About"
      TabPicture(4)   =   "Form1.frx":037A
      Tab(4).ControlCount=   1
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Frame6"
      Tab(4).Control(0).Enabled=   0   'False
      TabCaption(5)   =   "Exit"
      TabPicture(5)   =   "Form1.frx":0396
      Tab(5).ControlCount=   1
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame7"
      Tab(5).Control(0).Enabled=   0   'False
      Begin VB.Frame Frame7 
         Height          =   4665
         Left            =   -74850
         TabIndex        =   11
         Top             =   375
         Width           =   7890
         Begin VB.CommandButton Command5 
            Caption         =   "E&xit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   3150
            TabIndex        =   30
            ToolTipText     =   "Click here to exit the program"
            Top             =   3900
            Width           =   1590
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Created by Klaus Pedersen"
            Height          =   195
            Left            =   2985
            TabIndex        =   54
            Top             =   4350
            Width           =   1920
         End
         Begin VB.Image Image4 
            BorderStyle     =   1  'Fixed Single
            Height          =   1995
            Left            =   1050
            Picture         =   "Form1.frx":03B2
            Top             =   975
            Width           =   5835
         End
         Begin VB.Label Label14 
            Caption         =   $"Form1.frx":4B84
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   150
            TabIndex        =   29
            Top             =   225
            Width           =   7590
         End
      End
      Begin VB.Frame Frame6 
         Height          =   4665
         Left            =   150
         TabIndex        =   10
         Top             =   375
         Width           =   7890
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Help"
            Height          =   315
            Left            =   3900
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   3900
            Width           =   1140
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "I made this pic for a kid'a logo ;-)"
            Height          =   240
            Left            =   225
            TabIndex        =   50
            Top             =   2025
            Width           =   2415
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "(Cool name, isn't it?) ;-)"
            Height          =   240
            Left            =   300
            TabIndex        =   49
            Top             =   1800
            Width           =   2265
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Created by Klaus Pedersen"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   3360
            TabIndex        =   28
            Top             =   1125
            Width           =   3795
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FFFFFF&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   840
            Left            =   3600
            Shape           =   4  'Rounded Rectangle
            Top             =   3600
            Width           =   1740
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "nmb13978@vip.cybercity.dk"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   405
            TabIndex        =   26
            Top             =   4125
            Width           =   2055
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00FFFFFF&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   915
            Left            =   150
            Shape           =   4  'Rounded Rectangle
            Top             =   3525
            Width           =   2565
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   $"Form1.frx":4C30
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   690
            Left            =   2700
            TabIndex        =   25
            Top             =   3000
            Width           =   4965
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C00000&
            X1              =   2700
            X2              =   5325
            Y1              =   2925
            Y2              =   2925
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NM 2000 by Klaus Pedersen"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   2700
            TabIndex        =   24
            Top             =   2700
            Width           =   1995
         End
         Begin VB.Image Image3 
            Height          =   1275
            Left            =   375
            Picture         =   "Form1.frx":4CD3
            Top             =   450
            Width           =   2130
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00FFFFFF&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   1140
            Left            =   150
            Shape           =   4  'Rounded Rectangle
            Top             =   2625
            Width           =   7590
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            BorderColor     =   &H00FFFFFF&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   2790
            Left            =   150
            Shape           =   4  'Rounded Rectangle
            Top             =   225
            Width           =   2565
         End
      End
      Begin VB.Frame Frame5 
         Height          =   4665
         Left            =   -74850
         TabIndex        =   9
         Top             =   375
         Width           =   7890
         Begin VB.Frame Frame13 
            Caption         =   "Your computer"
            Height          =   1590
            Left            =   300
            TabIndex        =   42
            Top             =   2625
            Width           =   7365
            Begin VB.CommandButton Command11 
               Caption         =   "Advanced (cpu type, windows version, disk space, etc)"
               Height          =   315
               Left            =   150
               TabIndex        =   44
               Top             =   750
               Width           =   4815
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "You are logged in as: "
               Height          =   195
               Left            =   150
               TabIndex        =   43
               Top             =   300
               Width           =   1545
            End
         End
         Begin VB.Frame Frame11 
            Height          =   90
            Left            =   300
            TabIndex        =   40
            Top             =   2400
            Width           =   7365
         End
         Begin VB.Frame Frame9 
            Caption         =   "Options"
            Height          =   1815
            Left            =   300
            TabIndex        =   36
            Top             =   375
            Width           =   2790
            Begin VB.CommandButton Command6 
               Caption         =   "Advanced"
               Height          =   315
               Left            =   150
               TabIndex        =   38
               Top             =   525
               Width           =   1215
            End
            Begin VB.CommandButton Command7 
               Caption         =   "Reset..."
               Height          =   315
               Left            =   1425
               TabIndex        =   37
               Top             =   525
               Width           =   1215
            End
            Begin VB.Label Label15 
               Caption         =   "Notice: The reset button will not reset the advanced options!"
               Height          =   465
               Left            =   150
               TabIndex        =   39
               Top             =   1275
               Width           =   2490
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Basic options"
            Height          =   1815
            Left            =   3375
            TabIndex        =   31
            Top             =   375
            Width           =   4290
            Begin VB.CheckBox Check4 
               Caption         =   "&Display author note"
               Height          =   240
               Left            =   150
               TabIndex        =   35
               Top             =   1275
               Value           =   1  'Checked
               Width           =   1815
            End
            Begin VB.CheckBox Check3 
               Caption         =   "&Hide Tip of The Day"
               Height          =   240
               Left            =   150
               TabIndex        =   34
               Top             =   975
               Width           =   1815
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Show &Topinfo"
               Height          =   240
               Left            =   150
               TabIndex        =   33
               Top             =   675
               Value           =   1  'Checked
               Width           =   1815
            End
            Begin VB.CheckBox Check1 
               Caption         =   "&Show Statusbar"
               Height          =   240
               Left            =   150
               TabIndex        =   32
               Top             =   375
               Value           =   1  'Checked
               Width           =   1815
            End
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Created by Klaus Pedersen"
            Height          =   195
            Left            =   2985
            TabIndex        =   53
            Top             =   4350
            Width           =   1920
         End
      End
      Begin VB.Frame Frame4 
         Height          =   4665
         Left            =   -74850
         TabIndex        =   8
         Top             =   375
         Width           =   7890
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   3000
            Left            =   6375
            Top             =   3075
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   2
            Left            =   5925
            Top             =   3075
         End
         Begin VB.Frame Frame12 
            Height          =   90
            Left            =   150
            TabIndex        =   41
            Top             =   750
            Width           =   7590
         End
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Created by Klaus Pedersen"
            Height          =   195
            Left            =   2985
            TabIndex        =   52
            Top             =   4350
            Width           =   1920
         End
         Begin VB.Label Label20 
            Caption         =   "- Make some more advanced effects in this program"
            Height          =   315
            Left            =   225
            TabIndex        =   48
            Top             =   1575
            Width           =   4365
         End
         Begin VB.Label Label19 
            Caption         =   "- Chat console, where network users can join!"
            Height          =   240
            Left            =   225
            TabIndex        =   47
            Top             =   1275
            Width           =   5265
         End
         Begin VB.Label Label18 
            Caption         =   "- The possibility to send messages without using the Net Send command!"
            Height          =   240
            Left            =   225
            TabIndex        =   46
            Top             =   975
            Width           =   5340
         End
         Begin VB.Label Label17 
            Caption         =   "Here's some hints and suggestiongs, that you, who views this code may try to implement:"
            Height          =   465
            Index           =   0
            Left            =   225
            TabIndex        =   45
            Top             =   225
            Width           =   7515
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4665
         Left            =   -74850
         TabIndex        =   7
         Top             =   375
         Width           =   7890
         Begin VB.CommandButton Command2 
            Caption         =   "&Clear all"
            Height          =   315
            Left            =   3225
            TabIndex        =   22
            ToolTipText     =   "Cancel the send, and clears the receiver and message box (reset)"
            Top             =   2925
            Width           =   1515
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Send Message"
            Height          =   315
            Left            =   1650
            TabIndex        =   21
            ToolTipText     =   "Hit this to send the message"
            Top             =   2925
            Width           =   1515
         End
         Begin VB.Frame Frame8 
            Height          =   90
            Left            =   150
            TabIndex        =   20
            Top             =   750
            Width           =   7515
         End
         Begin VB.TextBox Text2 
            Height          =   1365
            Left            =   1650
            TabIndex        =   19
            ToolTipText     =   "Type your message here, you must write a message before sending."
            Top             =   1500
            Width           =   6015
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Left            =   1650
            TabIndex        =   17
            ToolTipText     =   "Type the name of whom you want this message to reach...as default, your own name is entered here..."
            Top             =   1050
            Width           =   6015
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Created by Klaus Pedersen"
            Height          =   195
            Left            =   2985
            TabIndex        =   51
            Top             =   4350
            Width           =   1920
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Your message:"
            Height          =   195
            Left            =   150
            TabIndex        =   18
            Top             =   1500
            Width           =   1050
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Receiver:"
            Height          =   195
            Left            =   150
            TabIndex        =   16
            Top             =   1125
            Width           =   690
         End
         Begin VB.Label Label6 
            Caption         =   $"Form1.frx":64EF
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   150
            TabIndex        =   15
            Top             =   225
            Width           =   7590
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4665
         Left            =   -74850
         TabIndex        =   6
         Top             =   375
         Width           =   7890
         Begin VB.CommandButton Command3 
            Caption         =   "See Tip of The Day"
            Height          =   315
            Left            =   2925
            TabIndex        =   23
            Top             =   3525
            Width           =   1890
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Created by Klaus Pedersen"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2970
            TabIndex        =   14
            Top             =   4350
            Width           =   1950
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "A KK Media® Production"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3060
            TabIndex        =   13
            Top             =   825
            Width           =   1770
         End
         Begin VB.Image Image2 
            Height          =   1935
            Left            =   1050
            Picture         =   "Form1.frx":65BD
            Top             =   1275
            Width           =   5775
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Welcome to NM 2000"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2400
            TabIndex        =   12
            Top             =   225
            Width           =   3090
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   90
      Left            =   -150
      TabIndex        =   1
      Top             =   1125
      Width           =   8715
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   0
      ScaleHeight     =   1140
      ScaleWidth      =   8490
      TabIndex        =   0
      Top             =   0
      Width           =   8490
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "This program sends Messages throughout the network by using the Net Send Command. Very easy to handle!!!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   375
         TabIndex        =   3
         Top             =   375
         Width           =   3765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "NM 2000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   2
         Top             =   150
         Width           =   2490
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   7575
         Picture         =   "Form1.frx":AD8F
         Top             =   300
         Width           =   480
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'
' This is a free-to-learn-by project by Klaus Pedersen
'
' This is one of my larger applications, and I would like
' you to take a look on it and see if you can learn something
' from this code.
'
' You may contact me at:
'
'  E-mail   :   nmb13978@vip.cybercity.dk
'  ICQ      :   53099224 (programmers will always be autorized ;-)
'
' Happy VB programming!
' - Klaus Pedersen

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Check1_Click()
 If Check1.Value = 1 Then
  StatusBar1.Visible = True
 Else
  StatusBar1.Visible = False
 End If
End Sub

Private Sub Check2_Click()
 If Check2.Value = 1 Then
  Picture1.Visible = True
 Else
  Picture1.Visible = False
 End If
End Sub

Private Sub Check3_Click()
 If Check3.Value = 1 Then
  Command3.Visible = False
 Else
  Command3.Visible = True
 End If
End Sub

Private Sub Check4_Click()
 If Check4.Value = 1 Then
  Label5.Visible = True
  Label24.Visible = True
  Label25.Visible = True
  Label26.Visible = True
  Label16.Visible = True
 Else
  Label5.Visible = False
  Label24.Visible = False
  Label25.Visible = False
  Label26.Visible = False
  Label16.Visible = False
 End If
End Sub

Private Sub Command1_Click()
 ' message sequence, make a string, and put it all together, if any errors, show them. If correct, SEND MESSAGE!!!
 Dim CombineNameMsg
CombineNameMsg = " " & Text1.Text & " " & Text2.Text
 If Text1.Text = "" Then
  MsgBox ("You must show me who to send the message to!"), vbExclamation, "Error"
 ElseIf Text1.Text = Environ$("username") Then
   MsgBox ("Your names must be unique - you can't send messages to a user, who has the same name as you ") & "(" & Environ$("username") & ")", vbExclamation, "Error"
 ElseIf Text2.Text = "" Then
   MsgBox ("You must supply a message!"), vbExclamation, "Error"
 Else
  retval = Shell("net send " & CombineNameMsg, vbHide)
  End If
End Sub

Private Sub Command10_Click()
' reset text3, text4 and text5
 Text3.Text = ""
 Text4.Text = ""
 Text5.Text = ""
End Sub

Private Sub Command11_Click()
' show a form
 Form3.Show
End Sub

Private Sub Command2_Click()
' reset text1 and text2
 Text1.Text = ""
 Text2.Text = ""
End Sub

Private Sub Command3_Click()
 frmTip.Show
End Sub

Private Sub Command4_Click()
' execute help file (the help file is not included)
 CDlg.HelpFile = "nm2k.hlp"
 CDlg.HelpCommand = cdlHelpContents
 CDlg.ShowHelp
End Sub

Private Sub Command5_Click()
 ' enable the timer
 Timer1.Enabled = True
End Sub

Private Sub Command6_Click()
 Form2.Show
 ' make a little effect, to make a better impression!
    Dim Counter As Integer
    Dim Workarea(5550) As String
    Form2.ProgressBar1.Min = LBound(Workarea)
    Form2.ProgressBar1.Max = UBound(Workarea)
    Form2.ProgressBar1.Visible = True
    Form2.ProgressBar1.Value = Form2.ProgressBar1.Min
    For Counter = LBound(Workarea) To UBound(Workarea)
        Workarea(Counter) = "Initial value" & Counter
        Form2.ProgressBar1.Value = Counter
Next Counter
    Form2.ProgressBar1.Visible = True
    Form2.ProgressBar1.Value = Form2.ProgressBar1.Min
 Unload Form2
 Form7.Show
End Sub

Private Sub Command7_Click()
 ' reset checks to their default values
 Check1.Value = 1
 Check2.Value = 1
 Check3.Value = 0
 Check4.Value = 1
End Sub

Private Sub Form_Load()
' initialize all variables and set text
 Text1.Text = Environ$("username")
 Load Form1
 Load Form2
 Load Form3
 Load Form4
 Load Form5
 Load Form6
 Load Form7
 Load frmTip
 Label22.Caption = "You are logged in as: " & Environ$("username") & "."
End Sub


Private Sub Label12_Click()
' a hyperlink, when clicked, execute the email
    ShellExecute 0&, vbNullString, "nmb13978@vip.cybercity.dk", vbNullString, "C:\", SW_SHOWNORMAL
End Sub

Private Sub Timer1_Timer()
' slide show
 Form1.Width = Form1.Width - 90
  Form1.Height = Form1.Height - 90
   Form1.Top = Form1.Top + 5
    Form1.Left = Form1.Left + 80
     Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
' unload the entire App from memory and then end
 Unload Form1
 Unload Form2
 Unload Form3
 Unload Form4
 Unload Form5
 Unload Form6
 Unload Form7
 Unload frmTip
 End
End Sub
