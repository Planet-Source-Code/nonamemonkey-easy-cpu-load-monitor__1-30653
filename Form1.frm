VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   390
   ClientLeft      =   750
   ClientTop       =   855
   ClientWidth     =   1695
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   390
   ScaleWidth      =   1695
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.Menu mMnu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mDelay 
         Caption         =   "Delay"
         Begin VB.Menu m5 
            Caption         =   "5 sec"
         End
         Begin VB.Menu m4 
            Caption         =   "4 sec"
         End
         Begin VB.Menu m3 
            Caption         =   "3 sec"
         End
         Begin VB.Menu m2 
            Caption         =   "2 sec"
         End
         Begin VB.Menu m1 
            Caption         =   "1 sec"
         End
         Begin VB.Menu m750 
            Caption         =   "750ms"
         End
         Begin VB.Menu m500 
            Caption         =   "500ms"
         End
         Begin VB.Menu m250 
            Caption         =   "250ms"
         End
         Begin VB.Menu m100 
            Caption         =   "100ms"
         End
         Begin VB.Menu m50 
            Caption         =   "50ms"
         End
      End
      Begin VB.Menu mPos 
         Caption         =   "Position"
         Begin VB.Menu mTL 
            Caption         =   "Top-Left"
         End
         Begin VB.Menu mTR 
            Caption         =   "Top-Right"
         End
         Begin VB.Menu mBL 
            Caption         =   "Bottom-Left"
         End
         Begin VB.Menu mBR 
            Caption         =   "Bottom-Right"
         End
      End
      Begin VB.Menu mSize 
         Caption         =   "Size"
         Begin VB.Menu mWM 
            Caption         =   "W"
            Begin VB.Menu mW 
               Caption         =   "1"
               Index           =   1
            End
            Begin VB.Menu mW 
               Caption         =   "2"
               Index           =   2
            End
            Begin VB.Menu mW 
               Caption         =   "3"
               Index           =   3
            End
            Begin VB.Menu mW 
               Caption         =   "4"
               Index           =   4
            End
         End
         Begin VB.Menu mHM 
            Caption         =   "H"
            Begin VB.Menu mH 
               Caption         =   "1"
               Index           =   1
            End
            Begin VB.Menu mH 
               Caption         =   "2"
               Index           =   2
            End
            Begin VB.Menu mH 
               Caption         =   "3"
               Index           =   3
            End
            Begin VB.Menu mH 
               Caption         =   "4"
               Index           =   4
            End
         End
      End
      Begin VB.Menu mOnTop 
         Caption         =   "On Top"
         Checked         =   -1  'True
      End
      Begin VB.Menu mASm 
         Caption         =   "AutoStart"
         Begin VB.Menu mAS 
            Caption         =   "Run under User"
            Index           =   0
         End
         Begin VB.Menu mAS 
            Caption         =   "Run as Service"
            Index           =   1
         End
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Call CPUStart
Timer1.Interval = 100
Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call CPUClose
End Sub

Private Sub Timer1_Timer()
Label1.Caption = CPUUsage()
Label1.Refresh
End Sub
