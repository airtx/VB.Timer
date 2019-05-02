VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1725
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   4440
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton Command5 
      Caption         =   "End"
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   1080
   End
   Begin VB.CommandButton Command3 
      Caption         =   "start"
      Height          =   735
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   1
      Text            =   "00"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Text            =   "5"
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private sec_Time As Integer

Private Sub Command1_Click()
    Text1.Text = CStr(Val(Text1.Text) + 5)
End Sub

Private Sub Command2_Click()
Text1.Text = CStr(Val(Text1.Text) - 5)
End Sub

Private Sub Command3_Click()
Dim min As Integer
Dim sec As Integer

min = Val(Text1.Text)
sec = Val(Text2.Text)
sec_Time = min * 60 + sec

Timer1.Interval = 1000
Timer1.Enabled = True



End Sub

Private Sub Command4_Click()
    sec_Time = 300
    Text1.Text = "5"
    Text2.Text = ""
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Form_Activate()
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    If (sec_Time > 0) Then
        sec_Time = sec_Time - 1
        Call Sub_ShowTime
        
    Else
        Timer1.Enabled = False
        v = MsgBox("Time is Over", 1, "Hello")
        
    End If
End Sub
Private Sub Sub_RunApp()
    Dim RetVal
        RetVal = Shell("explorer", 1)
        
End Sub

Private Sub Sub_ShowTime()
    Dim min As Integer
    Dim sec As Integer
    
    min = CInt(sec_Time \ 60)
    sec = sec_Time Mod 60
    
    Text1.Text = CStr(min)
    Text2.Text = CStr(sec)
End Sub
