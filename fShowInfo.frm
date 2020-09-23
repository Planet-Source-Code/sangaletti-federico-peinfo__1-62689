VERSION 5.00
Begin VB.Form fShowInfo 
   Caption         =   "Information"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   1
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   10095
   End
End
Attribute VB_Name = "fShowInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Me.Hide
    Me.Caption = "Information"
    txtInfo.Text = vbNullString
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtInfo.Width = Me.Width - 360
    txtInfo.Height = Me.Height - 1230
    Command1.Left = (Me.Width - Command1.Width) - 300
    Command1.Top = (Me.Height - Command1.Height) - 650
End Sub
