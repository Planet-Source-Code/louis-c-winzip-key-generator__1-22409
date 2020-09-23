VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About KeyGen"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2565
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   2565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E-&Mail Me"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "http://lpc-systems.0pi.com"
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
      Height          =   255
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Click to goto website."
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Coder/Programmer"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1860
      Left            =   120
      Picture         =   "Form2.frx":030A
      Top             =   120
      Width           =   2325
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell ("start mailto:lpc@pirateden.com"), vbHide
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Label2_Click()
Shell ("start iexplore.exe"), vbHide
End Sub
