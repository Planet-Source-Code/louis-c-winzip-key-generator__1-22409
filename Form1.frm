VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "WinZip Key Generator v.7x to 8x."
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5655
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&About"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      Picture         =   "Form1.frx":030A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 1998-2001 LPC Systems Inc ®. All rights reserved."
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
      Left            =   600
      TabIndex        =   7
      ToolTipText     =   "Click to goto website."
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Output"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Input Info"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Sub Command1_Click()
'    Text2 = GenerateKey(Text1)
'End Sub

Public Function GenerateKey(ByVal Name As String) As String
    Dim lPart1 As Long
    Dim lPart2 As Long
    Dim ch As Long
    Dim i As Long
    Dim n As Long
    'Name is trimmed and restricted to 40
    Name = Trim$(Name)
    Name = Mid$(Name, 1, 40)


    If Len(Name) = 0 Then
        Exit Function
    End If
    'The key is made up of two parts, both 4
    '     digit hex values
    'Part1 is quite tricky


    For i = 1 To Len(Name)
        ch = Asc(Mid$(Name, i, 1)) * &H100
        'Run through the calculations 8 times fo
        '     r each character


        For n = 1 To 8
            'Do different things based on the result
            '     of this wierd if


            If (((ch Xor lPart1) Mod &H10000) And &H8000&) = 0 Then
                'Bit shift 1 bit left
                lPart1 = lPart1 And &HFFFFFFF
                lPart1 = lPart1 * 2
            Else
                'Bit shift 1 bit left
                lPart1 = lPart1 And &HFFFFFFF
                lPart1 = lPart1 * 2
                'Xor with the magic number
                lPart1 = lPart1 Xor &H1021&
            End If
            ch = ch * 2
        Next n
    Next i
    'Add a bit for luck
    lPart1 = lPart1 + &H63&
    'Only want 4 digits
    lPart1 = lPart1 Mod &H10000
    'Part2 is very simple


    For i = 1 To Len(Name)
        lPart2 = lPart2 + (Asc(Mid$(Name, i, 1)) * (i - 1))
    Next i
    'Build up from 2 parts (making sure each
    '     part is 4 digits)
    GenerateKey = String$(4 - Len(CStr(Hex(lPart1))), "0") _
    & CStr(Hex(lPart1)) _
    & String$(4 - Len(CStr(Hex(lPart2))), "0") _
    & CStr(Hex(lPart2))
End Function

Private Sub Command2_Click()
Form2.Show vbModal, Me
End Sub

Private Sub Label3_Click()
Shell ("start iexplore.exe"), vbHide
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Text2.Text = GenerateKey(Text1)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
Text2.Text = GenerateKey(Text1)
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
Text2.Text = GenerateKey(Text2)
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
Text2.Text = GenerateKey(Text2)
End Sub
Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
MsgBox " WinZip Key Generator v.7x to 8x.", vbOKOnly
End Sub
