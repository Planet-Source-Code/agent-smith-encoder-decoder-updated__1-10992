VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Encoder - Decoder"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain2.frx":0000
   ScaleHeight     =   6225
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "COPY /\"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   2760
      Width           =   3465
   End
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmMain2.frx":0082
      Top             =   0
      Width           =   7095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CONVERT"
      Height          =   495
      Left            =   15
      TabIndex        =   1
      Top             =   2760
      Width           =   3570
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   7035
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6000
      Width           =   7095
      Begin VB.Image Image1 
         Height          =   240
         Left            =   -15
         Picture         =   "frmMain2.frx":1D8A
         Top             =   0
         Width           =   7275
      End
   End
   Begin VB.TextBox Text2 
      Height          =   2775
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   3240
      Width           =   7095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Encoder Decoder by James Gohl a.k.a. wasp53x - (wasp53x@hotmail.com)
'Thanx to Adolf Gerold for pointing out a few embarassing mistakes!

'Function below is needed
Private Function Convert(cString As String) As String
    
    For cCode = 1 To Len(cString)
        Conv = Conv + (100 / Len(cString)) '<<<Dont want the status bar?, then remove this code
        Image1.Width = (Picture1.Width / Len(cString)) * Conv * (Len(cString) / 100) '<<<Dont want the status bar?, then remove this code
        Convert = Convert + Chr(255 - Asc(Mid(cString, CInt(cCode), 1)))
    Next cCode
    Form_Load

End Function

'Code for using the functions:
Private Sub Command1_Click()
    Text2 = Convert(Text1)
End Sub

Private Sub Command2_Click()
    Text1 = Text2
    Text2 = ""
End Sub

Private Sub Form_Load()
    Image1.Width = 0 '<<<Dont want the status bar?, then remove this code
End Sub
