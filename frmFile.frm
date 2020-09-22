VERSION 5.00
Begin VB.Form frmFile 
   Caption         =   "Select a file to play with"
   ClientHeight    =   3060
   ClientLeft      =   2505
   ClientTop       =   2190
   ClientWidth     =   3780
   LinkTopic       =   "Form2"
   ScaleHeight     =   3060
   ScaleWidth      =   3780
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play the game"
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   2040
      TabIndex        =   7
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select a file to play with:"
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.OptionButton optOther 
         Caption         =   "Other (user selects file)"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   2775
      End
      Begin VB.OptionButton optTransportation 
         Caption         =   "Transportation"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   2655
      End
      Begin VB.OptionButton optEnglish 
         Caption         =   "English"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   2775
      End
      Begin VB.OptionButton optMath 
         Caption         =   "Math"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   2535
      End
      Begin VB.OptionButton optOriginal 
         Caption         =   "Original File"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPlay_Click()
    If optOriginal.Value = True Then
        strFile = "Library.txt"
        Unload Me
        Form1.Show
    ElseIf optMath.Value = True Then
        strFile = "Math.txt"
        Unload Me
        Form1.Show
    ElseIf optEnglish.Value = True Then
        strFile = "English.txt"
        Unload Me
        Form1.Show
    ElseIf optTransportation.Value = True Then
        strFile = "Transportation.txt"
        Unload Me
        Form1.Show
    Else: MsgBox "You need to select a file to continue"
    End If
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub optOther_Click()
    frmFileChoose.Show
    Unload Me
End Sub
