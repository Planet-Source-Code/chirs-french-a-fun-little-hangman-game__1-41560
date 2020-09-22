VERSION 5.00
Begin VB.Form frmFileChoose 
   Caption         =   "Choose a file to play with"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4470
   LinkTopic       =   "Form2"
   ScaleHeight     =   3210
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChoose 
      Caption         =   "Select with file"
      Height          =   855
      Left            =   2400
      TabIndex        =   1
      Top             =   2280
      Width           =   2055
   End
   Begin VB.FileListBox fleFile 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmFileChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChoose_Click()
    strFile = fleFile.FileName
    Unload Me
    Form1.Show
End Sub

Private Sub fleFile_DblClick()
    cmdChoose_Click
End Sub

Private Sub Form_Load()
    fleFile.Path = App.Path
End Sub
