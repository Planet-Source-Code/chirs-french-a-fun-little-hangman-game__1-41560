VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   1395
   ClientTop       =   1500
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   6045
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   3840
      TabIndex        =   5
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdGuess 
      Caption         =   "Guess"
      Height          =   615
      Left            =   2040
      TabIndex        =   4
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txtLetter 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label lblCheat 
      Height          =   615
      Left            =   5520
      TabIndex        =   20
      Top             =   3120
      Width           =   495
   End
   Begin VB.Line lne9 
      Visible         =   0   'False
      X1              =   720
      X2              =   600
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line lne8 
      Visible         =   0   'False
      X1              =   360
      X2              =   240
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line lne7 
      Visible         =   0   'False
      X1              =   360
      X2              =   600
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line lne6 
      Visible         =   0   'False
      X1              =   840
      X2              =   480
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line lne5 
      Visible         =   0   'False
      X1              =   120
      X2              =   480
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line lne4 
      Visible         =   0   'False
      X1              =   840
      X2              =   480
      Y1              =   2520
      Y2              =   2280
   End
   Begin VB.Line lne3 
      Visible         =   0   'False
      X1              =   120
      X2              =   480
      Y1              =   2520
      Y2              =   2280
   End
   Begin VB.Line lne2 
      Visible         =   0   'False
      X1              =   480
      X2              =   480
      Y1              =   1200
      Y2              =   2280
   End
   Begin VB.Shape shp1 
      Height          =   615
      Left            =   120
      Shape           =   3  'Circle
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   1920
      TabIndex        =   18
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblWord 
      Height          =   255
      Index           =   11
      Left            =   5640
      TabIndex        =   17
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblWord 
      Height          =   255
      Index           =   10
      Left            =   5280
      TabIndex        =   16
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblWord 
      Height          =   255
      Index           =   9
      Left            =   4920
      TabIndex        =   15
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblWord 
      Height          =   255
      Index           =   8
      Left            =   4560
      TabIndex        =   14
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblWord 
      Height          =   255
      Index           =   7
      Left            =   4200
      TabIndex        =   13
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblWord 
      Height          =   255
      Index           =   6
      Left            =   3840
      TabIndex        =   12
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblWord 
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   11
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblWord 
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   10
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblWord 
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   9
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblWord 
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   8
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblWord 
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   7
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblWord 
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   6
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblLetters 
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Letters guessed"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter guess:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Line Line3 
      X1              =   480
      X2              =   480
      Y1              =   600
      Y2              =   240
   End
   Begin VB.Line Line2 
      X1              =   480
      X2              =   1440
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line1 
      X1              =   1440
      X2              =   1440
      Y1              =   240
      Y2              =   2640
   End
   Begin VB.Label lne10 
      Caption         =   "N-O-O-O-O-O-O"
      Height          =   255
      Left            =   840
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strLetter As String
Dim i, intWordLNum As Integer
Dim strRawWord(20) As String
Dim intLen As Integer
Dim strWord(20) As String
Dim booguesed As Boolean
Dim intBody As Integer
Dim intGuessed As Integer
Dim inta As Integer

Private Sub cmdGuess_Click()
    Dim intRepeat As Integer
    strLetter = txtLetter.Text
    If strLetter = "cats and dogs" Then
        MsgBox strRawWord(inta)
        Exit Sub
    End If
    For i = 0 To 20
        If strLetter = strWord(i) Then
            intWordLNum = i
        Else
            intRepeat = intRepeat + 1
        End If
    Next i
    If intRepeat = 21 And booguesed = False Then
        Body
    Else
        Word
    End If
    lblLetters.Caption = lblLetters.Caption & strLetter & ","
    Check
End Sub

Sub Body()
    'MsgBox "Your guess is incorrect"
    intBody = intBody + 1
    Select Case intBody
    Case 1
        shp1.Visible = True
    Case 2
        lne2.Visible = True
    Case 3
        lne3.Visible = True
    Case 4
        lne4.Visible = True
    Case 5
        lne5.Visible = True
    Case 6
        lne6.Visible = True
    Case 7
        lne7.Visible = True
    Case 8
        lne8.Visible = True
    Case 9
        lne9.Visible = True
    Case 10
        lne10.Visible = True
        MsgBox "You have failed and died"
        Die
    End Select
    txtLetter.SetFocus
    txtLetter.SelStart = 0
    txtLetter.SelLength = Len(txtLetter.Text)
End Sub

Sub Check()
    Dim intCheck As Integer
    Dim booLetters As Boolean
    
    booLetters = True
    For i = 0 To Len(strRawWord(inta))
        If lblWord(i).Caption = "_" Then booLetters = False 'intCheck = intCheck + 1
    Next i
    'If intCheck = Len(strRawWord(inta)) Then
    ' If intCheck = 0 Then
    
    
    If booLetters Then
        MsgBox "You have won" & vbCrLf & "Good job"
        Unload Me
        frmFile.Show
    End If
End Sub

Sub Die()
    MsgBox "Your fatal word was" & vbCrLf & strRawWord(inta)
    Unload Me
    frmFile.Show
    Exit Sub
End Sub

Sub Word()
    If intWordLNum < 10 Then
        lblWord(intWordLNum).Caption = strWord(intWordLNum)
    End If
    strWord(intWordLNum) = ""
    'For i = 0 To 11
        'lblWord(i).Caption = strWord(i)
    'Next i
    For i = 1 To 10
        If strLetter = strWord(i) Then
            booguesed = True
            cmdGuess_Click
        End If
    Next i
    txtLetter.SetFocus
    txtLetter.SelStart = 0
    txtLetter.SelLength = Len(txtLetter.Text)
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim intFile As Integer
    Dim i2 As Integer
    Dim strRawWordTemp As String
    'If Not File_Exists(strFile) Then
        'Unload Me
        'frmFile.Show
        'Exit Sub
    'End If
    GoTo ByPass
HelpMe:
    MsgBox "File unavailable"
    frmFileChoose.Show
ByPass:
    Randomize
    inta = Int(Rnd * 20 + 1)
    i = 0
    intFile = FreeFile
    'MsgBox strFile
    'MsgBox App.Path & vbCr & Trim(App.Path) & "\Library.txt"
    On Error GoTo HelpMe
    Open (App.Path) & "/" & strFile For Input As #intFile
        For i = 0 To inta
            Line Input #intFile, strRawWord(i)
            Label3.Caption = strRawWord(i)
        Next i
    Close #intFile
    For i = 1 To Len(strRawWord(inta))
        lblWord(i).Caption = "_"
    Next i
    intLen = Len(strRawWord(inta))
    strRawWordTemp = strRawWord(inta)
    For i = 1 To intLen
        strWord(i) = Left(strRawWordTemp, 1)
        strRawWordTemp = Right(strRawWordTemp, (Len(strRawWordTemp) - 1))
    Next i
    'Label3.Caption = strRawWord(inta)

End Sub

Private Sub lblCheat_Click()
    MsgBox "The word is " & vbCrLf & strRawWord(inta)
End Sub

Private Sub txtLetter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then: cmdGuess_Click
End Sub

Function File_Exists(ByVal PathName As String, Optional Directory As Boolean) As Boolean
    If PathName <> "" Then
    'if file exists, then true, otherwise, false
        If IsMissing(Directory) Or Directory = False Then
            File_Exists = (Dir$(PathName) <> "")
        Else
            File_Exists = (Dir$(PathName, vbDirectory) <> "")
        End If
    End If
End Function

