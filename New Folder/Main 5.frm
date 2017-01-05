VERSION 5.00
Begin VB.Form Main 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   ForeColor       =   &H0000FFFF&
   Icon            =   "Main 5.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton Begin_cmd 
      Caption         =   "GO"
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   7320
      Width           =   3255
   End
   Begin VB.PictureBox Disp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      FillColor       =   &H00FF0000&
      Height          =   7095
      Left            =   0
      MousePointer    =   99  'Custom
      ScaleHeight     =   7035
      ScaleWidth      =   7035
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.Image HappyPic 
         Height          =   480
         Left            =   1680
         Picture         =   "Main 5.frx":0442
         Top             =   2400
         Width           =   480
      End
   End
   Begin VB.Label BlackLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   7800
      Width           =   975
   End
   Begin VB.Label WhiteLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   7320
      Width           =   975
   End
   Begin VB.Image WhitePic 
      Height          =   480
      Left            =   4200
      Picture         =   "Main 5.frx":0884
      Top             =   7320
      Width           =   480
   End
   Begin VB.Image BlackPic 
      Height          =   480
      Left            =   4200
      Picture         =   "Main 5.frx":0CC6
      Top             =   7800
      Width           =   480
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this is the baord and environment in which algorithms and people play

Option Explicit


Private Sub Begin_cmd_Click()
setup
End Sub

Sub setup()

Randomize

Main.Disp.Enabled = True
Main.Disp.Scale (0, 8)-(8, 0)

whoseTurn = white
notTurn = black
Main.Disp.MouseIcon = Main.WhitePic.Picture

'original setup
board(3, 3) = black
board(3, 4) = white
board(4, 3) = white
board(4, 4) = black
DrawChips

End Sub


Private Sub Command1_Click()
Joe.load_weights
Joe.neural_rater
End Sub

Private Sub Disp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xd As Integer
Dim yd As Integer
Dim you As Integer
xd = Int(X)
yd = Int(Y)
you = clicked_board(xd, yd)

End Sub

Function clicked_board(xd As Integer, yd As Integer)

If board(xd, yd) = 0 Then
        board(xd, yd) = whoseTurn
        If CalculateColourChanges(xd, yd, whoseTurn) = False Then
            board(xd, yd) = 0
        Else
        movesmade = movesmade + 1
        DrawChips
        whoseTurn = whoseTurn * -1
        End If
End If

white_score = find_score(white)
black_score = find_score(black)
Main.WhiteLabel.Caption = white_score
Main.BlackLabel.Caption = black_score

End Function

Sub DrawChips()

Dim xd As Integer
Dim yd As Integer
Dim i As Integer

Main.Disp.FillStyle = 0
Main.Disp.Cls

For i = 1 To 8
    Main.Disp.Line (0, i)-(8, i)
    Main.Disp.Line (i, 0)-(i, 8)
Next i

For xd = 0 To 7
    For yd = 0 To 7
            If board(xd, yd) = white Then
                Main.Disp.FillColor = QBColor(15)
                Main.Disp.Circle (xd + 0.5, yd + 0.5), 0.4
            Else
                If board(xd, yd) = black Then
                    Main.Disp.FillColor = QBColor(0)
                    Main.Disp.Circle (xd + 0.5, yd + 0.5), 0.4
                End If
            End If
    Next yd
Next xd

End Sub

Function CalculateColourChanges(xd As Integer, yd As Integer, colour As Integer)

Dim i As Integer
Dim flag As Boolean
Dim j As Integer
Dim m As Integer
Dim n As Integer
Dim numchanges As Integer
Dim move_made As Boolean
move_made = False

flag = True

For n = -1 To 1
    For m = -1 To 1
        If m = 0 And n = 0 Then
        Else
            i = 0
            j = 0
            flag = True
            Do Until flag = False
                i = i + 1
                If board(xd + (n * i), yd + (m * i)) <> 0 Then
                    flag = True
                    If board(xd + (n * i), yd + (m * i)) = colour Then
                        If i >= 2 Then
                            For j = 0 To i
                                board(xd + (n * j), yd + (m * j)) = colour
                            Next j
                            move_made = True
                            flag = False
                        Else
                            flag = False
                        End If
                    End If
                Else
                    flag = False
                End If
            Loop
        End If
    Next m
Next n

CalculateColourChanges = move_made

End Function


Function find_score(colour As Integer)
Dim xd As Integer
Dim yd As Integer
Dim temp_Score As Integer

For xd = 0 To 7
For yd = 0 To 7
    If board(xd, yd) = colour Then
        temp_Score = temp_Score + 1
    End If
        
Next yd
Next xd

find_score = temp_Score

End Function

Private Sub HappyPic_Click()

Dim xd As Integer
Dim yd As Integer
Dim you As Integer

xd = Int(Main.HappyPic.Left)
yd = Int(Main.HappyPic.Top)

you = clicked_board(xd, yd)
End Sub

