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
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
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
      Picture         =   "Main.frx":0000
      Top             =   7320
      Width           =   480
   End
   Begin VB.Image BlackPic 
      Height          =   480
      Left            =   4200
      Picture         =   "Main.frx":0442
      Top             =   7800
      Width           =   480
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Begin_cmd_Click()

Dim i As Integer
Dim j As Integer
Main.Disp.Enabled = True
Main.Disp.Scale (0, 8)-(8, 0)

For i = 1 To 8
Main.Disp.Line (0, i)-(8, i)
Main.Disp.Line (i, 0)-(i, 8)

Next i

whoseTurn = white

Main.Disp.MouseIcon = Main.WhitePic.Picture
occupied(3, 3) = True
occupied(3, 4) = True
occupied(4, 3) = True
occupied(4, 4) = True

'original setup
chipColour(3, 3) = black
chipColour(3, 4) = white
chipColour(4, 3) = white
chipColour(4, 4) = black
DrawChips

End Sub


Private Sub Disp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim blaah As String
Dim xd As Integer
Dim yd As Integer

xd = Int(X)
yd = Int(Y)
If occupied(xd, yd) = False Then
    occupied(xd, yd) = True
    
    If whoseTurn = white Then
        chipColour(xd, yd) = white
        If CalculateColourChanges(xd, yd, white) = False Then
            occupied(xd, yd) = False
            chipColour(xd, yd) = 0
        Else
        DrawChips
        whoseTurn = black
        find_legal_moves (black)
        Main.Disp.MouseIcon = Main.BlackPic.Picture
        End If
    Else
        If whoseTurn = black Then
            chipColour(xd, yd) = black
            If CalculateColourChanges(xd, yd, black) = False Then
                occupied(xd, yd) = False
                chipColour(xd, yd) = 0
            Else
                DrawChips
                whoseTurn = white
                find_legal_moves (white)
                Main.Disp.MouseIcon = Main.WhitePic.Picture
            End If
        End If
    End If
End If

white_score = find_score(white)
black_score = find_score(black)
Main.WhiteLabel.Caption = white_score
Main.BlackLabel.Caption = black_score
End Sub

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
        If occupied(xd, yd) = True Then
            If chipColour(xd, yd) = white Then
                Main.Disp.FillColor = QBColor(15)
                Main.Disp.Circle (xd + 0.5, yd + 0.5), 0.4
            Else
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
                If occupied(xd + (n * i), yd + (m * i)) = True Then
                    flag = True
                    If chipColour(xd + (n * i), yd + (m * i)) = colour Then
                        If i >= 2 Then
                            For j = 0 To i
                                chipColour(xd + (n * j), yd + (m * j)) = colour
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

Sub find_legal_moves(colour As Integer)

Dim xd As Integer
Dim yd As Integer
Dim i As Integer
Dim flag As Boolean
Dim j As Integer
Dim m As Integer
Dim n As Integer
Dim numchanges As Integer
Dim move_made As Boolean
Dim legal(0 To 7, 0 To 7) As Boolean
Dim is_move As Boolean
is_move = False
Dim points(0 To 7, 0 To 7) As Integer
Dim best_move As Integer
move_made = False

flag = True

For xd = 0 To 7
For yd = 0 To 7
        points(xd, yd) = 1
        For n = -1 To 1
        For m = -1 To 1
            If m = 0 And n = 0 Then
            Else
                i = 0
                j = 0
                flag = True
                Do Until flag = False
                    i = i + 1
                    If occupied(xd + (n * i), yd + (m * i)) = True Then
                        flag = True
                        If chipColour(xd + (n * i), yd + (m * i)) = colour Then
                            If i >= 2 And occupied(xd, yd) = False Then
                                legal(xd, yd) = True
                                points(xd, yd) = points(xd, yd) + i - 1
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
        If legal(xd, yd) = True Then
            Main.Disp.FillStyle = 0
            If whoseTurn = white Then
                Main.Disp.FillColor = RGB(256 - points(xd, yd) * 10, 256 - points(xd, yd) * 10, 256 - points(xd, yd) * 10)
            Else
                Main.Disp.FillColor = RGB(points(xd, yd) * 10, points(xd, yd) * 10, points(xd, yd) * 10)
            End If
            Main.Disp.Circle (xd + 0.5, yd + 0.5), 0.04 * points(xd, yd)
            If points(xd, yd) > best_move Then
                best_move = points(xd, yd)
            Else
                If points(xd, yd) = best_move And Rnd > 0.5 Then
                    best_move = points(xd, yd)
                End If
            End If
            is_move = True
        End If
        
Next yd
Next xd
        
If is_move = False Then
    If whoseTurn = black Then
        whoseTurn = white
    Else
        If whoseTurn = white Then
            whoseTurn = black
        End If
    End If
End If

End Sub

Function find_score(colour As Integer)
Dim xd As Integer
Dim yd As Integer
Dim temp_Score As Integer

For xd = 0 To 7
For yd = 0 To 7
    If occupied(xd, yd) = True And chipColour(xd, yd) = colour Then
        temp_Score = temp_Score + 1
    End If
        
Next yd
Next xd

find_score = temp_Score

End Function
