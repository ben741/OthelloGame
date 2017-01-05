Attribute VB_Name = "Variables"
Global whoseTurn As Integer
Global notTurn As Integer

Global chipX(0 To 63) As Integer
Global chipY(0 To 63) As Integer
Global board(-1 To 8, -1 To 8) As Integer
Global value_matrix(0 To 7, 0 To 7) As Integer
Global new_board(-1 To 8, -1 To 8, 5) As Integer
Global counter As Integer

Global Const black = -1
Global Const white = 1
Global white_score As Integer
Global black_score As Integer
Global messaged(-1 To 7, -1 To 7) As Integer

Global movesmade As Integer
Global weight(0 To 7, 0 To 7, 0 To 7, 0 To 7, 0 To 15) As Single


