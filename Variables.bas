Attribute VB_Name = "Variables"
Global whoseTurn As Integer

Global chipX(0 To 63) As Integer
Global chipY(0 To 63) As Integer
Global chipColour(-1 To 8, -1 To 8) As Integer


Global Const black = 10
Global Const white = 1567
Global occupied(-1 To 8, -1 To 8) As Boolean
Global white_score As Integer
Global black_score As Integer
