Attribute VB_Name = "Variables"
Global whoseTurn As Integer

Global chipX(0 To 63) As Integer
Global chipY(0 To 63) As Integer
Global chipColour(-1 To 8, -1 To 8) As Integer
Global value_matrix(0 To 7, 0 To 7) As Integer
Global new_board(-1 To 8, -1 To 8, 5) As Integer
Global counter As Integer

Global Const black = 10
Global Const white = 1567
Global occupied(-1 To 8, -1 To 8) As Boolean
Global white_score As Integer
Global black_score As Integer
Global messaged(-1 To 8, -1 To 8) As Integer

Global movesmade As Integer



Sub set_value_matrix()

'column 1
value_matrix(0, 0) = 10
value_matrix(0, 1) = 0
value_matrix(0, 2) = 2
value_matrix(0, 3) = 2
value_matrix(0, 4) = 2
value_matrix(0, 5) = 2
value_matrix(0, 6) = 0
value_matrix(0, 7) = 10

'column 2
value_matrix(1, 0) = 0
value_matrix(1, 1) = 0
value_matrix(1, 2) = 1
value_matrix(1, 3) = 1
value_matrix(1, 4) = 1
value_matrix(1, 5) = 1
value_matrix(1, 6) = 0
value_matrix(1, 7) = 0

'column 3
value_matrix(2, 0) = 2
value_matrix(2, 1) = 1
value_matrix(2, 2) = 2
value_matrix(2, 3) = 2
value_matrix(2, 4) = 2
value_matrix(2, 5) = 2
value_matrix(2, 6) = 1
value_matrix(2, 7) = 2

'column 4
value_matrix(3, 0) = 2
value_matrix(3, 1) = 1
value_matrix(3, 2) = 2
value_matrix(3, 3) = 1
value_matrix(3, 4) = 1
value_matrix(3, 5) = 2
value_matrix(3, 6) = 1
value_matrix(3, 7) = 2

'column 5
value_matrix(4, 0) = 2
value_matrix(4, 1) = 1
value_matrix(4, 2) = 2
value_matrix(4, 3) = 1
value_matrix(4, 4) = 1
value_matrix(4, 5) = 2
value_matrix(4, 6) = 1
value_matrix(4, 7) = 2

'column 6
value_matrix(5, 0) = 2
value_matrix(5, 1) = 1
value_matrix(5, 2) = 2
value_matrix(5, 3) = 2
value_matrix(5, 4) = 2
value_matrix(5, 5) = 2
value_matrix(5, 6) = 1
value_matrix(5, 7) = 2

'column 7
value_matrix(6, 0) = 0
value_matrix(6, 1) = 0
value_matrix(6, 2) = 1
value_matrix(6, 3) = 1
value_matrix(6, 4) = 1
value_matrix(6, 5) = 1
value_matrix(6, 6) = 0
value_matrix(6, 7) = 0

'column 8
value_matrix(7, 0) = 10
value_matrix(7, 1) = 0
value_matrix(7, 2) = 2
value_matrix(7, 3) = 2
value_matrix(7, 4) = 2
value_matrix(7, 5) = 2
value_matrix(7, 6) = 0
value_matrix(7, 7) = 10

End Sub
