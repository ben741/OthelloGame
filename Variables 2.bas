Attribute VB_Name = "Variables"
Global whoseTurn As Integer

Global chipX(0 To 63) As Integer
Global chipY(0 To 63) As Integer
Global chipColour(-1 To 8, -1 To 8) As Integer
Global value_matrix(0 To 7, 0 To 7) As Integer
Global new_board(-1 To 8, -1 To 8) As Integer

Global Const black = 10
Global Const white = 1567
Global occupied(-1 To 8, -1 To 8) As Boolean
Global white_score As Integer
Global black_score As Integer

Sub set_value_matrix()
Dim xd As Integer
Dim yd As Integer

For xd = 0 To 7
    For yd = 0 To 7
        value_matrix(xd, yd) = 1
    
    Next yd
Next xd

value_matrix(0, 0) = 5
value_matrix(0, 7) = 5
value_matrix(7, 0) = 5
value_matrix(7, 7) = 5
value_matrix(1, 1) = 0
value_matrix(1, 6) = 0
value_matrix(6, 1) = 0
value_matrix(6, 6) = 0

End Sub
