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

Global movesmade As Integer



Sub set_value_matrix()
Dim xd As Integer
Dim yd As Integer

Open "C:\My Documents\Benjamin\Othello game\matrix A.txt" For Input As 5

For xd = 0 To 7
    For yd = 0 To 7
        Input #5, value_matrix(xd, yd)
        Main.Disp.Circle (xd + 0.5, yd + 0.5), value_matrix(xd, yd) / 10
    Next yd
Next xd

End Sub
