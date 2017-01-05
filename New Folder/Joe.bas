Attribute VB_Name = "Joe"
Option Explicit
Global layer1(0 To 7, 0 To 7) As Integer
Global layer2(0 To 7, 0 To 7) As Integer

Sub neural_rater()
Dim xd As Integer
Dim yd As Integer
Dim xt As Integer
Dim yt As Integer



For xd = 0 To 7
    For yd = 0 To 7
        For xt = 0 To 7
            For yt = 0 To 7
                layer1(xd, yd) = board(xt, yt) * weight(xd, yd, xt, yt, 0)
                layer2(xd, yd) = layer2(xd, yd) + layer1(xd, yd)
            Next yt
        Next xt


        Main.Disp.Line (xd, yd)-(xd + 1, yd + 1), RGB((layer2(xd, yd) * 8 + 128), 100, 100), BF
        
    Next yd
Next xd
    
        
        
        
        

End Sub

Sub load_weights()

Dim xd As Integer
Dim yd As Integer
Dim xt As Integer
Dim yt As Integer
Dim i As Integer

Open "C:\My Documents\Benjamin\Othello game\weights.txt" For Input As 5

For xd = 0 To 7
    For yd = 0 To 7
        For xt = 0 To 7
            For yt = 0 To 7
                Input #5, weight(xd, yd, xt, yt, i)
            Next yt
        Next xt
    Next yd
Next xd

Close 5
End Sub


Sub save_weights()
Dim xd As Integer
Dim yd As Integer
Dim xt As Integer
Dim yt As Integer
Open "C:\My Documents\Benjamin\Othello game\weights.txt" For Output As 5

For xd = 0 To 7
    For yd = 0 To 7
        For xt = 0 To 7
            For yt = 0 To 7
                Print #5, Rnd 'weight(xd, yd, xt, yt)
            Next yt
        Next xt
    Next yd
Next xd

Close 5



End Sub
