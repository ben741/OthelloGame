Attribute VB_Name = "Al"
' AI algorithm #1

Sub AI_player_move(colour As Integer)

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
Dim xb As Integer
Dim yb As Integer
Dim nomove As Integer
Dim bestx As Integer
Dim besty As Integer
Dim you As Integer
Dim temp As Integer

If whoseTurn = white Then
nomove = find_legal_moves3(black)
Else
nomove = find_legal_moves3(white)
End If
Dim best_move As Integer
best_move = -32000
move_made = False


If board(0, 0) Or board(0, 7) Or board(7, 0) Or board(7, 7) Then
    load_matrix ("Matrix b")
Else
    load_matrix ("Matrix a")
End If
messenger (whoseTurn)
flag = True

For xb = 0 To 7
    For yb = 0 To 7
        new_board(xb, yb, 0) = board(xb, yb)
        
    Next yb
Next xb

For xd = 0 To 7
For yd = 0 To 7
    For xb = 0 To 7
    For yb = 0 To 7
        new_board(xb, yb, 0) = board(xb, yb)
        
    Next yb
    Next xb
    
    new_board(xd, yd, 0) = colour
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
                    If board(xd + (n * i), yd + (m * i)) <> 0 Then
                        flag = True
                        If board(xd + (n * i), yd + (m * i)) = colour Then
                            If i >= 2 And board(xd, yd) = 0 Then
                                For j = 0 To i
                                    new_board(xd + (n * j), yd + (m * j), 0) = colour
                                Next j
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
            For xb = 0 To 7
            For yb = 0 To 7
                If board(xb, yb) = whoseTurn Then ''''''''''''''''''''''''''
                    points(xd, yd) = points(xd, yd) + value_matrix(xb, yb)
                End If
            Next yb
            Next xb
            
            If whoseTurn = black Then
                temp = find_legal_moves2(white)
                If temp = 32067 Then
                    points(xd, yd) = -100
                Else
                    points(xd, yd) = (points(xd, yd) - (temp - nomove)) + value_matrix(xd, yd) * 100
                End If
            Else
                temp = find_legal_moves2(black)
                If temp = 32067 Then
                    points(xd, yd) = -100
                Else
                    points(xd, yd) = (points(xd, yd) - (temp - nomove)) + value_matrix(xd, yd) * 100
                End If
            End If
            'If whoseTurn = black Then
            '    points(xd, yd) = (points(xd, yd) - (find_legal_moves2(white) - nomove)) * value_matrix(xd, yd)
            'Else
            '    points(xd, yd) = (points(xd, yd) - (find_legal_moves2(black) - nomove)) * value_matrix(xd, yd)
            'End If
        
        
            Main.Disp.FillStyle = 0
            If points(xd, yd) > 0 Then
                'Main.Disp.FillColor = RGB(256 - points(xd, yd) * 10, 256 - points(xd, yd) * 10, 256 - points(xd, yd) * 10)
                'Main.Disp.FillColor = RGB(10, 100, 10)
            Else
                'Main.Disp.FillColor = RGB(points(xd, yd) * 10, points(xd, yd) * 10, points(xd, yd) * 10)
                'Main.Disp.FillColor = RGB(100, 10, 10)
            End If
            'Main.Disp.Circle (xd + 0.5, yd + 0.5), 0.04 * Abs(points(xd, yd)) + 0.05
            If points(xd, yd) > best_move Then
                best_move = points(xd, yd)
                bestx = xd
                besty = yd

            Else
                If points(xd, yd) = best_move And Rnd > 0.5 Then
                    best_move = points(xd, yd)
                    bestx = xd
                    besty = yd
                End If
            End If
            is_move = True
        End If
        
Next yd
Next xd
        
If legal(0, 0) = True Then
    bestx = 0
    besty = 0
End If
If legal(0, 7) = True Then
    bestx = 0
    besty = 7
End If
If legal(7, 0) = True Then
    bestx = 7
    besty = 0
End If
If legal(7, 7) = True Then
    bestx = 7
    besty = 7
End If
Main.HappyPic.Left = bestx + 0.2
Main.HappyPic.Top = besty + 0.8
you = Main.clicked_board(bestx, besty)


If is_move = False Then
    Beep
    If whoseTurn = black Then
        whoseTurn = white
        Main.Disp.MouseIcon = Main.WhitePic.Picture
    Else
        If whoseTurn = white Then
            whoseTurn = black
            Main.Disp.MouseIcon = Main.BlackPic.Picture
        End If
    End If
End If

End Sub

Function find_legal_moves2(colour As Integer)

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
Dim points(0 To 7, 0 To 7) As Long
Dim best_move As Integer
best_move = -32000
Dim xb As Integer
Dim yb As Integer

move_made = False
Dim nomove As Integer
    If colour = white Then
    nomove = find_legal_moves3(black)
    Else
    nomove = find_legal_moves3(white)
    End If
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
                    If new_board(xd + (n * i), yd + (m * i), counter) <> 0 Then
                        flag = True
                        If new_board(xd + (n * i), yd + (m * i), counter) = colour Then
                            If i >= 2 And board(xd, yd) = 0 Then
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
        
        For xb = 0 To 7
        For yb = 0 To 7
            If new_board(xb, yb, 0) = colour Then
                points(xd, yd) = points(xd, yd) + value_matrix(xb, yb)
            End If
        Next yb
        Next xb
        
        points(xd, yd) = points(xd, yd) * value_matrix(xd, yd)
        If legal(xd, yd) = True Then
 '           If counter <= 1 Then
 '           counter = counter + 1
 '               If colour = black Then
 '                   points(xd, yd) = (points(xd, yd) - (find_legal_moves2(white) - nomove)) * value_matrix(xd, yd)
 '               Else
 '                   points(xd, yd) = (points(xd, yd) - (find_legal_moves2(black) - nomove)) * value_matrix(xd, yd)
 '               End If
 '           End If
 '           counter = counter - 1
            If legal(xd, yd) = True Then
            If points(xd, yd) > best_move Then
                best_move = points(xd, yd)
            Else
                If points(xd, yd) = best_move And Rnd > 0.5 Then
                    best_move = points(xd, yd)
                End If
            End If
            is_move = True
            End If
        End If
Next yd
Next xd
If legal(0, 0) = True Then
    best_move = 32067
End If
If legal(0, 7) = True Then
    best_move = 32067
End If
If legal(7, 0) = True Then
    best_move = 32067
End If
If legal(7, 7) = True Then
    best_move = 32067
End If
find_legal_moves2 = best_move
End Function
Function find_legal_moves3(colour As Integer)

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
Dim xb As Integer
Dim yb As Integer

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
                    If board(xd + (n * i), yd + (m * i)) <> 0 Then
                        flag = True
                        If board(xd + (n * i), yd + (m * i)) = colour Then
                            If i >= 2 And board(xd, yd) = 0 Then
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
        
        For xb = 0 To 7
        For yb = 0 To 7
            If new_board(xb, yb, 0) = colour Then
                points(xd, yd) = points(xd, yd) + value_matrix(xb, yb)
            End If
        Next yb
        Next xb
        
        points(xd, yd) = points(xd, yd) * value_matrix(xd, yd)
        
        If legal(xd, yd) = True Then
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

If legal(0, 0) = True Then
    best_move = 320
End If
If legal(0, 7) = True Then
    best_move = 320
End If
If legal(7, 0) = True Then
    best_move = 320
End If
If legal(7, 7) = True Then
    best_move = 320
End If
find_legal_moves3 = best_move
End Function
Sub messenger(colour As Integer)


Dim xd As Integer
Dim yd As Integer
Dim cornerX As Integer
Dim cornerY As Integer
Dim flag As Boolean
Dim m As Integer
Dim n As Integer
Dim current As Integer
current = 9
Erase messaged

If board(0, 0) = colour Then
    messaged(0, 0) = current - 1
End If
If board(0, 7) = colour Then
    messaged(0, 7) = current - 1
End If
If board(7, 0) = colour Then
    messaged(7, 0) = current - 1
End If
If board(7, 7) = colour Then
    messaged(7, 7) = current - 1
End If

flag = True
Do Until flag = False
    current = current - 1
    For yd = 0 To 7
    For xd = 0 To 7
        flag = False

        If messaged(xd, yd) = current Then
            For n = -1 To 1
            For m = -1 To 1
                If board(xd + n, yd + m) <> 0 And messaged(xd + n, yd + m) = 0 Then
                    messaged(xd + n, yd + m) = current - 1
                    flag = True
                    Beep
                End If
                If board(xd + n, yd + m) <> 0 And xd + n > -1 And xd + n < 8 And yd + m > -1 And yd + m < 8 Then
                    value_matrix(xd + n, yd + m) = value_matrix(xd + n, yd + m) + current
                    Main.Disp.Circle (xd + n + 0.5, yd + n + 0.5), 0.1
                    flag = True
                End If
            Next m
            Next n
        End If

    Next xd
    Next yd
Loop

End Sub


Sub load_matrix(file_name As String)
Dim xd As Integer
Dim yd As Integer

Open "C:\My Documents\Benjamin\Othello game\" & file_name & ".txt" For Input As 5

For xd = 0 To 7
    For yd = 0 To 7
        Input #5, value_matrix(xd, yd)
        'Main.Disp.Circle (xd + 0.5, yd + 0.5), value_matrix(xd, yd) / 10
    Next yd
Next xd

Close 5

End Sub

