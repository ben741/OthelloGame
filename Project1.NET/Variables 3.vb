Option Strict Off
Option Explicit On
Module Variables
	Public whoseTurn As Short
	
	Public chipX(63) As Short
	Public chipY(63) As Short
	'UPGRADE_ISSUE: Declaration type not supported: Array with lower bound less than zero. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1051"'
	Public chipColour(-1 To 8, -1 To 8) As Short
	Public value_matrix(7, 7) As Short
	'UPGRADE_ISSUE: Declaration type not supported: Array with lower bound less than zero. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1051"'
	Public new_board(-1 To 8, -1 To 8, 5) As Short
	Public counter As Short
	
	Public Const black As Short = 10
	Public Const white As Short = 1567
	'UPGRADE_ISSUE: Declaration type not supported: Array with lower bound less than zero. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1051"'
	Public occupied(-1 To 8, -1 To 8) As Boolean
	Public white_score As Short
	Public black_score As Short
	
	Public movesmade As Short
	
	
	
	Sub set_value_matrix()
		Dim xd As Short
		Dim yd As Short
		
		FileOpen(5, "C:\My Documents\Benjamin\Othello game\matrix A.txt", OpenMode.Input)
		
		For xd = 0 To 7
			For yd = 0 To 7
				Input(5, value_matrix(xd, yd))
				'UPGRADE_ISSUE: PictureBox method Disp.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
				Main_Renamed.DefInstance.Disp.Circle (xd + 0.5, yd + 0.5), value_matrix(xd, yd) / 10
			Next yd
		Next xd
		
	End Sub
End Module