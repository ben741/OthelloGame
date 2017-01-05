Option Strict Off
Option Explicit On
Friend Class Main_Renamed
	Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'For the start-up form, the first instance created is the default instance.
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents Begin_cmd As System.Windows.Forms.Button
	Public WithEvents HappyPic As System.Windows.Forms.PictureBox
	Public WithEvents Disp As System.Windows.Forms.Panel
	Public WithEvents BlackLabel As System.Windows.Forms.Label
	Public WithEvents WhiteLabel As System.Windows.Forms.Label
	Public WithEvents WhitePic As System.Windows.Forms.PictureBox
	Public WithEvents BlackPic As System.Windows.Forms.PictureBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Main_Renamed))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.Begin_cmd = New System.Windows.Forms.Button
		Me.Disp = New System.Windows.Forms.Panel
		Me.HappyPic = New System.Windows.Forms.PictureBox
		Me.BlackLabel = New System.Windows.Forms.Label
		Me.WhiteLabel = New System.Windows.Forms.Label
		Me.WhitePic = New System.Windows.Forms.PictureBox
		Me.BlackPic = New System.Windows.Forms.PictureBox
		Me.BackColor = System.Drawing.SystemColors.Menu
		Me.Text = "Form1"
		Me.ClientSize = New System.Drawing.Size(476, 600)
		Me.Location = New System.Drawing.Point(4, 23)
		Me.ForeColor = System.Drawing.Color.Yellow
		Me.Icon = CType(resources.GetObject("Main.Icon"), System.Drawing.Icon)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "Main_Renamed"
		Me.Begin_cmd.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Begin_cmd.Text = "GO"
		Me.Begin_cmd.Size = New System.Drawing.Size(217, 81)
		Me.Begin_cmd.Location = New System.Drawing.Point(8, 488)
		Me.Begin_cmd.TabIndex = 3
		Me.Begin_cmd.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Begin_cmd.BackColor = System.Drawing.SystemColors.Control
		Me.Begin_cmd.CausesValidation = True
		Me.Begin_cmd.Enabled = True
		Me.Begin_cmd.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Begin_cmd.Cursor = System.Windows.Forms.Cursors.Default
		Me.Begin_cmd.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Begin_cmd.TabStop = True
		Me.Begin_cmd.Name = "Begin_cmd"
		Me.Disp.BackColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.Disp.Enabled = False
		Me.Disp.Size = New System.Drawing.Size(473, 473)
		Me.Disp.Location = New System.Drawing.Point(0, 0)
		Me.Disp.TabIndex = 0
		Me.Disp.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Disp.Dock = System.Windows.Forms.DockStyle.None
		Me.Disp.CausesValidation = True
		Me.Disp.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Disp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Disp.TabStop = True
		Me.Disp.Visible = True
		Me.Disp.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Disp.Name = "Disp"
		Me.HappyPic.Size = New System.Drawing.Size(32, 32)
		Me.HappyPic.Location = New System.Drawing.Point(112, 160)
		Me.HappyPic.Image = CType(resources.GetObject("HappyPic.Image"), System.Drawing.Image)
		Me.HappyPic.Enabled = True
		Me.HappyPic.Cursor = System.Windows.Forms.Cursors.Default
		Me.HappyPic.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.HappyPic.Visible = True
		Me.HappyPic.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.HappyPic.Name = "HappyPic"
		Me.BlackLabel.Text = "0"
		Me.BlackLabel.Font = New System.Drawing.Font("Arial", 18!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.BlackLabel.Size = New System.Drawing.Size(65, 33)
		Me.BlackLabel.Location = New System.Drawing.Point(320, 520)
		Me.BlackLabel.TabIndex = 2
		Me.BlackLabel.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.BlackLabel.BackColor = System.Drawing.Color.Transparent
		Me.BlackLabel.Enabled = True
		Me.BlackLabel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.BlackLabel.Cursor = System.Windows.Forms.Cursors.Default
		Me.BlackLabel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.BlackLabel.UseMnemonic = True
		Me.BlackLabel.Visible = True
		Me.BlackLabel.AutoSize = False
		Me.BlackLabel.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.BlackLabel.Name = "BlackLabel"
		Me.WhiteLabel.Text = "0"
		Me.WhiteLabel.Font = New System.Drawing.Font("Arial", 18!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.WhiteLabel.Size = New System.Drawing.Size(65, 33)
		Me.WhiteLabel.Location = New System.Drawing.Point(320, 488)
		Me.WhiteLabel.TabIndex = 1
		Me.WhiteLabel.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.WhiteLabel.BackColor = System.Drawing.Color.Transparent
		Me.WhiteLabel.Enabled = True
		Me.WhiteLabel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.WhiteLabel.Cursor = System.Windows.Forms.Cursors.Default
		Me.WhiteLabel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.WhiteLabel.UseMnemonic = True
		Me.WhiteLabel.Visible = True
		Me.WhiteLabel.AutoSize = False
		Me.WhiteLabel.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.WhiteLabel.Name = "WhiteLabel"
		Me.WhitePic.Size = New System.Drawing.Size(32, 32)
		Me.WhitePic.Location = New System.Drawing.Point(280, 488)
		Me.WhitePic.Image = CType(resources.GetObject("WhitePic.Image"), System.Drawing.Image)
		Me.WhitePic.Enabled = True
		Me.WhitePic.Cursor = System.Windows.Forms.Cursors.Default
		Me.WhitePic.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.WhitePic.Visible = True
		Me.WhitePic.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.WhitePic.Name = "WhitePic"
		Me.BlackPic.Size = New System.Drawing.Size(32, 32)
		Me.BlackPic.Location = New System.Drawing.Point(280, 520)
		Me.BlackPic.Image = CType(resources.GetObject("BlackPic.Image"), System.Drawing.Image)
		Me.BlackPic.Enabled = True
		Me.BlackPic.Cursor = System.Windows.Forms.Cursors.Default
		Me.BlackPic.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.BlackPic.Visible = True
		Me.BlackPic.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.BlackPic.Name = "BlackPic"
		Me.Controls.Add(Begin_cmd)
		Me.Controls.Add(Disp)
		Me.Controls.Add(BlackLabel)
		Me.Controls.Add(WhiteLabel)
		Me.Controls.Add(WhitePic)
		Me.Controls.Add(BlackPic)
		Me.Disp.Controls.Add(HappyPic)
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As Main_Renamed
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As Main_Renamed
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New Main_Renamed()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	
	Private Sub Begin_cmd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Begin_cmd.Click
		Randomize()
		Dim i As Short
		Dim j As Short
		Main_Renamed.DefInstance.Disp.Enabled = True
		'UPGRADE_ISSUE: PictureBox method Disp.Scale was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		Main_Renamed.DefInstance.Disp.Scale (0, 8) - (8, 0)
		
		For i = 1 To 8
			'UPGRADE_ISSUE: PictureBox method Disp.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Main.Disp.Line (0, i) - (8, i)
			'UPGRADE_ISSUE: PictureBox method Disp.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Main.Disp.Line (i, 0) - (i, 8)
			
		Next i
		
		whoseTurn = white
		
		'UPGRADE_ISSUE: PictureBox property Disp.MouseIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		Main_Renamed.DefInstance.Disp.MouseIcon = Main_Renamed.DefInstance.WhitePic.Image
		occupied(3, 3) = True
		occupied(3, 4) = True
		occupied(4, 3) = True
		occupied(4, 4) = True
		
		'original setup
		chipColour(3, 3) = black
		chipColour(3, 4) = white
		chipColour(4, 3) = white
		chipColour(4, 4) = black
		DrawChips()
		
		Variables.set_value_matrix()
		movesmade = 4
	End Sub
	
	
	Private Sub Disp_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles Disp.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim xd As Short
		Dim yd As Short
		Dim you As Short
		xd = Int(X)
		yd = Int(Y)
		'UPGRADE_WARNING: Couldn't resolve default property of object clicked_board(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		you = clicked_board(xd, yd)
		
	End Sub
	
	Function clicked_board(ByRef xd As Short, ByRef yd As Short) As Object
		
		Dim blaah As String
		Static end_game As Short
		
		If occupied(xd, yd) = False Then
			occupied(xd, yd) = True
			
			If whoseTurn = white Then
				chipColour(xd, yd) = white
				'UPGRADE_WARNING: Couldn't resolve default property of object CalculateColourChanges(xd, yd, white). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				If CalculateColourChanges(xd, yd, white) = False Then
					occupied(xd, yd) = False
					chipColour(xd, yd) = 0
				Else
					movesmade = movesmade + 1
					DrawChips()
					whoseTurn = black
					find_legal_moves((black))
					'UPGRADE_ISSUE: PictureBox property Disp.MouseIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
					Main_Renamed.DefInstance.Disp.MouseIcon = Main_Renamed.DefInstance.BlackPic.Image
				End If
			Else
				If whoseTurn = black Then
					chipColour(xd, yd) = black
					'UPGRADE_WARNING: Couldn't resolve default property of object CalculateColourChanges(xd, yd, black). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					If CalculateColourChanges(xd, yd, black) = False Then
						occupied(xd, yd) = False
						chipColour(xd, yd) = 0
					Else
						movesmade = movesmade + 1
						DrawChips()
						whoseTurn = white
						'find_legal_moves (white)
						'UPGRADE_ISSUE: PictureBox property Disp.MouseIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
						Main_Renamed.DefInstance.Disp.MouseIcon = Main_Renamed.DefInstance.WhitePic.Image
					End If
				End If
			End If
		End If
		
		If occupied(0, 0) Or occupied(0, 7) Or occupied(7, 0) Or occupied(7, 7) Then
			If end_game = 1 Then
				load_new_matrix()
				end_game = 1
			End If
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object find_score(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		white_score = find_score(white)
		'UPGRADE_WARNING: Couldn't resolve default property of object find_score(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		black_score = find_score(black)
		Main_Renamed.DefInstance.WhiteLabel.Text = CStr(white_score)
		Main_Renamed.DefInstance.BlackLabel.Text = CStr(black_score)
	End Function
	
	Sub DrawChips()
		Dim xd As Short
		Dim yd As Short
		Dim i As Short
		
		'UPGRADE_ISSUE: PictureBox property Disp.FillStyle was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		Main_Renamed.DefInstance.Disp.FillStyle = 0
		'UPGRADE_ISSUE: PictureBox method Disp.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		Main.Disp.Cls()
		
		For i = 1 To 8
			'UPGRADE_ISSUE: PictureBox method Disp.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Main.Disp.Line (0, i) - (8, i)
			'UPGRADE_ISSUE: PictureBox method Disp.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Main.Disp.Line (i, 0) - (i, 8)
			
		Next i
		
		For xd = 0 To 7
			For yd = 0 To 7
				If occupied(xd, yd) = True Then
					If chipColour(xd, yd) = white Then
						'UPGRADE_ISSUE: PictureBox property Disp.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
						Main_Renamed.DefInstance.Disp.FillColor = QBColor(15)
						'UPGRADE_ISSUE: PictureBox method Disp.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
						Main_Renamed.DefInstance.Disp.Circle (xd + 0.5, yd + 0.5), 0.4
					Else
						'UPGRADE_ISSUE: PictureBox property Disp.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
						Main_Renamed.DefInstance.Disp.FillColor = QBColor(0)
						'UPGRADE_ISSUE: PictureBox method Disp.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
						Main_Renamed.DefInstance.Disp.Circle (xd + 0.5, yd + 0.5), 0.4
					End If
					
				End If
			Next yd
		Next xd
		
	End Sub
	
	Function CalculateColourChanges(ByRef xd As Short, ByRef yd As Short, ByRef colour As Short) As Object
		
		Dim i As Short
		Dim flag As Boolean
		Dim j As Short
		Dim m As Short
		Dim n As Short
		Dim numchanges As Short
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
		
		'UPGRADE_WARNING: Couldn't resolve default property of object CalculateColourChanges. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		CalculateColourChanges = move_made
		
	End Function
	
	Sub find_legal_moves(ByRef colour As Short)
		
		Dim xd As Short
		Dim yd As Short
		Dim i As Short
		Dim flag As Boolean
		Dim j As Short
		Dim m As Short
		Dim n As Short
		Dim numchanges As Short
		Dim move_made As Boolean
		Dim legal(7, 7) As Boolean
		Dim is_move As Boolean
		is_move = False
		Dim points(7, 7) As Short
		Dim xb As Short
		Dim yb As Short
		Dim nomove As Short
		If whoseTurn = white Then
			'UPGRADE_WARNING: Couldn't resolve default property of object find_legal_moves3(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			nomove = find_legal_moves3(black)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object find_legal_moves3(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			nomove = find_legal_moves3(white)
		End If
		Dim best_move As Short
		best_move = -1000
		move_made = False
		
		flag = True
		
		For xb = 0 To 7
			For yb = 0 To 7
				new_board(xb, yb, 0) = chipColour(xb, yb)
				
			Next yb
		Next xb
		
		For xd = 0 To 7
			For yd = 0 To 7
				For xb = 0 To 7
					For yb = 0 To 7
						new_board(xb, yb, 0) = chipColour(xb, yb)
						
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
								If occupied(xd + (n * i), yd + (m * i)) = True Then
									flag = True
									If chipColour(xd + (n * i), yd + (m * i)) = colour Then
										If i >= 2 And occupied(xd, yd) = False Then
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
				
				
				For xb = 0 To 7
					For yb = 0 To 7
						If chipColour(xb, yb) = whoseTurn Then ''''''''''''''''''''''''''
							points(xd, yd) = points(xd, yd) + value_matrix(xb, yb)
						End If
					Next yb
				Next xb
				If whoseTurn = black Then
					'UPGRADE_WARNING: Couldn't resolve default property of object find_legal_moves2(white). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					points(xd, yd) = (points(xd, yd) * value_matrix(xd, yd)) - find_legal_moves2(white)
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object find_legal_moves2(black). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
					points(xd, yd) = (points(xd, yd) * value_matrix(xd, yd)) - find_legal_moves2(black)
				End If
				'If whoseTurn = black Then
				'    points(xd, yd) = (points(xd, yd) - (find_legal_moves2(white) - nomove)) * value_matrix(xd, yd)
				'Else
				'    points(xd, yd) = (points(xd, yd) - (find_legal_moves2(black) - nomove)) * value_matrix(xd, yd)
				'End If
				
				If legal(xd, yd) = True Then
					'UPGRADE_ISSUE: PictureBox property Disp.FillStyle was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
					Main_Renamed.DefInstance.Disp.FillStyle = 0
					If points(xd, yd) > 0 Then
						'Main.Disp.FillColor = RGB(256 - points(xd, yd) * 10, 256 - points(xd, yd) * 10, 256 - points(xd, yd) * 10)
						'UPGRADE_ISSUE: PictureBox property Disp.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
						Main_Renamed.DefInstance.Disp.FillColor = RGB(10, 100, 10)
					Else
						'Main.Disp.FillColor = RGB(points(xd, yd) * 10, points(xd, yd) * 10, points(xd, yd) * 10)
						'UPGRADE_ISSUE: PictureBox property Disp.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
						Main_Renamed.DefInstance.Disp.FillColor = RGB(100, 10, 10)
					End If
					'UPGRADE_ISSUE: PictureBox method Disp.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
					Main_Renamed.DefInstance.Disp.Circle (xd + 0.5, yd + 0.5), 0.04 * System.Math.Abs(points(xd, yd)) + 0.05
					If points(xd, yd) > best_move Then
						best_move = points(xd, yd)
						Main_Renamed.DefInstance.HappyPic.Left = VB6.TwipsToPixelsX(xd + 0.2)
						Main_Renamed.DefInstance.HappyPic.Top = VB6.TwipsToPixelsY(yd + 0.8)
					Else
						If points(xd, yd) = best_move And Rnd() > 0.5 Then
							best_move = points(xd, yd)
						End If
					End If
					is_move = True
				End If
				
			Next yd
		Next xd
		
		If is_move = False Then
			Beep()
			If whoseTurn = black Then
				whoseTurn = white
				'UPGRADE_ISSUE: PictureBox property Disp.MouseIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
				Main_Renamed.DefInstance.Disp.MouseIcon = Main_Renamed.DefInstance.WhitePic.Image
			Else
				If whoseTurn = white Then
					whoseTurn = black
					'UPGRADE_ISSUE: PictureBox property Disp.MouseIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
					Main_Renamed.DefInstance.Disp.MouseIcon = Main_Renamed.DefInstance.BlackPic.Image
				End If
			End If
		End If
		
	End Sub
	
	Function find_legal_moves2(ByRef colour As Short) As Object
		
		Dim xd As Short
		Dim yd As Short
		Dim i As Short
		Dim flag As Boolean
		Dim j As Short
		Dim m As Short
		Dim n As Short
		Dim numchanges As Short
		Dim move_made As Boolean
		Dim legal(7, 7) As Boolean
		Dim is_move As Boolean
		is_move = False
		Dim points(7, 7) As Short
		Dim best_move As Short
		Dim xb As Short
		Dim yb As Short
		
		move_made = False
		Dim nomove As Short
		If colour = white Then
			'UPGRADE_WARNING: Couldn't resolve default property of object find_legal_moves3(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			nomove = find_legal_moves3(black)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object find_legal_moves3(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
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
				
				For xb = 0 To 7
					For yb = 0 To 7
						If new_board(xb, yb, 0) = colour Then
							points(xd, yd) = points(xd, yd) + value_matrix(xb, yb)
						End If
					Next yb
				Next xb
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
							If points(xd, yd) = best_move And Rnd() > 0.5 Then
								best_move = points(xd, yd)
							End If
						End If
						is_move = True
					End If
				End If
			Next yd
		Next xd
		
		'UPGRADE_WARNING: Couldn't resolve default property of object find_legal_moves2. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		find_legal_moves2 = best_move
	End Function
	Function find_legal_moves3(ByRef colour As Short) As Object
		
		Dim xd As Short
		Dim yd As Short
		Dim i As Short
		Dim flag As Boolean
		Dim j As Short
		Dim m As Short
		Dim n As Short
		Dim numchanges As Short
		Dim move_made As Boolean
		Dim legal(7, 7) As Boolean
		Dim is_move As Boolean
		is_move = False
		Dim points(7, 7) As Short
		Dim best_move As Short
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
								If chipColour(xd + (n * i), yd + (m * i)) <> 0 Then
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
				points(xd, yd) = points(xd, yd) * value_matrix(xd, yd)
				If legal(xd, yd) = True Then
					If points(xd, yd) > best_move Then
						best_move = points(xd, yd)
					Else
						If points(xd, yd) = best_move And Rnd() > 0.5 Then
							best_move = points(xd, yd)
						End If
					End If
					is_move = True
				End If
				
			Next yd
		Next xd
		
		'UPGRADE_WARNING: Couldn't resolve default property of object find_legal_moves3. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		find_legal_moves3 = best_move
	End Function
	Function find_score(ByRef colour As Short) As Object
		Dim xd As Short
		Dim yd As Short
		Dim temp_Score As Short
		
		For xd = 0 To 7
			For yd = 0 To 7
				If occupied(xd, yd) = True And chipColour(xd, yd) = colour Then
					temp_Score = temp_Score + 1
				End If
				
			Next yd
		Next xd
		
		'UPGRADE_WARNING: Couldn't resolve default property of object find_score. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		find_score = temp_Score
		
	End Function
	
	Private Sub HappyPic_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles HappyPic.Click
		
		Dim xd As Short
		Dim yd As Short
		Dim you As Short
		
		xd = Int(VB6.PixelsToTwipsX(Main_Renamed.DefInstance.HappyPic.Left))
		yd = Int(VB6.PixelsToTwipsY(Main_Renamed.DefInstance.HappyPic.Top))
		
		'UPGRADE_WARNING: Couldn't resolve default property of object clicked_board(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		you = clicked_board(xd, yd)
	End Sub
	
	Sub load_new_matrix()
		Dim xd As Short
		Dim yd As Short
		
		FileOpen(5, "C:\My Documents\Benjamin\Othello game\matrix b.txt", OpenMode.Input)
		
		For xd = 0 To 7
			For yd = 0 To 7
				Input(5, value_matrix(xd, yd))
				'UPGRADE_ISSUE: PictureBox method Disp.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
				Main_Renamed.DefInstance.Disp.Circle (xd + 0.5, yd + 0.5), value_matrix(xd, yd) / 10
			Next yd
		Next xd
		
		FileClose(5)
		
	End Sub
End Class