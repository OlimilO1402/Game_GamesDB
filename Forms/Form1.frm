VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9615
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton BtnOpen 
      Caption         =   "&Open"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "All Games"
      Height          =   4695
      Left            =   3240
      TabIndex        =   9
      Top             =   480
      Width           =   6375
      Begin VB.CommandButton BtnAddScore 
         Caption         =   "&Add Score"
         Height          =   375
         Left            =   3240
         TabIndex        =   6
         Top             =   240
         Width           =   3015
      End
      Begin VB.ListBox LBPlayerScore 
         Height          =   3960
         ItemData        =   "Form1.frx":1782
         Left            =   3240
         List            =   "Form1.frx":1784
         TabIndex        =   7
         Top             =   600
         Width           =   3015
      End
      Begin VB.ListBox LBGames 
         Height          =   3960
         ItemData        =   "Form1.frx":1786
         Left            =   120
         List            =   "Form1.frx":1788
         TabIndex        =   5
         Top             =   600
         Width           =   3015
      End
      Begin VB.CommandButton BtnNewGame 
         Caption         =   "New &Game"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3015
      End
      Begin VB.ListBox LBPlayerScoreSorted 
         Height          =   1815
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "All Players"
      Height          =   4695
      Left            =   0
      TabIndex        =   8
      Top             =   480
      Width           =   3255
      Begin VB.CommandButton BtnNewPlayer 
         Caption         =   "New &Player"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3015
      End
      Begin VB.ListBox LBPlayers 
         Height          =   3960
         ItemData        =   "Form1.frx":178A
         Left            =   120
         List            =   "Form1.frx":178C
         MultiSelect     =   2  'Erweitert
         TabIndex        =   3
         Top             =   600
         Width           =   3015
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   " ? "
      Begin VB.Menu mnuHelpInfo 
         Caption         =   "Info"
      End
   End
   Begin VB.Menu mnuLBAPl 
      Caption         =   "mnuLBAPl"
      Visible         =   0   'False
      Begin VB.Menu mnuLBAPlAddToCurGame 
         Caption         =   "Add to current game"
      End
      Begin VB.Menu mnuLBAPlEdit 
         Caption         =   "Edit"
      End
   End
   Begin VB.Menu mnuLBAPiG 
      Caption         =   "mnuLBAPiG"
      Visible         =   0   'False
      Begin VB.Menu mnuLBAPiGSorted 
         Caption         =   "Sorted"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_CurGame        As Game
Private m_CurPlayer      As Player
Private m_CurPlayerScore As PlayerScore

'Private Sub Command1_Click()
'    MsgBox LBPlayers.Width & vbCrLf & LBGames.Width & vbCrLf & LBPlayerScore.Width
'End Sub

Private Sub Form_Load()
    
    Me.Caption = App.ProductName
    'AllPlayers mit AllGames verknüpfen
    'Set m_AllGames.AllPlayers = MApp.Players
End Sub

Private Function Integer_TryParse(ByVal s As String, int_out As Long) As Boolean
    'parst den ersten Integer aus dem String
Try: On Error GoTo Catch
    s = Left$(s, InStr(1, s, ","))
    If Len(s) = 0 Then Exit Function
    int_out = CLng(s)
    Integer_TryParse = True
Catch:
End Function

Private Sub UpdateAllLB()
    Dim li As Long: li = LBPlayers.ListIndex
    MApp.Players.ToListBox LBPlayers
    LBPlayers.ListIndex = li
    li = LBGames.ListIndex
    MApp.Games.ToListBox LBGames
    If LBGames.ListCount > 0 Then LBGames.ListIndex = li
    If Not m_CurGame Is Nothing Then
        li = LBPlayerScore.ListIndex
        m_CurGame.ToListBox LBPlayerScore: LBPlayerScore.ListIndex = li
        li = LBPlayerScoreSorted.ListIndex
        m_CurGame.ToListBox LBPlayerScoreSorted: LBPlayerScoreSorted.ListIndex = li
    End If
End Sub
Private Sub BtnOpen_Click()
    
    MApp.FileOpen
    
    'MApp.Players.ToListBox LBPlayers
    'MApp.Games.ToListBox LBGames
    UpdateAllLB
End Sub
Private Sub BtnSave_Click()
    MApp.FileSave
End Sub

Private Sub BtnNewPlayer_Click()
    
    Dim s As String: s = InputBox("Geben Sie den Namen des Neuen Mitspielers an:", "Spielername:")
    
    If Len(s) = 0 Then Exit Sub
    
    Set m_CurPlayer = MApp.Players.Add(MNew.Player(MApp.Players.NextID, s))
    Dim i As Long: i = LBPlayers.ListIndex
    MApp.Players.ToListBox LBPlayers
    If i >= 0 Then LBPlayers.Selected(i) = True
    
    'MApp.Players.AddPlayerByName InputBox("Geben Sie den Namen des Neuen Mitspielers an:", "Spielername:")
    'MApp.Players.ToListBox LBPlayers
    
    'UpdateAllLB
End Sub
Private Sub BtnNewGame_Click()
    Dim s As String: s = InputBox("Geben Sie eine Bezeichnung für das Spiel an: (z.B.: Skat, Backgammon, Schach...)", "Spielbezeichnung:")
    If Len(s) = 0 Then Exit Sub
    Dim d As Date: d = Now
    
    Set m_CurGame = MApp.Games.Add(MNew.Game(MApp.Games.NextID, d, s))
    
    MApp.Games.ToListBox LBGames
    
    s = "Sie haben ein neues " & s & "-Spiel angelegt am: " & CStr(d) & vbNewLine
    s = s & "mit Doppelklick bzw. rechte Maustaste in die Liste 'All Players' fügen Sie Spieler hinzu."
    MsgBox s
        'Set m_CurGame = MNew.Game(0, Now, s)
        'm_CurGame.Datum = Now
        'm_CurGame.Name = s
        'm_AllGames.AddGame mCurGame
        'm_AllGames.ToListBox LBGames
        'UpdateAllLB
        'LBGames.ListIndex = LBGames.ListCount - 1
    'End If
End Sub

Private Sub BtnAddScore_Click()
    Call LBPlayerScore_Click
    If m_CurPlayerScore Is Nothing Then
        If LBPlayerScore.ListCount = 0 Then
        Else
            MsgBox "Select a player first."
            Exit Sub
        End If
    End If
    Dim NewListIndex As Long: NewListIndex = LBPlayerScore.ListIndex
    Dim Name As String: Name = m_CurPlayerScore.Player.Name ' MApp.Players.Item(m_CurPlayerScore.ID)
    Dim Pts As String: Pts = InputBox("Diese Punkte werden zum Punktestand (" & CStr(m_CurPlayerScore.Score) & ") des Spielers " & Name & " hinzuaddiert:", "Neue Punkte des Spielers: " & Name)
    If Len(Pts) > 0 Then
        If IsNumeric(Pts) Then
            m_CurPlayerScore.AddScore CDbl(Pts)
            m_CurGame.ToListBox LBPlayerScore 'InGameList
            m_CurGame.ToListBox LBPlayerScoreSorted
        Else
            MsgBox "Please give a number."
        End If
    End If
    'NewListIndex = NewListIndex + 1 'öhm, why??
    If NewListIndex = LBPlayerScore.ListCount Then NewListIndex = 0
    LBPlayerScore.ListIndex = NewListIndex
    'End If
End Sub

Private Sub Form_Resize()
    Dim brdr As Single: brdr = 8 * Screen.TwipsPerPixelX
    Dim L As Single, T As Single, W As Single, H As Single
    Dim W3 As Single: W3 = Me.ScaleWidth / 3
    Dim FH As Single: FH = Me.ScaleHeight - Frame1.Top
    
    
    L = 0:   T = Frame1.Top
    W = W3:  H = FH
    If W > 0 And H > 0 Then Frame1.Move L, T, W, H
    
    L = W3
    W = 2 * W3
    If W > 0 And H > 0 Then Frame2.Move L, T, W, H
        
    L = brdr:          T = BtnNewPlayer.Top
    W = W3 - L - brdr: H = BtnNewPlayer.Height
    If W > 0 And H > 0 Then BtnNewPlayer.Move L, T, W, H
        
    L = brdr:          T = LBPlayers.Top
    W = W3 - L - brdr: H = FH - T - brdr
    If W > 0 And H > 0 Then LBPlayers.Move L, T, W, H
    
    
    L = brdr:                    T = BtnNewGame.Top
    W = (2 * W3 - 3 * brdr) / 2: H = BtnNewGame.Height
    If W > 0 And H > 0 Then BtnNewGame.Move L, T, W, H
    
    L = brdr:                    T = LBGames.Top
    W = (2 * W3 - 3 * brdr) / 2: H = FH - T - brdr
    If W > 0 And H > 0 Then LBGames.Move L, T, W, H
    
    L = brdr + W + brdr:         T = BtnAddScore.Top
    W = W:                       H = BtnAddScore.Height
    If W > 0 And H > 0 Then BtnAddScore.Move L, T, W, H
    
    T = LBPlayerScore.Top
    H = FH - T - brdr
    If W > 0 And H > 0 Then LBPlayerScore.Move L, T, W, H
    If W > 0 And H > 0 Then LBPlayerScoreSorted.Move L, T, W, H
    
End Sub

Private Sub LBGames_Click()
    'beim Klick in die Listbox das aktuelle Spiel setzen
    Set m_CurPlayerScore = Nothing
    If LBGames.ListCount = 0 Then Exit Sub
    If LBGames.ListIndex < 0 Then Exit Sub
    Dim s As String: s = LBGames.List(LBGames.ListIndex)
    If Len(s) = 0 Then Exit Sub
    Dim ID As Long
    If Not Integer_TryParse(s, ID) Then
        MsgBox "Could not get ID from listentry"
        Exit Sub
    End If
    If Not MApp.Games.Contains(CStr(ID)) Then Exit Sub
    Set m_CurGame = MApp.Games.Item(ID) 'GetGameByIDNr(CLng(Left$(str, InStr(str, ","))))
    m_CurGame.ToListBox LBPlayerScore
    m_CurGame.ToListBox LBPlayerScoreSorted
End Sub

Private Sub mnuFileNew_Click()
    MApp.FileNew
    Set m_CurGame = Nothing
    Set m_CurPlayer = Nothing
    Set m_CurPlayerScore = Nothing
    
    UpdateAllLB
End Sub

Private Sub mnuFileOpen_Click()
    MApp.FileOpen
    UpdateAllLB
End Sub

Private Sub mnuFileSave_Click()
    MApp.FileSave
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuHelpInfo_Click()
    MsgBox MApp.InfoStr
End Sub

Private Sub mnuLBAPiGSorted_Click()
    mnuLBAPiGSorted.Checked = Not mnuLBAPiGSorted.Checked
    If mnuLBAPiGSorted.Checked Then
        LBPlayerScoreSorted.ZOrder 0
    Else
        LBPlayerScore.ZOrder 0
    End If
End Sub

Private Sub mnuLBAPlEdit_Click()
    If LBPlayers.ListIndex >= 0 Then
        Dim isel As Long:   isel = LBPlayers.ListIndex
        Dim aStr As String: aStr = LBPlayers.List(isel)
        Dim ID As Long:       ID = CLng(Left$(aStr, InStr(aStr, ",")))
        Dim pl As Player: Set pl = MApp.Players.Item(CStr(ID))
        aStr = InputBox("Geben sie einen neuen Namen für den Spieler " & pl.Name & " ein:", "Neuer Name:", pl.Name)
        If Len(aStr) > 0 Then
            pl.Name = aStr
            MApp.Players.ToListBox LBPlayers
            LBPlayers.ListIndex = isel
            MApp.Games.ToListBox LBGames
        End If
    End If
End Sub

Private Sub LBPlayerScore_Click()
'beim Klick in die Listbox den aktuellen Spieler im aktuellen Spiel auswählen
    Call SetCurPlayerScore
End Sub
Private Sub SetCurPlayerScore()
    
    If LBPlayerScore.ListCount = 0 Then Exit Sub
    If LBPlayerScore.ListIndex < 0 Then Exit Sub
    Dim s As String: s = LBPlayerScore.List(LBPlayerScore.ListIndex)
    s = Right(s, Len(s) - InStr(1, s, ","))
    Dim ID As Long
    If Not Integer_TryParse(s, ID) Then
        MsgBox "Could not get ID"
        Exit Sub
    End If
    If Not m_CurGame.Contains(CStr(ID)) Then Exit Sub
    Set m_CurPlayerScore = m_CurGame.Item(CStr(ID))
    
'        Dim s As String: s = LBPlayerScore.List(LBPlayerScore.ListIndex)
'        Dim p1 As Long: p1 = InStr(1, s, ",") + 1
'        Dim p2 As Long: p2 = InStr(p1, s, ",")
'        Dim sID As String: sID = Trim$(Mid$(s, p1, p2 - p1))
    'End If
End Sub

'OK das mit der Sorted-ListBox ist ein Schmarrn, bei numerischen Werte geht das nicht richtig
'weil 10 vor 9 einsortiert wird, das ist Mist 10, 7, 8, 9
Private Sub LBPlayerScore_DblClick()
    SetCurPlayerScore
    Dim s As String: s = InputBox("Geben Sie den Gesamtpunktestand des Spielers: " & m_CurPlayerScore.Player.Name & " im aktuellen Spiel an:", "Gesamtspielstand des Spielers:", CStr(m_CurPlayerScore.Score))
    If Len(s) > 0 Then
        If Not IsNumeric(s) Then
            MsgBox "Geben Sie eine ganze Zahl ein."
            Exit Sub
        End If
        m_CurPlayerScore.Score = CDbl(s)
        m_CurGame.ToListBox LBPlayerScore
        m_CurGame.ToListBox LBPlayerScoreSorted
    End If
End Sub

Private Sub mnuLBAPlAddToCurGame_Click()
    If m_CurGame Is Nothing Then Exit Sub
    If LBPlayers.ListCount = 0 Then
        MsgBox "Select player first"
        Exit Sub
    End If
    Dim i As Long: i = LBPlayers.ListIndex
    AddPlToCurGame LBPlayers.List(i)
    
    'm_curgame.Add mnew.PlayerScore(
'
'            For i = 0 To LBPlayers.ListCount - 1
'                If LBPlayers.Selected(i) Then
'                    LBPlayers.Selected(i) = False
'                    StrE = LBPlayers.List(i)
'                    AddPlToCurGame (StrE)
'                End If
'            Next
            m_CurGame.ToListBox LBPlayerScore
            m_CurGame.ToListBox LBPlayerScoreSorted
            LBPlayerScore.ListIndex = LBPlayerScore.ListCount - 1
'        End If
'    End If
End Sub
Private Sub AddPlToCurGame(sEntry As String)
    Dim ID As Long: ID = CLng(Left$(sEntry, InStr(1, sEntry, ",")))
    Dim pl As Player: Set pl = MApp.Players.Item(CStr(ID))
    If pl Is Nothing Then
        MsgBox "Player not found: " & ID
        Exit Sub
    End If
    m_CurGame.Add MNew.PlayerScore(pl)
End Sub

Private Sub LBPlayerScore_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu mnuLBAPiG
End Sub
Private Sub LBPlayerScoreSorted_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu mnuLBAPiG
End Sub

Private Sub LBPlayers_DblClick()
    If LBPlayers.ListCount = 0 Then Exit Sub
    If m_CurGame Is Nothing Then Exit Sub
    Dim s As String: s = LBPlayers.List(LBPlayers.ListIndex)
    Dim ID As String: ID = CLng(Left$(s, InStr(1, s, ",")))
    If Len(s) = 0 Then Exit Sub
    Set m_CurPlayer = MApp.Players.Item(ID)
    If m_CurGame.Contains(ID) Then
        MsgBox "Der Spieler ist bereits ein Mitspieler im diesem Spiel."
        Exit Sub
    End If
    m_CurGame.Add MNew.PlayerScore(m_CurPlayer)
    
    'Dim StrE As String
    '
    'AddPlToCurGame LBPlayers.List(LBPlayers.ListIndex)
    m_CurGame.ToListBox LBPlayerScore
    m_CurGame.ToListBox LBPlayerScoreSorted
    LBPlayerScore.ListIndex = LBPlayerScore.ListCount - 1
    'End If
End Sub

Private Sub LBPlayers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuLBAPl, , , , mnuLBAPlAddToCurGame
    End If
End Sub
