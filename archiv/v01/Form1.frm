VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame2 
      Caption         =   "All Games"
      Height          =   2535
      Left            =   2520
      TabIndex        =   5
      Top             =   600
      Width           =   5415
      Begin VB.CommandButton BtnAddScore 
         Caption         =   "&Add Score"
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
      Begin VB.ListBox PlayersInGameList 
         Height          =   1815
         ItemData        =   "Form1.frx":0000
         Left            =   3240
         List            =   "Form1.frx":0002
         TabIndex        =   8
         Top             =   600
         Width           =   2055
      End
      Begin VB.ListBox GamesList 
         Height          =   1815
         ItemData        =   "Form1.frx":0004
         Left            =   120
         List            =   "Form1.frx":0006
         TabIndex        =   7
         Top             =   600
         Width           =   3015
      End
      Begin VB.CommandButton BtnNewGame 
         Caption         =   "New &Game"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3015
      End
      Begin VB.ListBox PlayersInGameListSorted 
         Height          =   1815
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.CommandButton BtnOpen 
      Caption         =   "&Open"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton BtnSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "All Players"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2295
      Begin VB.CommandButton BtnNewPlayer 
         Caption         =   "New &Player"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.ListBox PlayersList 
         Height          =   1815
         ItemData        =   "Form1.frx":0008
         Left            =   120
         List            =   "Form1.frx":000A
         MultiSelect     =   2  'Erweitert
         TabIndex        =   1
         Top             =   600
         Width           =   2055
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mAllPlayers As New Players
Private mPlayersPFN As String
Private mAllGames As New Games
Private mGamesPFN As String
Private mCurGame As Game
Private mCurPiG As PlayerInGame

Private Sub Form_Load()
  Me.Caption = App.ProductName
  mPlayersPFN = App.Path & "\" & "AllPlayers.pls"
  mGamesPFN = App.Path & "\" & "AllGames.gms"
  'AllPlayers mit AllGames verknüpfen
  Set mAllGames.AllPlayers = mAllPlayers
End Sub
Private Sub UpdateAllLB()
Dim LI As Long
  LI = PlayersList.ListIndex
  Call mAllPlayers.ToListBox(PlayersList): PlayersList.ListIndex = LI
  LI = GamesList.ListIndex
  Call mAllGames.ToListBox(GamesList): GamesList.ListIndex = LI
  If Not mCurGame Is Nothing Then
    LI = PlayersInGameList.ListIndex
    Call mCurGame.ToListBox(PlayersInGameList): PlayersInGameList.ListIndex = LI
    LI = PlayersInGameListSorted.ListIndex
    Call mCurGame.ToListBox(PlayersInGameListSorted): PlayersInGameListSorted.ListIndex = LI
  End If
End Sub
Private Sub BtnOpen_Click()
  mAllPlayers.ReadFromFile (mPlayersPFN)
  Call mAllPlayers.ToListBox(PlayersList)
  mAllGames.ReadFromFile (mGamesPFN)
  Call mAllGames.ToListBox(GamesList)
  'UpdateAllLB
End Sub
Private Sub BtnSave_Click()
  mAllPlayers.SaveToFile (mPlayersPFN)
  mAllGames.SaveToFile (mGamesPFN)
End Sub

Private Sub BtnNewPlayer_Click()
  mAllPlayers.AddPlayerByName (InputBox("Geben Sie den Namen des Neuen Mitspielers an:", "Spielername:"))
  Call mAllPlayers.ToListBox(PlayersList)
  'UpdateAllLB
End Sub
Private Sub BtnNewGame_Click()
Dim aStr As String
  aStr = InputBox("Geben Sie eine Bezeichnung für das Spiel an: (z.B.: Skat, Backgammon, Schach...)", "Spielbezeichnung:")
  If Len(aStr) > 0 Then
    Set mCurGame = New Game
    mCurGame.Datum = Now
    mCurGame.GameName = aStr
    Call mAllGames.AddGame(mCurGame)
    Call mAllGames.ToListBox(GamesList)
    'UpdateAllLB
    GamesList.ListIndex = GamesList.ListCount - 1
    aStr = "Sie haben ein neues " & aStr & "-Spiel angelegt am: " & CStr(mCurGame.Datum) & vbNewLine
    aStr = aStr & "mit Doppelklick bzw. rechte Maustaste in die Liste 'All Players' fügen Sie Spieler hinzu."
    MsgBox aStr
  End If
End Sub
Private Sub BtnAddScore_Click()
Dim Pts As String, Name As String, NewListIndex As Long
  Call PlayersInGameList_Click
  If Not mCurPiG Is Nothing Then
    NewListIndex = PlayersInGameList.ListIndex
    Name = mAllPlayers.GetNameByIDNr(mCurPiG.IdNr)
    Pts = InputBox("Diese Punkte werden zum Punktestand (" & CStr(mCurPiG.Score) & ") des Spielers " & Name & " hinzuaddiert:", "Neue Punkte des Spielers: " & Name)
    If Len(Pts) > 0 Then
      If IsNumeric(Pts) Then
        mCurPiG.AddScore (CLng(Pts))
        Call mCurGame.ToListBox(PlayersInGameList)
        Call mCurGame.ToListBox(PlayersInGameListSorted)
        'UpdateAllLB
      Else
        MsgBox "Geben Sie eine ganze Zahl ein."
      End If
    End If
    NewListIndex = NewListIndex + 1
    If NewListIndex = PlayersInGameList.ListCount Then NewListIndex = 0
    PlayersInGameList.ListIndex = NewListIndex
  End If
End Sub

Private Sub GamesList_Click()
'beim Klick in die Listbox das aktuelle Spiel setzen
Dim str As String
  str = GamesList.List(GamesList.ListIndex)
  Set mCurGame = mAllGames.GetGameByIDNr(CLng(Left$(str, InStr(str, ","))))
  Call mCurGame.ToListBox(PlayersInGameList)
  Call mCurGame.ToListBox(PlayersInGameListSorted)
End Sub

Private Sub mnuLBAPiGSorted_Click()
  mnuLBAPiGSorted.Checked = Not mnuLBAPiGSorted.Checked
  If mnuLBAPiGSorted.Checked Then
    PlayersInGameListSorted.ZOrder 0
  Else
    PlayersInGameList.ZOrder 0
  End If
End Sub

Private Sub mnuLBAPlEdit_Click()
Dim aStr As String, Pl As Player, isel As Long
  If PlayersList.ListIndex >= 0 Then
    isel = PlayersList.ListIndex
    aStr = PlayersList.List(isel)
    Set Pl = mAllPlayers.GetPlayerByIDNr(CLng(Left$(aStr, InStr(aStr, ","))))
    aStr = InputBox("Geben sie einen neuen Namen für den Spieler " & Pl.Name & " ein:", "Neuer Name:", Pl.Name)
    If Len(aStr) > 0 Then
      Pl.Name = aStr
      Call mAllPlayers.ToListBox(PlayersList)
      PlayersList.ListIndex = isel
      Call mAllGames.ToListBox(GamesList)
    End If
  End If
End Sub

Private Sub PlayersInGameList_Click()
'beim Klick in die Listbox den aktuellen Spieler im aktuellen Spiel auswählen
  Call SetCurPiG
End Sub
Private Sub SetCurPiG()
Dim aStr As String, C1Pos As Long, C2Pos As Long, strID As String
  If PlayersInGameList.ListCount > 0 Then
    If PlayersInGameList.ListIndex < 0 Then PlayersInGameList.ListIndex = 0
    aStr = PlayersInGameList.List(PlayersInGameList.ListIndex)
    C1Pos = InStr(1, aStr, ",") + 1
    C2Pos = InStr(C1Pos, aStr, ",")
    strID = Trim$(Mid$(aStr, C1Pos, C2Pos - C1Pos))
    Set mCurPiG = mCurGame.GetPlayerInGameByIdNr(CLng(strID))
  End If
End Sub
Private Sub PlayersInGameList_DblClick()
Dim Pts As String
  Call SetCurPiG
  Pts = InputBox("Geben Sie den Gesamtpunktestand des Spielers: " & mAllPlayers.GetNameByIDNr(mCurPiG.IdNr) & " im aktuellen Spiel an:", "Gesamtspielstand des Spielers:", CStr(mCurPiG.Score))
  If Len(Pts) > 0 Then
    If IsNumeric(Pts) Then
      mCurPiG.Score = CLng(Pts)
      Call mCurGame.ToListBox(PlayersInGameList)
      Call mCurGame.ToListBox(PlayersInGameListSorted)
    Else
      MsgBox "Geben Sie eine ganze Zahl ein."
    End If
  End If
End Sub

Private Sub mnuLBAPlAddToCurGame_Click()
Dim StrE As String, i As Long
  If PlayersList.SelCount > 0 Then
    If Not mCurGame Is Nothing Then
      For i = 0 To PlayersList.ListCount - 1
        If PlayersList.Selected(i) Then
          PlayersList.Selected(i) = False
          StrE = PlayersList.List(i)
          AddPlToCurGame (StrE)
        End If
      Next
      Call mCurGame.ToListBox(PlayersInGameList)
      Call mCurGame.ToListBox(PlayersInGameListSorted)
      PlayersInGameList.ListIndex = PlayersInGameList.ListCount - 1
    End If
  End If
End Sub
Private Sub AddPlToCurGame(StrE As String)
Dim PiG As New PlayerInGame
  PiG.IdNr = CLng(Left$(StrE, InStr(1, StrE, ",")))
  Call mCurGame.AddPlayerInGame(PiG)
End Sub

Private Sub PlayersInGameList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then PopupMenu mnuLBAPiG
End Sub
Private Sub PlayersInGameListSorted_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then PopupMenu mnuLBAPiG
End Sub

Private Sub PlayersList_DblClick()
Dim StrE As String
  If PlayersList.ListCount > 0 Then
    If Not mCurGame Is Nothing Then
      AddPlToCurGame (PlayersList.List(PlayersList.ListIndex))
      Call mCurGame.ToListBox(PlayersInGameList)
      Call mCurGame.ToListBox(PlayersInGameListSorted)
      PlayersInGameList.ListIndex = PlayersInGameList.ListCount - 1
    End If
  End If
End Sub

Private Sub PlayersList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
    PopupMenu mnuLBAPl, , , , mnuLBAPlAddToCurGame
  End If
End Sub
