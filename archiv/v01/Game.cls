VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public GameNr As Long
'die GameNr würde schon reichen aber sie ist nicht sehr aussagekräftig
'deshalb will man natürlich den Zeitpunkt des Spieles wissen.
Public Datum As Date
'Eine Bezeichnung für das Spiel: Kegeln, Backgammon, Skat ...
Public GameName As String
'einen Zeiger auf die Liste aller Player, um SpielerNamen zur IdNr herauszufinden
Public AllPlayers As Players
'Die Liste der Spieler die an diesem Spiel teilnehmen:
Private mCol As New Collection


Public Sub AddPlayerInGame(PiG As PlayerInGame)
  If IsInGame(PiG.IdNr) Then
    MsgBox "Der Spieler " & AllPlayers.GetNameByIDNr(PiG.IdNr) & " ist bereits im Spiel."
  Else
    mCol.Add PiG
  End If
End Sub
Public Function IsInGame(IdNr As Long) As Boolean
Dim PiG As PlayerInGame
  IsInGame = True
  For Each PiG In mCol
    If PiG.IdNr = IdNr Then Exit Function
  Next
  IsInGame = False
End Function
Public Function GetPlayerInGameByIdNr(IdNr As Long) As PlayerInGame
  For Each GetPlayerInGameByIdNr In mCol
    If GetPlayerInGameByIdNr.IdNr = IdNr Then
      Exit Function
    End If
  Next
End Function
Public Sub ToListBox(LB As VB.ListBox)
Dim PiG As PlayerInGame, Name As String
  LB.Clear
  For Each PiG In mCol
    Name = AllPlayers.GetNameByIDNr(PiG.IdNr)
    LB.AddItem CStr(PiG.Score) & ", " & CStr(PiG.IdNr) & ", " & Name
  Next
End Sub

Public Function ToString() As String
  ToString = CStr(GameNr) & ", " & CStr(Datum) & ", " & GameName
End Function

Public Function ToFileStr() As String
Dim PiG As PlayerInGame
  ToFileStr = ToString & ", "
  For Each PiG In mCol
    ToFileStr = ToFileStr & PiG.ToFileStr & ", "
  Next
End Function
