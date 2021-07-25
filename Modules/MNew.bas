Attribute VB_Name = "MNew"
Option Explicit

Public Function Player(ByVal aID As Long, aName As String) As Player
    Set Player = New Player: Player.New_ aID, aName
End Function

Public Function Game(ByVal aID As Long, aDate As Date, aName As String) As Game
    Set Game = New Game: Game.New_ aID, aDate, aName
End Function

Public Function PlayerScore(aPlayer As Player) As PlayerScore
    Set PlayerScore = New PlayerScore: PlayerScore.New_ aPlayer
End Function
