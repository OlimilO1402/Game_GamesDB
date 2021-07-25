VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlayerInGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'die ID des Spielers:
Public IdNr As Long

'die Punkte die der Spieler während eines Spiels gemacht hat:
Public Score As Long

Public Sub AddScore(Sc As Long)
  Score = Score + Sc
End Sub

Public Function ToFileStr() As String
  ToFileStr = CStr(IdNr) & ", " & CStr(Score)
End Function
