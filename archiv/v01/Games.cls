VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Games"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Klasse speichert in einer Collection alle Spiele
'die stattgefunden haben
Private mCol As Collection
Private mLastGamIDNr As Long
Public AllPlayers As Players

Private Sub Class_Initialize()
  Call Clear
End Sub

Public Sub AddGame(G As Game)
  mLastGamIDNr = mLastGamIDNr + 1
  G.GameNr = mLastGamIDNr
  Set G.AllPlayers = AllPlayers
  mCol.Add G
End Sub

Public Sub Clear()
  Set mCol = New Collection
End Sub
'Alle Spiele in einer Listbox anzeigen:
Public Sub ToListBox(LB As VB.ListBox)
Dim Gm As Game
  LB.Clear
  For Each Gm In mCol
    LB.AddItem Gm.ToString
  Next
End Sub

Public Function GetGameByIDNr(IdNr As Long) As Game
  For Each GetGameByIDNr In mCol
    If GetGameByIDNr.GameNr = IdNr Then Exit Function
  Next
End Function

'#############################'  Datei einlesen und abspeichern  '#############################'
'Alle Spiele aus Datei lesen
Public Sub ReadFromFile(PFN As String)
Dim FNr As Integer, mLine As String, StrArr() As String
Dim Gm As Game, i As Long, PiG As PlayerInGame
TryE: On Error GoTo CatchE
  FNr = FreeFile
  Open PFN For Input As #FNr
    Call Clear
    Do While Not EOF(FNr)
      Line Input #FNr, mLine
      StrArr() = Split(mLine, ",")
      Set Gm = New Game
      Set Gm.AllPlayers = AllPlayers
      Gm.GameNr = CLng(StrArr(0))
      mLastGamIDNr = Max(mLastGamIDNr, Gm.GameNr)
      Gm.Datum = CDate(StrArr(1))
      Gm.GameName = StrArr(2)
      For i = 3 To UBound(StrArr) - 1
        If IsOdd(i) Then
          Set PiG = New PlayerInGame
          PiG.IdNr = CLng(StrArr(i))
        Else
          PiG.Score = CLng(StrArr(i))
          Call Gm.AddPlayerInGame(PiG)
        End If
      Next
      mCol.Add Gm
    Loop
  Close #FNr
  Exit Sub
CatchE:
  Close #FNr
End_Try:
End Sub
Private Function IsOdd(i As Long) As Boolean 'ist Zahl ungerade
  If i Mod 2 Then IsOdd = True
End Function
Private Function Max(LngVal1 As Long, LngVal2 As Long) As Long
  If LngVal1 > LngVal2 Then Max = LngVal1 Else Max = LngVal2
End Function
'Alle Spiele in Datei speichern
Public Sub SaveToFile(PFN As String)
TryE: On Error GoTo CatchE
'Alle Spiele und zugehörige IDNr der Spieler mit jeweiligem Punktestand abspeichern.
Dim FNr As Integer, Gm As Game, mbr As VbMsgBoxResult
  mbr = vbOK
  If mCol.Count = 0 Then mbr = MsgBox("Die Liste AllGames ist leer, soll die Datei überschrieben werden?", vbOKCancel)
  If mbr = vbOK Then
    FNr = FreeFile
    Open PFN For Output As #FNr
      For Each Gm In mCol
        Print #FNr, Gm.ToFileStr
      Next
    Close #FNr
  End If
  Exit Sub
CatchE:
  Close #FNr
End_Try:
End Sub

