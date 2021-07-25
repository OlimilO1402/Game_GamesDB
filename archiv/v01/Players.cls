VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Players"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mCol As Collection
Private mLastIDNr As Long


Private Sub Class_Initialize()
  Call Clear
End Sub
Public Sub Clear()
  Set mCol = New Collection
End Sub
'einen neuen Mitspieler erstellen:
Public Sub AddPlayerByName(aName As String)
  If Len(aName) > 0 Then
    mLastIDNr = mLastIDNr + 1 'muß unique sein
    AddPlayer (aName)
  End If
End Sub

'eine Private Sub AddPlayer, deshalb weil die IDNr intern entweder
'von AddPlayerByName oder
'von ReadFromfile festgelegt wird.
Private Sub AddPlayer(aName As String)
Dim NewPlayer As New Player
  NewPlayer.Name = aName
  NewPlayer.IdNr = mLastIDNr
  mCol.Add NewPlayer
End Sub

Public Function GetNameByIDNr(IdNr As Long) As String
  GetNameByIDNr = GetPlayerByIDNr(IdNr).Name
End Function
Public Function GetPlayerByIDNr(IdNr As Long) As Player
'es gäbe unter VB6 natürlich noch eine viel elegantere Möglichkeit
'nämlich die IDNr als Key mit in der Collection abzuspeichern
'genau für den Fall wäre der Key wahrscheinlich gedacht
  For Each GetPlayerByIDNr In mCol
    If GetPlayerByIDNr.IdNr = IdNr Then
      Exit Function
    End If
  Next
End Function

'Alle Spieler in einer Listbox anzeigen:
Public Sub ToListBox(LB As VB.ListBox)
Dim Pl As Player
  LB.Clear
  For Each Pl In mCol
    LB.AddItem Pl.ToString
  Next
End Sub

'#############################'  Datei einlesen und abspeichern  '#############################'
'Alle Spieler aus Datei lesen
Public Sub ReadFromFile(PFN As String)
Dim FNr As Integer, mLine As String, CPos As Long
TryE: On Error GoTo CatchE
  FNr = FreeFile
  Open PFN For Input As #FNr
    Call Clear
    Do While Not EOF(FNr)
      Line Input #FNr, mLine
      CPos = InStr(1, mLine, ",")
      mLastIDNr = CLng(Trim(Left$(mLine, CPos)))
      AddPlayer (Trim(Right$(mLine, Len(mLine) - CPos)))
    Loop
  Close #FNr
  Exit Sub
CatchE:
  Close #FNr
End_Try:
End Sub

'Alle Spieler in Datei speichern
Public Sub SaveToFile(PFN As String)
Dim FNr As Integer, Pl As Player, mbr As VbMsgBoxResult
TryE: On Error GoTo CatchE
  mbr = vbOK
  If mCol.Count = 0 Then mbr = MsgBox("Die Liste AllPlayers ist leer, soll die Datei überschrieben werden?", vbOKCancel)
  If mbr = vbOK Then
    FNr = FreeFile
    Open PFN For Output As #FNr
      For Each Pl In mCol
        Print #FNr, Pl.ToString
      Next
    Close #FNr
  End If
  Exit Sub
CatchE:
  Close #FNr
End_Try:
End Sub

