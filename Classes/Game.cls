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
Private m_ID   As Long
Private m_Date As Date
Public Name    As String
Private m_List As Collection 'Of PlayerScore
'die GameNr würde schon reichen aber sie ist nicht sehr aussagekräftig
'deshalb will man natürlich den Zeitpunkt des Spieles wissen.
'Eine Bezeichnung für das Spiel: Kegeln, Backgammon, Skat ...
'einen Zeiger auf die Liste aller Player, um SpielerNamen zur IdNr herauszufinden
'Private MApp.Players As Players
'Die Liste der Spieler die an diesem Spiel teilnehmen:
'warum nicht einfach nur eine Liste Player?
'wozu ist die Klasse PlayerInGame?
'ist dazu um den Punktestand zu speichern
'd.h. die Klasse PlayerInGame müßte eigentlich PlayerScore heißen oder nur Score

Private Sub Class_Initialize()
    Set m_List = New Collection
End Sub
Friend Sub New_(ByVal aID As Long, ByVal aStartDate As Date, aName As String)
    m_ID = aID: m_Date = aStartDate: Name = aName
End Sub

Public Function Add(aObj As PlayerScore) As PlayerScore
    Set Add = aObj
    m_List.Add Add, CStr(Add.ID)
End Function

Public Sub Clear()
    Class_Initialize
End Sub

Public Function Contains(ByVal key As String) As Boolean
    On Error Resume Next
    If IsEmpty(m_List(key)) Then: 'DoNothing
    Contains = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Property Get Count() As Long
    Count = m_List.Count
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Get Item(ByVal key As String) As PlayerScore
    Set Item = m_List.Item(key)
End Property

Public Property Get StartDate() As Date
    StartDate = m_Date
End Property

Public Property Get List() As Collection
    Set List = m_List
End Property

Public Function ToStr() As String
    ToStr = CStr(m_ID) & ", " & CStr(m_Date) & ", " & Me.Name '& IIf(m_List.Count > 0, ", ", "")
End Function

'Game braucht 2 verschiedene Serialisierungen:
' * für die ListBox
' * für die Datei
'in der Datei kommen hinter dem Spielnamen gleich Score und ID für alle Mitspieler,
'wobei die Namen der Mitspieler in der Datei redundant, aber in der GUI erforderlich sind.
Public Sub ToListBox(aLB As VB.ListBox)
    aLB.Clear
    Dim ps As PlayerScore
    For Each ps In m_List
        aLB.AddItem ps.ToStr
        
        'CStr(ps.Score) & ", " & CStr(ps.Player.ID) & ", " & ps.Player.Name
        'Dim Name As String: Name = AllPlayers.GetNameByIDNr(PiG.IDNr)
        'aLB.AddItem CStr(PiG.Score) & ", " & CStr(PiG.IDNr) & ", " & Name
    Next
End Sub

'Public Sub AddPlayerInGame(PiG As PlayerInGame)
'    If IsInGame(PiG.IDNr) Then
'        MsgBox "Der Spieler " & AllPlayers.GetNameByIDNr(PiG.IDNr) & " ist bereits im Spiel."
'    Else
'        m_Col.Add PiG
'    End If
'End Sub
'
'Public Function IsInGame(ByVal IDNr As Long) As Boolean
'    Dim PiG As PlayerInGame
'    For Each PiG In m_Col
'        If PiG.IDNr = IDNr Then
'            IsInGame = True
'            Exit Function
'        End If
'    Next
'End Function
'
'Public Function GetPlayerInGameByIdNr(IDNr As Long) As PlayerInGame
'    For Each GetPlayerInGameByIdNr In m_Col
'        If GetPlayerInGameByIdNr.IDNr = IDNr Then Exit Function
'    Next
'End Function
'
'
'Public Function ToFileStr() As String
'    Dim PiG As PlayerInGame
'    ToFileStr = ToStr & ", "
'    For Each PiG In mCol
'        ToFileStr = ToFileStr & PiG.ToFileStr & ", "
'    Next
'End Function
