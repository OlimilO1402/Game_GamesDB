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
Private m_List   As Collection 'Of Player
'Private m_LastIDNr As Long _'Blödsinn, die LastID ist einfach die Anzahl der Elemente
Private m_NextID As Long

'Halt Stop, die ID darf nicht einfach nur Count sein,
'wenn die ID immer für einen neuen Player gelten soll
'so muss die ID globally sein.
'weil wenn ein Spieler gelöscht wird, dann kann durch abspeichern der Datei und wieder einlesen der Datei
'd.h. wir brauchen MaxVersion

Private Sub Class_Initialize()
    Set m_List = New Collection
    m_NextID = 1 'Die ID ist 1-basiert
End Sub

'einen neuen Mitspieler hinzufügen
'die Nächste ID hochzählen
'Ein Spieler kann auf 2 Arten hinzugefügt werden:
' * es ist noch kein Spieler vorhanden, der User wählt "Neuen Spieler" hinzufügen
'   die nächste ID muss immer um eins Höher sein als die ID des vorherigen Spielers
' * der User liest die Datenbankdatei ein, jeder Spieler muss mit der eingelesenen
'   ID hinzugefügt werden
'   wird danach ein Neuer Spieler hinzugefügt, dann muss er wieder die maximale ID
'   erhalten.
'   was ist wenn der User einen Spieler per Dialog hinzufügt, und erst dann die Datei einliest?
'   das darf nicht passieren, die Datenbank muss durch das Programm automatisch bei Programmstart
'   geladen werden.
'
Public Function Add(aObj As Player) As Player
    'm_NextID = MaxL(m_Col.Count, m_NextID) + 1
    Set Add = aObj
    m_List.Add Add, CStr(Add.ID)
    m_NextID = MaxL(Add.ID, m_NextID)
    m_NextID = m_NextID + 1
    'If Len(aName) > 0 Then
    '    mLastIDNr = mLastIDNr + 1 'muß unique sein
    '    AddPlayer (aName)
    'End If
End Function

Public Sub Clear()
    Class_Initialize
End Sub

Public Property Get Count() As Long
    Count = m_List.Count
End Property

Public Property Get Item(ByVal key As String) As Player
    Set Item = m_List.Item(key)
End Property

Public Property Get List() As Collection 'Of Player
    Set List = m_List
End Property

Public Property Get NextID() As Long
    NextID = m_NextID
End Property

'eine Private Sub AddPlayer, deshalb weil die IDNr intern entweder
'von AddPlayerByName oder
'von ReadFromfile festgelegt wird.
'Private Sub AddPlayer(aName As String)
'    Dim NewPlayer As New Player
'    NewPlayer.Name = aName
'    NewPlayer.IDNr = mLastIDNr
'    mCol.Add NewPlayer
'End Sub

'Public Function GetNameByIDNr(IDNr As Long) As String
'    GetNameByIDNr = GetPlayerByIDNr(IDNr).Name
'End Function
'Public Function GetPlayerByIDNr(IDNr As Long) As Player
'    'es gäbe unter VB6 natürlich noch eine viel elegantere Möglichkeit
'    'nämlich die IDNr als Key mit in der Collection abzuspeichern
'    'genau für den Fall wäre der Key wahrscheinlich gedacht
'    For Each GetPlayerByIDNr In mCol
'        If GetPlayerByIDNr.IDNr = IDNr Then
'            Exit Function
'        End If
'    Next
'End Function

'Alle Spieler in einer Listbox anzeigen:
Public Sub ToListBox(aLB As VB.ListBox)
    aLB.Clear
    Dim pl As Player
    For Each pl In m_List
        aLB.AddItem pl.ToStr 'MSerial.Player(pl)
    Next
End Sub
