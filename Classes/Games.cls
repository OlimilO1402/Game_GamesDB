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
'Klasse speichert in einer Collection alle Spiele die stattgefunden haben
Private m_List As Collection 'Of Game
Private m_NextID As Long

Private Sub Class_Initialize()
    Set m_List = New Collection
    m_NextID = 1
End Sub

Public Function Add(aObj As Game) As Game
    Set Add = aObj
    m_List.Add Add, CStr(Add.ID)
    m_NextID = MaxL(Add.ID, m_NextID)
    m_NextID = m_NextID + 1
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

Public Property Get Item(ByVal ID As Long) As Game
    Set Item = m_List.Item(CStr(ID))
End Property

Public Property Get List() As Collection
    Set List = m_List
End Property

Public Property Get NextID() As Long
    'NextID ist super und muss unbedingt in die List Klasse!!!
    NextID = m_NextID
End Property

'Alle Spiele in einer Listbox anzeigen:
Public Sub ToListBox(aLB As VB.ListBox)
    aLB.Clear
    Dim Game As Game
    For Each Game In m_List
        aLB.AddItem Game.ToStr
    Next
End Sub

'Public Function GetGameByIDNr(IDNr As Long) As Game
'    For Each GetGameByIDNr In mCol
'        If GetGameByIDNr.Nr = IDNr Then Exit Function
'    Next
'End Function
'#############################'  Datei einlesen und abspeichern  '#############################'
''Alle Spiele aus Datei lesen
'Public Sub ReadFromFile(PFN As String)
'    Dim PiG As PlayerInGame
'Try: On Error GoTo Finally
'    Dim FNr As Integer: FNr = FreeFile
'    Open PFN For Input As FNr
'    Call Clear
'    Do While Not EOF(FNr)
'        Dim sLine As String: Line Input #FNr, sLine
'        Dim sArr() As String: sArr() = Split(sLine, ",")
'        Dim Gm As Game: Set Gm = New Game
'        Set Gm.AllPlayers = AllPlayers
'        Gm.Nr = CLng(sArr(0))
'        mLastGamIDNr = Max(mLastGamIDNr, Gm.Nr)
'        Gm.Datum = CDate(sArr(1))
'        Gm.Name = sArr(2)
'        Dim i As Long
'        For i = 3 To UBound(sArr) - 1
'            If IsOdd(i) Then
'                Set PiG = New PlayerInGame
'                PiG.IDNr = CLng(sArr(i))
'            Else
'                PiG.Score = CLng(sArr(i))
'                Call Gm.AddPlayerInGame(PiG)
'            End If
'        Next
'        mCol.Add Gm
'    Loop
'Finally:
'    Close #FNr
'End Sub
'
'Private Function IsOdd(ByVal i As Long) As Boolean 'ist Zahl ungerade
'    If i Mod 2 Then IsOdd = True
'End Function
'
''Alle Spiele in Datei speichern
'Public Sub SaveToFile(PFN As String)
'Try: On Error GoTo Finally
''Alle Spiele und zugehörige IDNr der Spieler mit jeweiligem Punktestand abspeichern.
'    If m_Col.Count = 0 Then If MsgBox("Die Liste AllGames ist leer, soll die Datei überschrieben werden?", vbOKCancel) = vbCancel Then Exit Sub
'    Dim FNr As Integer: FNr = FreeFile
'    Open PFN For Output As #FNr
'    Dim Gm As Game
'    For Each Gm In mCol
'        Print FNr, Gm.ToFileStr
'    Next
'Finally:
'    Close #FNr
'End Sub
'
