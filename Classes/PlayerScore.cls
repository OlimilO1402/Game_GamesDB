VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlayerScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Tja, soll die Zuordnung über IDs erfolgen oder über eine Objekt-Referenz?
'die ID des Spielers: nope!!!
'Public IdNr As Long
Private m_Player As Player
'Private m_Game   As Game
'AArrgh Nein, nicht über die ID, ID ist relationale-DB wir sind aber Objektorientiert
'd.h. die Klasse PlayerInGame mmuss eine Zuordnung zwischen einem Spieler und einem Spiel sein.
'wir brauchen neben Score auch die Möglichkeit einen Spielzustand einzufrieren
'z.B. für Monopoly, welcher Spieler hat wieviel Geld, welche Straßen und wieviele Häuser und Hotels darauf.
'wie könnte man das allgemein lösen?
'da Spiele sehr komplex sein können, eigentlich nur über eine String-Beschreibung
'bei der sich die Spieler selber überlegen wie sie den SPielstand einfrieren.
'ganz schnell geht es z.B. über Fotos
'die Punkte die der Spieler während eines Spiels gemacht hat:
Public Score As Double

Friend Sub New_(aPlayer As Player)
    Set m_Player = aPlayer
End Sub

'ähm wozu ist die Klasse eigentlich nochmal genau da?
'möglicherweise ist die Klasse eigentlich obsolet
Public Property Get Player() As Player
    Set Player = m_Player
End Property

Public Property Get ID() As Long
    ID = m_Player.ID
End Property

'Public Property Get Game() As Game
'    Set Game = m_Game
'End Property

'und dann muss es einen Abbruchspielstand geben,
'wir brauchen eine Klasse, mit variablen Properties, je Spiel
'die Klasse muss sich reproduzieren können, für jeden Spieler das gleiche Set an Properties
'zum Speichern eines Spielstandes
'z.B. Monopoly:
'* Spielerstandpunkt auf dem Spielfeld (Gefängnis, Los, Bahnhof, Straße,
'* Einbehaltene Ereigniskarten: Name
'* Einbehaltene Gemeinschaftskarten: Name
'* Geld
'  - Zwanziger: Anzahl
'  - Tausender: Anzahl
'  - Zehntsder: Anzahl
'
'* Straßen
'  - Straßengruppe1
'    + Straße: Name
'      ? Häuser: Anzahl
'      ? Hotels: Anzahl
'    + Straße: Name
'
Public Sub AddScore(ByVal Value As Double)
    Score = Score + Value
End Sub

Public Function ToStr(Optional bExclNam As Boolean = False) As String
    'die ID immer zuerst!!!
    'OK wir müssen hier mit der Kompatilbiltät brechen!!!
    ToStr = CStr(Me.Score) & ", " & CStr(Me.ID) & IIf(bExclNam, "", ", " & m_Player.Name)
End Function
