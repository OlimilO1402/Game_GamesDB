Attribute VB_Name = "MApp"
Option Explicit
'Spieleabendverwaltung
'Sie kennen sicher die Situation, eine Gruppe von bspw 7 Freunden verabredet sich zu einem geselligen Spieleabend.
'Es gibt Spiele bei denen können alle Freunde mitspielen, z.B. Mikado, hier ist die Anzahl der Spieler beliebig.
'bei manchen Spielen können aber nur eine maximale Anzahl Personen mitspielen. z.b. bei Schach oder bei Mensch-
'Ärgere-Dich-Nicht, hier können maximal 2 bzw 4 Personen mitspielen.
'
'Es gibt eine Liste mit allen Spielern MApp.Players
'und eine Liste mit allen Spielen, nicht jeder Spieler macht bei jedem Spiel mit
'
Public Players As Players
Public Games   As Games
'Public CurGame As Game
'Public CurPiG  As PlayerInGame

Public PFNPlayers As String
Public PFNGames   As String

Sub Main()
    PFNPlayers = App.Path & "\" & "AllPlayers_2021.pls"
    PFNGames = App.Path & "\" & "AllGames_2021.gms"

    FileNew
    
    FrmMain.Show
End Sub

Public Property Get InfoStr() As String
    With App
        InfoStr = .CompanyName & _
            " " & .EXEName & _
            " " & .Major & _
            "." & .Minor & _
            "." & .Revision & vbCrLf & _
                  .FileDescription
    End With
End Property
Public Function MaxL(ByVal V1 As Long, ByVal V2 As Long) As Long
    If V1 > V2 Then MaxL = V1 Else MaxL = V2
End Function

Public Sub FileNew()
    Set Players = New Players
    Set Games = New Games
End Sub

Public Sub FileSave()
    SavePlayers
    SaveGames
End Sub

Private Sub SavePlayers()
Try: On Error GoTo Finally
    'Kill PFNPlayers
    Dim FNr As Integer: FNr = FreeFile
    Open PFNPlayers For Binary Access Write As FNr
    MSerial.Players Players
    Put FNr, , MSerial.ToStr
    MSerial.Clear
Finally:
    Close FNr
End Sub

Private Sub SaveGames()
Try: On Error GoTo Finally
    'Kill GamesPFN
    Dim FNr As Integer: FNr = FreeFile
    Open PFNGames For Binary Access Write As FNr
    MSerial.Games Games
    Put FNr, , MSerial.ToStr
    MSerial.Clear
Finally:
    Close FNr
End Sub

Public Sub FileOpen()
    MDeseri.TryRead_Players PFNPlayers, Players
    MDeseri.TryRead_Games PFNGames, Games
End Sub

