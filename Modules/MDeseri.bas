Attribute VB_Name = "MDeseri"
Option Explicit

'#############################'   Spieler-Datei einlesen   '#############################'
'Alle Spieler aus Datei lesen
'Format z.B.:
'1, Oliver
'2, Peter
'3, Paul
'4, Mary
'5, Simon
Public Function TryRead_Players(aPFN As String, pls_out As Players) As Boolean
Try: On Error GoTo Finally
    Dim FNr As Integer: FNr = FreeFile
    Open aPFN For Input As #FNr
    'Call Clear
    If pls_out.Count Then pls_out.Clear
    Dim sLine As String, CPos As Long
    Do While Not EOF(FNr)
        Line Input #FNr, sLine
        CPos = InStr(1, sLine, ",")
        Dim ID As Long:   ID = CLng(Trim(Left$(sLine, CPos)))
        Dim Nm As String: Nm = Trim$(Mid(sLine, CPos + 1))
        pls_out.Add MNew.Player(ID, Nm)
        'AddPlayer (Trim(Right$(sLine, Len(sLine) - CPos)))
    Loop
    TryRead_Players = True
Finally:
    Close #FNr
End Function

'#############################'   Spiele-Datei einlesen   '#############################'
'Alle Spiele aus Datei lesen
'Format z.B.:
'ID, DateTime          , Name,     Sco, ID, Sco, ID, Sco, ID . . .
'1, 25.07.2021 11:48:04, Chess,    150, 1, 300, 5
'2, 25.07.2021 11:48:10, Monopoly,  10, 2,  85, 3,  65, 4
'3, 25.07.2021 11:48:38, Monopoly,   5, 1,   9, 5,   8, 3
'4, 25.07.2021 11:49:06, Chess,     15, 2,  25, 4

'Ähm Nö,
'* Bei Serialisierung:
'   es muss nur gelesen werden -> Serialisierung von außen,
'* Bei Deserialisierung:
'   es muss geschrieben werden -> Deserialisierung von innen, mit einer public Parse-Methode
'hier: egal weil nur sehr wenig Daten
Public Function TryRead_Games(aPFN As String, gms_out As Games) As Boolean
Try: On Error GoTo Finally
    Dim FNr As Integer: FNr = FreeFile
    Open aPFN For Input As #FNr
    'Call Clear
    If gms_out.Count Then gms_out.Clear
    Dim sLine As String, CPos As Long
    Do While Not EOF(FNr)
        
        Line Input #FNr, sLine
        sLine = Trim$(sLine)
        
        Dim gm As Game
        If TryRead_Game(sLine, gm) Then
            gms_out.Add gm
        End If
        
    Loop
    TryRead_Games = True
Finally:
    Close #FNr
End Function

Private Function TryRead_Game(ByVal sLine As String, gm_out As Game) As Boolean
Try: On Error GoTo Catch
    Dim CPos As Long
    
    If Len(sLine) = 0 Then Exit Function
    CPos = InStr(1, sLine, ",")
    Dim ID As Long:   ID = CLng(Left$(sLine, CPos - 1)):  sLine = Trim$(Mid$(sLine, CPos + 1))
    
    If Len(sLine) = 0 Then Exit Function
    CPos = InStr(1, sLine, ",")
    Dim Dt As Date:   Dt = CDate(Left$(sLine, CPos - 1)): sLine = Trim$(Mid$(sLine, CPos + 1))
    
    If Len(sLine) = 0 Then Exit Function
    CPos = InStr(1, sLine, ",")
    Dim Nm As String: Nm = Left$(sLine, CPos - 1):        sLine = Trim$(Mid$(sLine, CPos + 1))
    
    Set gm_out = MNew.Game(ID, Dt, Nm)
    Read_PlayerScores sLine, gm_out
    
    TryRead_Game = True
    Exit Function
Catch:
    MsgBox "Error during reading game object: " & ID & ", " & Dt & ", " & Nm
End Function

Private Sub Read_PlayerScores(ByVal sLine As String, gm_out As Game)
Try: On Error GoTo Catch
    Dim sLineL As String
    Do Until Len(sLine) = 0
        Dim CPos As Long
        CPos = InStr(1, sLine, ",")
        CPos = InStr(CPos + 1, sLine, ",")
        'sLineL = IIf(CPos, Left$(sLine, CPos - 1), sLine) 'Nooo!
        'sLine = IIf(CPos, Trim(Mid(sLine, CPos + 1)), "") 'Nooo!
        If CPos Then
            sLineL = Left$(sLine, CPos - 1)
            sLine = Trim(Mid(sLine, CPos + 1))
        Else
            sLineL = sLine
            sLine = ""
        End If
        If Len(sLineL) Then
            Dim ps As PlayerScore
            If TryRead_PlayerScore(sLineL, ps) Then
                gm_out.Add ps
            End If
        End If
    Loop
    Exit Sub
Catch:
    MsgBox "Error during reading PlayerScores: " & sLine
End Sub

Private Function TryRead_PlayerScore(ByVal sLine As String, ps_out As PlayerScore) As Boolean
Try: On Error GoTo Catch
    If Len(sLine) = 0 Then Exit Function
        
    Dim CPos As Long: CPos = InStr(1, sLine, ",")
    
    Dim Sc As Double: Sc = CDbl(Left$(sLine, CPos - 1)): sLine = Trim$(Mid$(sLine, CPos + 1))
    
    Dim ID As Long:   ID = CLng(sLine)
    
    Dim pl As Player: Set pl = MApp.Players.Item(CStr(ID))
    If Not pl Is Nothing Then
        Set ps_out = MNew.PlayerScore(pl)
        ps_out.Score = Sc
    End If
    TryRead_PlayerScore = True
    Exit Function
Catch:
    MsgBox "Error during reading PlayerScore : " & ID & ", " & Sc
End Function

