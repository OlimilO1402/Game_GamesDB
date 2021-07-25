Attribute VB_Name = "MSerial"
Option Explicit
Private sb As String

Public Sub Clear()
    sb = vbNullString
End Sub

Public Function ToStr() As String
    ToStr = sb
End Function

Public Sub Player(Obj As Player)
    Dim s As String
    With Obj
        s = s & .ToStr & vbCrLf 'CStr(.ID) & ", " & .Name
    End With
    sb = sb & s
End Sub

'Public Function ToStr() As String
'    ToStr = CStr(IDNr) & ", " & Name 'und Sonstiges
'End Function
'Try: On Error GoTo Finally
'    Dim mbr As VbMsgBoxResult    'mbr = vbOK
'    If mCol.Count = 0 Then mbr = MsgBox("Die Liste AllPlayers ist leer, soll die Datei überschrieben werden?", vbOKCancel)
'    If mbr = vbCancel Then Exit Sub
'    Dim FNr As Integer: FNr = FreeFile
'    Open aPFN For Output As #FNr
'    Dim Pl As Player
'    For Each Pl In m_Col
'        Print #FNr, Pl.ToStr
'    Next
'Finally:
'    Close FNr
'End Function

Public Sub Players(Objs As Players)
    Dim s As String
    With Objs
        Dim Obj As Player
        For Each Obj In .List
            MSerial.Player Obj
            's = s & vbCrLf
        Next
    End With
    sb = sb & s
End Sub

'Public Sub PlayerScore(Obj As PlayerScore)
'    Dim s As String: s = Obj.ToStr(True)
'    sb = sb & s
'End Sub

Public Sub Game(g As Game)
    Dim s As String
    With g
        s = .ToStr & IIf(.List.Count > 0, ", ", "")
        Dim ps As PlayerScore
        Dim i As Long
        For i = 1 To .List.Count - 1
            Set ps = .List.Item(i)
            s = s & ps.ToStr(True) & ", "
        Next
        Set ps = .List.Item(.List.Count)
        s = s & ps.ToStr(True)
    End With
    sb = sb & s
End Sub

Public Sub Games(gs As Games)
    With gs
        Dim g As Game
        For Each g In .List
            MSerial.Game g
            sb = sb & vbCrLf
        Next
    End With
End Sub

