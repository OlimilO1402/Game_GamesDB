VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Eine ID deswegen, weil der gleiche Name
'mehrmals auftauchen könnte
Private m_ID As Long
Public Name  As String  'der Name kann nachträglich editiert werden
'Man könmnte hier noch mehr sonstige personenbezogene Daten speichern
'jetzt mal nur Vornamen

Friend Sub New_(ByVal aID As Long, aName As String)
    m_ID = aID: Name = aName
End Sub

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Function ToStr() As String
    ToStr = CStr(m_ID) & ", " & Name 'und Sonstiges
End Function
