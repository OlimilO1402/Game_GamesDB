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
Public IdNr As Long
'Eine ID deswegen, weil der gleiche Name
'mehrmals auftauchen könnte

Public Name As String
'Man könmnte hier noch mehr sonstige personenbezogene Daten speichern
'jetzt mal nur Vornamen

Public Function ToString() As String
  ToString = CStr(IdNr) & ", " & Name 'und Sonstiges
End Function
