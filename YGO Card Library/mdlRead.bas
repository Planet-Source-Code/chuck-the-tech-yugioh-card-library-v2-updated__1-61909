Attribute VB_Name = "mdlRead"
'needed by ReadINI Function
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'needed by WriteINI Function
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long

'User defined type to set values for each card
Public Type Card
  'sets the Frame for use with sorting cards
  Frame As Integer
  'Card name
  Name As String
  'Card Attribute
  Attribute As Integer
  'Card Icon if Spell or Trap Card
  Icon As String
  'Card Type
  Type As String
  'Card Description
  Description As String
  'Monster Level
  Level As Integer
  'number of tribute(s) required to summon monster
  Cost As Integer
  'monster's attack points
  Attack As Integer
  'monster's defence points
  Defence As Integer
  'These 4 will be used for card effects.  More might be added
  Phase As String
  Spell As String
  Value As String
  Value2 As String
End Type

Public Function ReadINI(Section, KeyName, filename As String) As String
Dim sRet As String
  
  sRet = String(998, Chr(0))
  ReadINI = Left(sRet, GetPrivateProfileString(Section, KeyName, "", sRet, Len(sRet), filename))
End Function

Public Function WriteINI(Section, KeyName, NewString As String, filename As String) As String
  Dim sWet As String
  sWet = WritePrivateProfileString(Section, KeyName, NewString, filename)
End Function


'Gets the card info
Public Function Get_Card(strCard As String) As Card
Dim vardata As Variant

  vardata = Split(ReadINI("Cards", strCard, App.Path & "\data\cards.dat"), "|")
  
On Error GoTo nocard
  
  With Get_Card
    .Frame = vardata(0)
    .Name = vardata(1)
    .Attribute = vardata(2)
    .Icon = vardata(3)
    .Type = vardata(4)
    .Description = vardata(5)
    If vardata(6) = "" Then .Level = 0 Else .Level = vardata(6)
    If vardata(7) = "" Then .Cost = 0 Else .Cost = vardata(7)
    If vardata(8) = "" Then .Attack = -1 Else .Attack = vardata(8)
    If vardata(9) = "" Then .Defence = -1 Else .Defence = vardata(9)
  End With

Exit Function

'used if card not found
nocard:
    Get_Card.Name = ""
End Function

Public Function Get_Card_Small(strCard As Integer) As Card
Dim vardata As Variant

  vardata = Split(ReadINI("Cards", strCard, App.Path & "\data\cards.dat"), "|")

  With Get_Card_Small
    .Frame = vardata(0)
    .Name = vardata(1)
    .Attribute = vardata(2)
    .Icon = vardata(3)
    .Type = vardata(4)
    .Description = vardata(5)
    If vardata(6) = "" Then .Level = 0 Else .Level = vardata(6)
    If vardata(7) = "" Then .Cost = 0 Else .Cost = vardata(7)
    If vardata(8) = "" Then .Attack = -1 Else .Attack = vardata(8)
    If vardata(9) = "" Then .Defence = -1 Else .Defence = vardata(9)
  End With

End Function



