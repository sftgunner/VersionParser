Attribute VB_Name = "Module1"
Function VER2NUM(version As String)
'Allows version numbers (eg 12.7.2) to be sorted numerically by converting to a number. Works until a revision number exceeds 1 million
   Dim versionsplit As Variant
   Dim revisiondefinition As Integer
   Dim revisionsegmentvalue As Variant
   Dim versioninteger As Double
   Dim runtime As Integer
   
   versionsplit = Split(version, ".")
   versioninteger = 0
   runtime = 1
   
   For Each revisionsegmentvalue In versionsplit
   'NB: VBA arrays start from 0
   
   versioninteger = versioninteger + (revisionsegmentvalue * (10 ^ ((runtime * -1) + 1 - (5 * (runtime - 1)))))
   runtime = runtime + 1
   Next
   
   VER2NUM = versioninteger
End Function
