Attribute VB_Name = "Module1"
Function VER2NUM(version As String)
'Allows version numbers (eg 12.7.2) to be sorted numerically by converting to a number. Works until a revision number exceeds 2^10
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
   
   versioninteger = versioninteger + (revisionsegmentvalue * (2 ^ ((runtime * -1) + 1 - (10 * (runtime - 1)))))
   runtime = runtime + 1
   Next
   
   VER2NUM = versioninteger
End Function

Function NUM2VER(number As Double)
    Dim versionarray() As String
    Dim revisionsegmentvalue As Integer
    i = 0
    Do While number > 0
    ReDim Preserve versionarray(i)
    versionarray(i) = CStr(Round(number))
    number = (number - Round(number)) * (2 ^ 11)
    i = i + 1
    Loop
    
    NUM2VER = Join(versionarray, ".")
End Function
