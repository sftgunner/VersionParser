Attribute VB_Name = "VersionParser"
'VersionParser Module

Function VER2NUM(ByVal version As String, Optional ByVal revisionprecision As Integer = 4)
'Optimum revisionprecision = 4 -> allows for version numbers up to 16, and depth of 10 subversions
'Previous default = 10 -> allows for version numbers up to 1024 and depth of 4.
'Maximum total revisions (for all subversions combined > 1.099 trillion subversions) irrespective of revisionprecision. Simply chooses whether focus is on depth or breadth.
'For standard x.x.x versioning, revisionprecision can be set to 25, allowing for only 847 billion conbinations, 33 million per subversion (as 3 not a power of two)
'Allows version numbers (eg 12.7.2) to be sorted numerically by converting to a number.
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
   
   versioninteger = versioninteger + (revisionsegmentvalue * (2 ^ ((runtime * -1) + 1 - (revisionprecision * (runtime - 1)))))
   runtime = runtime + 1
   Next
   
   VER2NUM = versioninteger
End Function

Function NUM2VER(ByVal number As Double, Optional ByVal revisionprecision As Integer = 4)
    Dim versionarray() As String
    Dim revisionsegmentvalue As Integer
    i = 0
    Do While number > 0
    ReDim Preserve versionarray(i)
    versionarray(i) = CStr(Round(number))
    number = (number - Round(number)) * (2 ^ (revisionprecision + 1))
    i = i + 1
    Loop
    
    NUM2VER = Join(versionarray, ".")
End Function
