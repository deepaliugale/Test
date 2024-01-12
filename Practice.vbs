Function ShowFreeSpace(drvPath)
   Dim fso, d, s
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set d = fso.GetDrive(fso.GetDriveName(drvPath))
   s = "Drive " & UCase(drvPath) & " - " 
   s = s & d.VolumeName  & " "
   s = s & "Free Space: " & FormatNumber(d.FreeSpace/1024, 0) 
   s = s & " Kbytes"
   ShowFreeSpace = s

End Function
s = ShowFreeSpace("D:\")
string1= "Deepali is tester"
If (StrComp(string1, "Deepali")) Then
	'WScript.Echo "Pass"
Else
	'WScript.Echo "Fail"
End If

line = "this is the text to look in to, it contains the search pattern"
Set RE = New RegExp
RE.IgnoreCase = True
RE.Pattern = "search.*tern"
If RE.Test(Line) Then 
'WScript.Echo IsNumeric(Line)
Else 
'WScript.echo VarType(Line)
End If
 wuu = Split("Deepali is a tester in ACN", " ")
 For i=0 To Lbound(wuu)
 'WScript.Echo Replace(wuu(i),"Dee","SEE")
 Next

Dim str1: str1 = Null 
Dim str2: str2 = Null
WScript.Echo StrComp(str1, str2, 1)