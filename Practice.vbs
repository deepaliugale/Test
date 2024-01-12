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
's = ShowFreeSpace("D:\")
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
'WScript.Echo StrComp(str1, str2, 1)

Set objDic= CreateObject("Scripting.Dictionary")

objDic.Add "Item" , 4
'WScript.Echo objDic.Count
'WScript.Echo objDic.Exists(4)
'WScript.Echo objDic.Item("Item")


'factorialValue = 1

'userInput = Inputbox ("Enter any number")

'For i = 1 to userInput

'    factorialValue = factorialValue * i 
'Msgbox  factorialValue
'Next 
origStr = "efqieufge933"
Function singleCharArr(inputString)
If inputString <> "" Then
      Dim oRegEx, colChars
      Set oRegEx = New RegExp
      oRegEx.Global = True
      oRegEx.IgnoreCase = True
      oRegEx.Pattern = ".{1}"
      Set colChars = oRegEx.Execute(inputString)
      Dim arrChars()
      For i = 0 To colChars.Count - 1
          Redim Preserve arrChars(i)
          arrChars(i) = colChars.Item(i)
      Next
      Set colChars = Nothing
      Set oRegEx = Nothing
      singleCharArr = arrChars
      Erase arrChars
  End If
End Function

strAry = singleCharArr(Trim(origStr))
Dim counts(20) 
For i = 0 To UBound(StrAry) - 1
    For j = 0 To UBound(StrAry) - 1
        If (StrAry(i) = StrAry(j)) Then
            cnt =cnt+1
        End If
        counts(i) = cnt
    Next
    cnt = 0
    If counts(i)>1 Then
    	'WScript.Echo(strAry(i) &" occurred more than one => " & counts(i))
    Else
    	'WScript.Echo(strAry(i)&" occurred once => " & counts(i))
    End If
Next


