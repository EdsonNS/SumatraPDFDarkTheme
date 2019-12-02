

sVBSPath = Left(WScript.ScriptFullName,(Len(WScript.ScriptFullName) - (Len(WScript.ScriptName) + 1)))
REM msgbox sVBSPath

Dim fileTXT : fileTXT = sVBSPath & "\" & "SumatraPDF-settings.txt"

Dim strMod : strMod = WScript.Arguments.Item(0)

Const LightTheme = "LightTheme"
Const DarTheme = "DarkTheme"

If (strMod=DarTheme) Then
	strSearch1 = "BackgroundColor = #ffffff"
	strReplace1 = "BackgroundColor = #6e6e6e"
	strSearch2 = "TextColor = #000000"
	strReplace2 = "TextColor = #ffffff"
End If
If (strMod=LightTheme) Then
	strSearch1 = "BackgroundColor = #6e6e6e"
	strReplace1 = "BackgroundColor = #ffffff"
	strSearch2 = "TextColor = #ffffff"
	strReplace2 = "TextColor = #000000"
End If


Const ForReading = 1
Const ForWriting = 2
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(fileTXT, ForReading)
strText = objFile.ReadAll
objFile.Close
strNewText = Replace(strText, strSearch1, strReplace1)
strNewText = Replace(strNewText, strSearch2, strReplace2)
Set objFile = objFSO.OpenTextFile(fileTXT, ForWriting)
objFile.WriteLine strNewText
objFile.Close