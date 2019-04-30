' ==============================================================
'
' Author : Christophe Avonture
' Date	 : June 2018
'
' Open every .txt files under the current directory and count the
' number of lines in each files
'
' Display the informations as a markdown table
'
' CHANGES
' 20190425 - Improve parameters by allowing "." to specify the
'               current folder and allowing multiple extensions
'               like "bas;cls;csv;php;xml"
'
' ==============================================================

Option Explicit

Dim sFolderName, sExtensions
Dim objFSO
Dim bShowHelp

Class clsHelper

	Sub ForceCScriptExecution()

		Dim sArguments, Arg, sCommand

		If Not LCase(Right(WScript.FullName, 12)) = "\cscript.exe" Then

			' Get command lines paramters'
			sArguments = ""
			For Each Arg In WScript.Arguments
				sArguments=sArguments & Chr(34) & Arg & Chr(34) & Space(1)
			Next

			sCommand = "cmd.exe cscript.exe //nologo " & Chr(34) & _
			WScript.ScriptFullName & Chr(34) & Space(1) & Chr(34) & sArguments & Chr(34)

			' 1 to activate the window
			' true to let the window opened
			Call CreateObject("Wscript.Shell").Run(sCommand, 1, true)

			' This version of the script (started with WScript) can be terminated
			wScript.quit

		End If

	End Sub

End Class

Class clsFolders

	Dim objFSO, objFile

	Private bVerbose

	Public Property Let verbose(bYesNo)
		bVerbose = bYesNo
	End Property

	Private Sub Class_Initialize()
		bVerbose = False
		Set objFSO = CreateObject("Scripting.FileSystemObject")
	End Sub

	Private Sub Class_Terminate()
		Set objFSO = Nothing
	End Sub

	' @url https://gist.github.com/sholsinger/943116/caf67a2504d6e45e4acc49597fac5f1bb6033ba2#gistcomment-1967571
	Private Function in_array(needle, haystack)

		Dim hay

		in_array = False

		needle = trim(needle)

		For Each hay in haystack
			If trim(hay) = needle Then
				in_array = True
				Exit For
			End If
		Next

	End Function

	' @url https://github.com/zeirishin/phpvbs/blob/master/functions/strings/explode.vbs
	Private Function explode(delimiter,string,limit)

		explode = false

		If len(delimiter) = 0 Then Exit Function

		If len(limit) = 0 Then limit = 0

		If limit > 0 Then
			explode = Split(string,delimiter,limit)
		Else
			explode = Split(string,delimiter)
		End If

	End Function

	' --------------------------------------------------
	' Get the list of files with the specified extension
	' and return a Dictionary object with, for each file,
	' the number of lines in the file
	'
	' Parameters :
	'
	' sFolder : the folder to scan
	' sExtensions : the extension(s) to search (f.i. "txt" or "csv;txt")
	'
	' Remark : if files are big, this function can take a while
	' so just be patient
	'
	' See documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/Folders.md#countlines
	' --------------------------------------------------
	Public Function countLines(sFolder, sExtensions)

		Dim objDict
		Dim objFile, objContent
		Dim wCountLines
		Dim arrExtensions

		sExtensions = Trim(sExtensions)

		' Remove final ";" if there are ones
		If (Right(sExtensions, 1) = ";") Then
			Do
				sExtensions = Left(sExtensions, Len(sExtensions) - 1)
			Loop While (Right(sExtensions, 1) = ";") or (sExtensions = "")

		End if

		' Convert the list (f.i. bas;cls;csv;json) into an array
		arrExtensions = explode(";", sExtensions, 0)

		Set objDict = CreateObject("Scripting.Dictionary")

		If Not Right(sFolder, 1) = "\" Then
			sFolder = sFolder & "\"
		End If

		' Loop any files
		For Each objFile In objFSO.GetFolder(sFolder).Files

			If (in_array(LCase(objFSO.GetExtensionName(objFile.Name)), arrExtensions)) Then

				Set objContent = objFSO.OpenTextFile(sFolder & objFile.Name, 1)

				objContent.ReadAll

				wCountLines = objContent.Line

				objdict.Add objFile.Name, wCountLines

			End if

		Next

		Set objContent = Nothing
		Set objFile = Nothing

		Set countLines = objdict

	End Function

	' --------------------------------------------------
	' Return the current folder i.e. the folder from where
	' the script has been started
	'
	' See documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/Folders.md#getcurrentfolder
	' --------------------------------------------------
	Public Function getCurrentFolder()

		Dim sFolder

		Set objFile = objFSO.GetFile(Wscript.ScriptName)
		sFolder = objFSO.GetParentFolderName(objFile) & "\"
		Set objFile = Nothing

		getCurrentFolder = sFolder

	End Function

End Class

' --------------------------------------------------------
'
' Variables initialization
'
' --------------------------------------------------------
Private Sub initialization()

	bShowHelp = False

	' Initialize our parameters
	sFolderName = ""
	sExtensions = ""

	Set objFSO = CreateObject("Scripting.FileSystemObject")

End Sub

' --------------------------------------------------------
'
' Get the folder parameter from command line options
' Make sure folder name ends with a \
'
' --------------------------------------------------------
Private Sub getFolder(sArgument)

   Dim wshShell

	sFolderName = sArgument

	If Trim(sFolderName) = "." Then
		Set wshShell = CreateObject("WScript.Shell")
		sFolderName = wshShell.CurrentDirectory
		Set wshShell = Nothing
	End If

	' Path should be absolute

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	sFolderName = objFSO.GetAbsolutePathName(sFolderName)
	Set objFSO = Nothing

	On Error Resume next

	' Oups, the folder doesn't exists
	If Err.Number <> 0 Then
		sFolderName = ""
		Err.clear
	End If

	On Error Goto 0

	If (sFolderName <> "") Then
		If Not Right(sFolderName, 1) = "\" Then
			sFolderName = sFolderName & "\"
		End If
	End if

End Sub

' --------------------------------------------------------
'
' Get the list of extensions to scan (f.i. "csv;txt")
'
' --------------------------------------------------------
Private Sub getExtensions(sArgument)
	sExtensions = Trim(sArgument)
End Sub

' --------------------------------------------------------
'
' Finalization
'
' --------------------------------------------------------
Private Sub finalize()

	Set objFSO = Nothing

End Sub

' --------------------------------------------------------
'
' Helper, process command line parameters / options
'
'   Parameters starts with a "-" (like in "-i=")
'   Options starts with a "/" (like in "/force")
'
' --------------------------------------------------------
Private Sub getParameters()

	Dim wCount, I
	Dim sArgument

	if (wScript.Arguments.Count = 0) Then
		bShowHelp = True
		Exit Sub
	End If

	wCount = wScript.Arguments.Count - 1

	' Process arguments one by one
	For I = 0 To wCount

		' Get the argument
		sArgument = Trim(wScript.Arguments(I))

		If (Left(sArgument, 3) = "-i=") Then
			' -i is for the name of the input folder
			Call getFolder(Right(sArgument, Len(sArgument) - 3))
		ElseIf (Left(sArgument, 3) = "-e=") Then
			' -e is for the list of extensions
			Call getExtensions(Right(sArgument, Len(sArgument) - 3))
		Else
			If (sArgument = "/?") Then
				bShowHelp = true
			ElseIf (LCase(sArgument) = "/help") Then
				bShowHelp = true
			End If
		End If

	Next

End Sub

' --------------------------------------------------------
'
' Somes parameters are mandatory and can't be missing. If one
' of them isn't specified on the command line.
' Quit the script
'
' --------------------------------------------------------
Private Sub validateParameters()

	' Start validation, these variables can't be empty
	If (Trim(sFolderName) = "") Or (Trim(sExtensions) = "") Then

		If (Trim(sFolderName) = "") Then
			wScript.echo "Error - You need to specify the folder to " & _
				"scan for getting files and count lines."
			wScript.echo ""
			wScript.echo "Please use the -i option for this purpose."
		ElseIf (Trim(sExtensions) = "") Then
			wScript.echo "Error - You need to specify file's extensions " & _
				"to scan (f.i. ""php"")."
			wScript.echo ""
			wScript.echo "Please use the -e option for this purpose."
		End if

		wScript.echo ""
		wScript.echo "In case of need, run this script with the /? " & _
			"parameters to get help"
		wScript.echo ""

		' -1 = Something goes wrong
		wScript.Quit -1

	End If

End Sub

' --------------------------------------------------------
'
' Display how to use the script and the list of parameters
' Quit the script
'
' --------------------------------------------------------
Sub ShowHelp()

	wScript.echo "==============="
	wScript.echo "= Count_Lines ="
	wScript.echo "==============="
	wScript.echo ""
	wScript.echo "Scan a folder and count the number of text lines in files."
	wScript.echo ""
	wScript.echo "Usage: " & wScript.ScriptName & " -i=input_folder -e=csv;txt"
	wScript.echo ""
	wScript.echo "-i=xxx    Name of the folder to scan (or just a '.' for the current folder)"
	wScript.echo "-e=xxx    One or more extensions (f.i. 'txt' or 'csv;txt;xml')"
	wScript.echo ""
	wScript.echo "/?        Show this help screen (or /help)"
	wScript.echo ""
	wScript.echo "Examples: "
	wScript.echo "    cscript " & Wscript.ScriptName & " -i=. -e=csv"
	wScript.echo "    cscript " & Wscript.ScriptName & " -i=. -e=csv;txt;xml"
	wScript.echo "    cscript " & Wscript.ScriptName & " -i=C:\Data\ -e=csv"
	wScript.echo ""
	wScript.echo " To get more info, please read https://github.com/cavo789/vbs_count_lines"

	' And quit
	wScript.Quit -1

End sub

' --------------------------------------------------------
'
' ENTRY POINT
'
' --------------------------------------------------------

Dim cFolders, cHelper
Dim objDict, objKey
Dim wCount, wTotal
Dim sFileName

	Call initialization()

	Call getParameters()

	' Show help screen and stop
	If bShowHelp Then Call showHelp()

	' Run validations and in case of problem, display an error message and stop
	Call validateParameters

	Set cHelper = New clsHelper
	Call cHelper.ForceCScriptExecution()
	Set cHelper = Nothing

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	Set cFolders = New clsFolders

	Set objDict = cFolders.countLines(sFolderName, sExtensions)

	wTotal = 0

	wScript.echo "| Filename | Count |" & vbCrLf & "| --- | --- |"

	For Each objKey In objDict

		sFileName = objKey

		wCount = objDict(objKey)

		' The last line is an empty one, don't count it
		wCount = wCount - 1

		' Add and calculate a total
		wTotal = wTotal + wCount

		wScript.echo "| " & sFileName & " | " & FormatNumber(wCount, 0) & " |"

	Next

	wScript.echo "| TOTAL | " & FormatNumber(wTotal, 0) & " |"

	wScript.echo ""

	Set cFolders = Nothing

	Call finalize()
