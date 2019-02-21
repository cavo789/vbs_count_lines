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
' ==============================================================

Option Explicit

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

    ' --------------------------------------------------
    ' Get the list of files with the specified extension
    ' and return a Dictionary object with, for each file,
    ' the number of lines in the file
    '
    ' Parameters :
    '
    ' sFolder : the folder to scan
    ' sExtension : the extension to search (f.i. "txt")
    '
    ' Remark : if files are big, this function can take a while
    ' so just be patient
    '
    ' See documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/Folders.md#countlines
    ' --------------------------------------------------
    Public Function countLines(sFolder, sExtension)

        Dim objDict
        Dim objFile, objContent
        Dim wCountLines

        Set objDict = CreateObject("Scripting.Dictionary")

        If Not Right(sFolder, 1) = "\" Then
            sFolder = sFolder & "\"
        End If

        ' Loop any files
        For Each objFile In objFSO.GetFolder(sFolder).Files

            If (LCase(objFSO.GetExtensionName(objFile.Name)) = sExtension) Then

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

Sub ShowHelp()

    wScript.echo " ==============="
    wScript.echo " = Count_Lines ="
    wScript.echo " ==============="
    wScript.echo ""
    wScript.echo " Please specify the name of the folder and the extension "
    wScript.echo "to scan; f.i.: "
    wScript.echo " " & wScript.ScriptName & " 'C:\Temp\FolderName' txt"
    wScript.echo ""
    wScript.echo "To get more info, please read https://github.com/cavo789/vbs_count_lines"
    wScript.echo ""

    wScript.quit

End sub

Dim cFolders, cHelper
Dim objFSO, objDict, objKey
Dim sFolderName, sFileName, sExtension
Dim wCount, wTotal

    ' Get the first argument (f.i. "C:\Temp\db1.accdb")
    If (wScript.Arguments.Count < 2) Then

        Call ShowHelp

    Else

        ' Get the path specified on the command line, folder to scan
        sFolderName = Wscript.Arguments.Item(0)

        ' and the extension to scan
        sExtension = Wscript.Arguments.Item(1)

        Set cHelper = New clsHelper
        Call cHelper.ForceCScriptExecution()
        Set cHelper = Nothing

        Set objFSO = CreateObject("Scripting.FileSystemObject")

        Set cFolders = New clsFolders

        Set objDict = cFolders.countLines(sFolderName, sExtension)

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
        Set objFSO = Nothing

    End if
