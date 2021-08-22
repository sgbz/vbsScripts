On Error Resume Next

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

query = MsgBox("Scan current directory?", vbYesNo)

If query = 6 Then
	folderPATH = fso.GetAbsolutePathName("./")
Else
	folderPATH = InputBox("Paste directory path.", "")
End If

If fso.FolderExists(folderPATH) Then
  
	OutStr = folderPATH + vbCrLf
	OutStr = OutStr + AllFolders(folderPATH)

	Function inArray(element, arr)
		For i = 0 To Ubound(arr) 
			If Trim(arr(i)) = Trim(element) Then 
				inArray = True
				Exit Function
			Else 
				inArray = False
			End If  
		Next 
	End Function

	Function convertSize(value)
		i = 0
		unit = Array("b","Kb","Mb","Gb","Tb", "Pb", "Eb", "Zb", "Yb")

		do until value < 1024
			i = i +1
			value = value / 1024
		loop
	
		convertSize = round(value, 2)&" "& unit(i)
	End Function

	Function AllFolders(WDir)
    	Output = ""
    	Set SubF = fso.GetFolder(WDir).SubFolders

    	For Each Folder In SubF
        	Output = Output + WDir + "\" + Folder.Name + vbCrLf
        	Output = Output + AllFolders(WDir + "\" + Folder.Name)
    	Next 

    	AllFolders = Output
	End Function

	allDirectories = Split(OutStr, vbcrlf)

	Dim allExt()
	Dim ext()

	i = 0

	For Each directory In allDirectories
		For Each file In fso.GetFolder(directory).Files
			ReDim Preserve allExt(i)
			allExt(i) = LCase(fso.GetExtensionName(file.Name))
			totalSize = totalSize + file.Size
			i = i + 1
		Next
	Next

	If UBound(allExt) < 1 Then
		MsgBox "No files in current directory."
		WScript.Quit
	End If

	For r = 0 To UBound(allExt)
		ReDim Preserve ext(r)
	
		If inArray(allExt(r), ext) Then
				ext(r) = null
		Else
				ext(r) = allExt(r)
		End If
	
		If ext(r) <> "" Then
			result = ext(r)&" "&result
			extensionsCount = extensionsCount + 1	
		End If	
	next

	extArray = Split(Trim(result), " ")

	fso.CreateTextFile("extensions.txt").Write("Directory: """ & folderPATH & """."& vbCrLf)
	Set report = fso.OpenTextFile("extensions.txt", 8)

	report.WriteLine("Total extension types in this folder: " & extensionsCount &" ("& Trim(LCase(result)) &")."& vbCrLf)


	num = 0
	For Each extension In extArray
		filesCount = 0
		num = num + 1
		For Each directory In allDirectories
			For Each file In fso.GetFolder(directory).Files
				If LCase(fso.GetExtensionName(file.name)) = extension then
					filesCount = filesCount + 1
					size = size + file.Size
					fileType = fso.GetFile(fso.BuildPath(directory, file.name)).Type
				End If
			Next
		Next
		report.WriteLine(num &") "& fileType & " (" & LCase(extension) &") - " &" Total: "& filesCount-1 & ", Size: " & convertSize(size) & ".")
	Next

	report.WriteLine(vbcrlf & "Files in total: " & UBound(allExt) & ", total size: " & convertSize(totalSize) & ".")
	report.WriteLine("Scan date and time: " & WeekdayName(Weekday(now())) & ", " & Now() & ".")
	report.Close
	
	MsgBox "Scanning complete."

Else
	MsgBox "Directory does not exist."
End If
