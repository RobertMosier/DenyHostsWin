'##############################################################################
'# FILE: LogRotate.vbs
'# DATE: 2018.03.14
'#   BY: Robert Mosier
'#       Rotate logs by archiving with incremental integer prefix on the file
'#       name.
'# DEPENDENCIES:
'#       HelperLib.vbs
'#       The following settings must be defined prior to class creation;
'#           LR_LOG_PATH
'#           LR_LOG_FILENAME
'#           LR_MAX_ARCHIVES
'##############################################################################
Option Explicit


Class LogRotate
	'User settings
	Private logPath, logFilename, maxArchives
	
	
	'##########################################################################
	'Initialize the class with user settings.
	Public Sub Class_Initialize()
		logPath		= ExpandEnvVars(LR_LOG_PATH)
		logFilename	= ExpandEnvVars(LR_LOG_FILENAME)
		maxArchives	= ExpandEnvVars(LR_MAX_ARCHIVES)
	End Sub 'Class_Initialize
	
	
	'##########################################################################
	'Get the full file path name of an archived log file.
	'Parameters
	'	iNum As Integer:	A number that corresponds to the zero indexed age
	'		of an archived log file. Each time rotation occurs the index of
	'		each archive is incremented until it reaches maxArchives when it is
	'		deleted.
	'Returns
	'	getArchivePathName As String:	The full path and file name to an
	'		archived log file. This file does not necessarily exists.
	Private Function getArchivePathName(iNum)
		getArchivePathName = logPath & "\" & iNum & "." & logFilename
	End Function 'getArchivePathName
	
	
	'##########################################################################
	'If the oldest archived log file possible (determined by MAX_ARCHIVES - 1)
	'exists it will be deleted.
	Private Sub deleteMaxArchive()
		Dim oFso: Set oFso = CreateObject("Scripting.FileSystemObject")
		Dim sFile: sFile = getArchivePathName(maxArchives - 1)

		If oFso.FileExists(sFile) Then
			oFso.DeleteFile(sFile)
		End If
		
		Set oFso = Nothing
	End Sub 'deleteMaxArchive
	
	
	'##########################################################################
	'Age an archived log file by incrementing its file name prefix.
	'Parameters
	'	iNum As Integer:	A number that corresponds to the zero indexed age
	'		of an archived log file. Each time rotation occurs the index of
	'		each archive is incremented until it reaches maxArchives when it is
	'		deleted.
	Private Sub ageArchive(iNum)
		Dim oFso: Set oFso = CreateObject("Scripting.FileSystemObject")
		Dim sFile1: sFile1 = getArchivePathName(iNum)
		Dim sFile2: sFile2 = getArchivePathName(iNum + 1)
		
		If oFso.FileExists(sFile1) And Not oFso.FileExists(sFile2) Then
			oFso.MoveFile sFile1, sFile2
		End If
		
		Set oFso = Nothing
	End Sub 'ageArchive
	
	
	'##########################################################################
	'Archive the log file by prefixing its file name with 0.
	Private Sub archiveLog()
		Const FW_STOP = "%WINDIR%\System32\netsh.exe advfirewall set allprofiles state off"
		Const FW_START = "%WINDIR%\System32\netsh.exe advfirewall set allprofiles state on"
		
		Dim oShell: Set oShell = CreateObject("WScript.shell")
		Dim oFso: Set oFso = CreateObject("Scripting.FileSystemObject")
		Dim sFile1: sFile1 = logPath & "\" & logFilename
		Dim sFile2: sFile2 = getArchivePathName(0)
		
		'Turn the advfirewall off to release the lock on the log file
		oShell.run "cmd /c " & FW_STOP, 0, True
		'Wait a second for the file lock to release
		WScript.Sleep 1000
		
		If oFso.FileExists(sFile1) And Not oFso.FileExists(sFile2) Then
			oFso.MoveFile sFile1, sFile2
		End If
		
		'Turn the advfirewall back on
		oShell.run "cmd /c " & FW_START, 0, True
		
		Set oFso = Nothing
	End Sub 'archiveLog
	
	
	'##########################################################################
	'Rotate the logs by deleting the oldest, aging the archives, & archiving
	'current log file. It is assumed that the logging software will create its
	'own new log if missing.
	Public Sub Rotate()
		Dim i
		
		deleteMaxArchive

		'age all archives which must be done from oldest to newest
		For i = (maxArchives - 1) To 0 Step -1
			ageArchive i
		Next

		archiveLog
	End Sub 'Rotate
	
End Class 'LogRotate
