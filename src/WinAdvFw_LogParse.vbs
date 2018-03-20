'##############################################################################
'# FILE: WinAdvFw_LogParse.vbs
'# DATE: 2018.03.06
'#   BY: Robert Mosier
'#       Parses file defined by LOG_FILE. Creates comprehensive listing of
'#       unique IP addresses. Exposes date, time, IP, and count for each IP.
'# DEPENDENCIES:
'#       HelperLib.vbs
'#       The following settings must be defined prior to class creation;
'#           FW_LOG_FILE
'#           FW_PORT
'##############################################################################
Option Explicit


'##############################################################################
'This class is instantiated for each unique IP matching in the log and the object
'references are maintained within the fwSuspects Dictionary object in
'WinAdvFw_LogParse class.
Class Suspect
	'Public properties?!? Shame on me, I know I'm just too lazy!
	Public count	'number of times this IP was matched
	Public first	'date/time of first match
	Public last		'date/time of last match
End Class


Class WinAdvFw_LogParse

	'User settings
	Private fwLogFile, fwPort
	
	'A Dictionary object for the IPv4 addresses
	Private fwSuspects
	
	'Variables used for summary (these may not even get used)
	Private fwTotalLines, fwMatchCount, fwLogStart, fwLogEnd

	
	'##########################################################################
	'Properties for summary (read-only publicly)
	Public Property Get TotalLines
		TotalLines = fwTotalLines
	End Property

	Public Property Get MatchCount
		MatchCount = fwMatchCount
	End Property

	Public Property Get LogStart
		LogStart = fwLogStart
	End Property

	Public Property Get LogEnd
		LogEnd = fwLogEnd
	End Property

	
	'##########################################################################
	'Class constants implemented as private read-only properties for self
	'documenting code. It beats Integer values sprinkled around the code with
	'little to no meaning right?
	'Format of log entries within the file
	'#Fields: date time action protocol src-ip dst-ip src-port dst-port size tcpflags tcpsyn tcpack tcpwin icmptype icmpcode info path
	Private Property Get FIELD_DATE
		FIELD_DATE = 0
	End Property
	
	Private Property Get FIELD_TIME
		FIELD_TIME = 1
	End Property
	
	Private Property Get FIELD_SRC_IP
		FIELD_SRC_IP = 4
	End Property
	
	Private Property Get FIELD_DST_PORT
		FIELD_DST_PORT = 7
	End Property

	
	'##########################################################################
	'Wrapper properties for fwSuspects to expose Dictionary object. Again read-only.
	Public Property Get UniqueIPCount
		UniqueIPCount = fwSuspects.Count
	End Property
	
	Public Property Get UniqueIPAddresses
		UniqueIPAddresses = fwSuspects.Keys
	End Property
	
	Public Function GetUniqueIP_Hits(ipv4)
		GetUniqueIP_Hits = fwSuspects.Item(ipv4).count
	End Function
	
	Public Function GetUniqueIP_First(ipv4)
		GetUniqueIP_First = fwSuspects.Items(ipv4).first
	End Function
	
	Public Function GetUniqueIP_Last(ipv4)
		GetUniqueIP_Last = fwSuspects.Items(ipv4).last
	End Function

	
	'##########################################################################
	'Initialize the class with user defined settings.
	'See Settings.vbs for more detailed info.
	Public Sub Class_Initialize()
		'User configurable settings
		fwLogFile	= ExpandEnvVars(FW_LOG_FILE)
		fwPort		= ExpandEnvVars(FW_PORT)
		
		'Dictionary object to contain the Suspect objects
		Set fwSuspects = CreateObject("Scripting.Dictionary")
		
		'Variables used for summary
		fwTotalLines = 0
		fwMatchCount = 0
		
		'Read the log file so suspect IP addresses can be identified
		ParseLogFile
	End Sub 'Class_Initialize

	
	'##########################################################################
	'Add Suspect objects to the dictionary of fwSuspects.
	'Parameters
	'	matchedLineArray As Array:	an array of Strings split from a matched
	'		log entry line which should be 16 items for valid log entries.
	'		See the read-only properties above beginning with FIELD_ for more
	'		info.
	Private Sub AddSuspect(matchedLineArray)
		Dim objSuspect
		
		'Have we encountered this IP before?
		If fwSuspects.Exists(matchedLineArray(FIELD_SRC_IP)) Then
			'Lets work with the existing suspect object from the dictionary
			Set objSuspect = fwSuspects(matchedLineArray(FIELD_SRC_IP))
			
			'Update the suspect properties
			objSuspect.count = objSuspect.count + 1
			objSuspect.last = matchedLineArray(FIELD_DATE) & " " & matchedLineArray(FIELD_TIME)
			
			'Update the dictionary item to reflect our changes
			Set fwSuspects.Item(matchedLineArray(FIELD_SRC_IP)) = objSuspect
		Else
			'This is a unique IP so lets make a new suspect object
			Set objSuspect = new Suspect
			
			'Set the suspect properties accordingly
			objSuspect.first = matchedLineArray(FIELD_DATE) & " " & matchedLineArray(FIELD_TIME)
			objSuspect.last = objSuspect.first
			objSuspect.count = 1
			'objSuspect.ip = matchedLineArray(FIELD_SRC_IP)
			
			'Insert the suspect object into the dictionary of fwSuspects
			'fwSuspects.Add objSuspect.ip, objSuspect
			fwSuspects.Add matchedLineArray(FIELD_SRC_IP), objSuspect
		End If
	End Sub 'AddSuspect

	
	'##########################################################################
	'Add Suspect objects to the dictionary of fwSuspects.
	Private Sub ParseLogFile()
		'Variables for reading in the log file defined by fwLogFile
		Dim fso:	Set fso	= CreateObject("Scripting.FileSystemObject")
		Dim f:		Set f	= fso.OpenTextFile(fwLogFile)
		Dim logLine, splitLine

		'Loop for each line in the fwLogFile. Only ends when we reach the end of file
		Do Until f.AtEndOfStream
			'Read the next line from the log file
			logLine = f.ReadLine
			
			'Skip blank lines
			If Len(logLine) > 0 Then
			
				'Skip commented lines
				If Left(logLine, 1) <> "#" Then
				
					'Split log entries into respective components
					splitLine = Split(logLine, " ")
					
					'Get the first line date/time
					If fwTotalLines = 0 Then
						fwLogStart = splitLine(FIELD_DATE) & " " & splitLine(FIELD_TIME)
					End If
					
					'Is logLine a valid log entry?
					If UBound(splitLine) = 16 Then
					
						'Add valid log entry lines to fwTotalLines
						fwTotalLines = fwTotalLines + 1
						
						'Is the log entry a match for the port
						If splitLine(FIELD_DST_PORT) = fwPort Then
							'Increment fwMatchCount so we have a count of all lines matched by port
							fwMatchCount = fwMatchCount + 1
							
							'save the date/time of each line so we have the last left after
							fwLogEnd = splitLine(FIELD_DATE) & " " & splitLine(FIELD_TIME)
							
							'Finally we actually add the suspect object
							AddSuspect splitLine
						End If
					End If
				End If
			End If
		Loop

		f.Close
	End Sub 'ParseLogFile

End Class 'WinAdvFw_LogParse
