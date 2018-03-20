'##############################################################################
'# FILE: Test_WinAdvFw_LogParse.vbs
'# DATE: 2018.03.14
'#   BY: Robert Mosier
'#       Tests the WinAdvFw_LogParse class and demonstrates basic usage.
'# DEPENDENCIES:
'#       WinAdvFw_LogParse.vbs
'##############################################################################
Option Explicit


Sub Test_WinAdvFw_LogParse()
	Dim lp: Set lp = New WinAdvFw_LogParse
	Dim sIP
	
	For Each sIP In lp.UniqueIPAddresses
		WScript.Echo sIP & Space(15 - Len(sIP)) & " Hits: " & lp.GetUniqueIP_Hits(sIP)
	Next

	WScript.Echo "Total lines:   " & lp.TotalLines
	WScript.Echo "Matched lines: " & lp.MatchCount
	WScript.Echo "Unique IPs:    " & lp.UniqueIPCount
	WScript.Echo "Log start:     " & lp.LogStart
	WScript.Echo "Log end:       " & lp.LogEnd
End Sub
