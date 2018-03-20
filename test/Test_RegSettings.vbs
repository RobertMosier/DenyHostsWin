'##############################################################################
'# FILE: Test_RegSettings.vbs
'# DATE: 2018.03.19
'#   BY: Robert Mosier
'#       Tests the RegSettings class and demonstrates basic usage.
'# DEPENDENCIES:
'#       RegSettings.vbs
'##############################################################################
Option Explicit


Sub Test_RegSettings()
	Dim rs: Set rs = New RegSettings

	WScript.Echo "DenyHostsWin"
	WScript.Echo "    Current Version:     " & rs.CurrentVersion
	WScript.Echo "    Threshold:           " & rs.Threshold
	WScript.Echo "    Firewall Rules"
	WScript.Echo "        Base Name:       " & rs.FwRuleBaseName
	WScript.Echo "        Group Name:      " & rs.FwRuleGroupName
	WScript.Echo "        Description:     " & rs.FwRuleDescription
	WScript.Echo "        Local Ports:     " & rs.FwRuleLocalPorts
	WScript.Echo "    Firewall Log"
	WScript.Echo "        Log File:        " & rs.FwLogFile
	WScript.Echo "        Port:            " & rs.FwPort
	WScript.Echo "    Email Notify"
	WScript.Echo "        On Threshold:    " & rs.NotifyOnThreshold
	WScript.Echo "        Threshold:       " & rs.NotifyThreshold
	WScript.Echo "        Email From:      " & rs.NotifyEmailFrom
	WScript.Echo "        Email To:        " & rs.NotifyEmailTo
	WScript.Echo "        Email Subject:   " & rs.NotifyEmailSubject
	WScript.Echo "        Template File:   " & rs.NotifyEmailMsgTemplate
	WScript.Echo "        SMTP Auth User:  " & rs.NotifySMTPAuthUser
	WScript.Echo "        SMTP Auth Pass:  " & rs.NotifySMTPAuthPass
	WScript.Echo "    Log Rotate"
	WScript.Echo "        Log Path:        " & rs.LrLogPath
	WScript.Echo "        Log Filename:    " & rs.LrLogFilename
	WScript.Echo "        Max Archives:    " & rs.LrMaxArchives
End Sub
