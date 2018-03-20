'##############################################################################
'# FILE: DenyHostsWin.vbs
'# DATE: 2018.03.14
'#   BY: Robert Mosier
'#       The main class of the application. This software is intended to find
'#       potential remote attackers in the Windows Advanced Firewall log and
'#       block their IP addresses from specified network ports.
'# DEPENDENCIES:
'#       HelperLib.vbs
'#       WinAdvFw.vbs
'#       WinAdvFw_LogParse.vbs
'#       IPv4Addr.vbs
'#       GMailNotify.vbs
'##############################################################################
Option Explicit


Sub Main
	Dim sIP, oIP
	Dim aAddrs
	Dim lp, fw, lr, notify
	
	'Read settings from Registry
	InitSettings
	
	'Create LogRotate
	Set lr = New LogRotate
	'Lets rotate the logs first so we don't miss any entries from the firewall
	lr.Rotate
	'We are done with LogRotate so clean up
	Set lr = Nothing

	'Create WinAdvFw_LogParse
	Set lp = New WinAdvFw_LogParse
	'Create WinAdvFw
	Set fw = New WinAdvFw
	'Create GMailNotify
	Set notify = New GMailNotify
	
	If lp.UniqueIPCount > 0 Then
	
		'Check each IP if it has more than THRESHOLD hits
		For Each sIP In lp.UniqueIPAddresses
			If lp.GetUniqueIP_Hits(sIP) > THRESHOLD Then
			
				'Create the IPv4Addr object so we can validate addresses
				Set oIP = New IPv4Addr
				If oIP.ValidIPv4(sIP) Then
				
					'Since this IP is over the THRESHOLD we should block it in the firewall
					fw.BlockIPv4Address sIP
					
					'Add the blocked IP to the array of IPv4Addr
					oIP.SetAddress sIP
					ArrayPush aAddrs, oIP
				End If
			End If
		Next
		
		If IsArray(aAddrs) Then
			'Sort the IP addresses for easily identifying trends like
			'multiple addresses within same subnet
			SortAddrs aAddrs
			
			'Place our found values in the email message template and send it
			notify.preprocessMessage aAddrs
			notify.Send()
		End If
		
	End If
	
	'Clean up
	Set fw = Nothing
	Set lp = Nothing
	Set notify = Nothing
	
	'TODO:	It would also be nice to setup some error/debug logging.
End Sub


Sub SortAddrs(aIPv4)
	Dim i, j, TempValue

	For i = LBound(aIPv4) To UBound(aIPv4)
		For j = LBound(aIPv4) To UBound(aIPv4) - 1
		
			If aIPv4(j).ToInteger > aIPv4(j + 1).ToInteger Then
				Set TempValue = aIPv4(j + 1)
				Set aIPv4(j + 1) = aIPv4(j)
				Set aIPv4(j) = TempValue
			End If
			
		Next
	Next
End Sub


Sub InitSettings
	'Declare and create RegSettings class
	Dim rs: Set rs = New RegSettings

	'Read settings from Registry into global variables.
	'These used to be constants but are now variables,
	'just have not changed the names to match convention yet.
	THRESHOLD					= rs.Threshold
	FW_RULE_BASE_NAME			= ExpandEnvVars(rs.FwRuleBaseName)
	FW_RULE_GROUP_NAME			= ExpandEnvVars(rs.FwRuleGroupName)
	FW_RULE_DESCRIPTION			= ExpandEnvVars(rs.FwRuleDescription)
	FW_RULE_LOCAL_PORTS			= ExpandEnvVars(rs.FwRuleLocalPorts)
	FW_LOG_FILE					= ExpandEnvVars(rs.FwLogFile)
	FW_PORT						= ExpandEnvVars(rs.FwPort)
	
	If rs.NotifyOnThreshold > 0 Then
		NOTIFY_ON_THRESHOLD = True
	Else
		NOTIFY_ON_THRESHOLD = False
	End If
	
	NOTIFY_THRESHOLD			= rs.NotifyThreshold
	NOTIFY_EMAIL_FROM			= ExpandEnvVars(rs.NotifyEmailFrom)
	NOTIFY_EMAIL_TO				= ExpandEnvVars(rs.NotifyEmailTo)
	NOTIFY_EMAIL_SUBJECT		= ExpandEnvVars(rs.NotifyEmailSubject)
	NOTIFY_EMAIL_MSG_TEMPLATE	= ExpandEnvVars(rs.NotifyEmailMsgTemplate)
	NOTIFY_SMTP_AUTH_USER		= rs.NotifySMTPAuthUser
	NOTIFY_SMTP_AUTH_PASS		= rs.NotifySMTPAuthPass
	LR_LOG_PATH					= ExpandEnvVars(rs.LrLogPath)
	LR_LOG_FILENAME				= ExpandEnvVars(rs.LrLogFilename)
	LR_MAX_ARCHIVES				= rs.LrMaxArchives
End Sub
