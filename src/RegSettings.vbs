'##############################################################################
'# FILE: RegSettings.vbs
'# DATE: 2018.03.19
'#   BY: Robert Mosier
'#       Provide access to registry settings.
'# DEPENDENCIES:
'#       None
'##############################################################################
Option Explicit


Class RegSettings

	Private Property Get VERSION()
		VERSION = "18.03"
	End Property

	Private Property Get KEY_BASE()
		KEY_BASE = "HKLM\SOFTWARE\DenyHostsWin"
	End Property

	Private Property Get KEY_CURRENT_VERSION()
		KEY_CURRENT_VERSION = KEY_BASE & "\CurrentVersion"
	End Property

	Private Property Get KEY_THRESHOLD()
		KEY_THRESHOLD = ""
		
		If getRegValue(KEY_CURRENT_VERSION) = VERSION Then
			KEY_THRESHOLD = KEY_BASE & "\" & VERSION & "\Threshold"
		End If
	End Property

	Private Property Get KEY_FWRULEBASENAME()
		KEY_FWRULEBASENAME = ""
		
		If getRegValue(KEY_CURRENT_VERSION) = VERSION Then
			KEY_FWRULEBASENAME = KEY_BASE & "\" & VERSION & "\FirewallRules\FwRuleBaseName"
		End If
	End Property

	Private Property Get KEY_FWRULEGROUPNAME()
		KEY_FWRULEGROUPNAME = ""
		
		If getRegValue(KEY_CURRENT_VERSION) = VERSION Then
			KEY_FWRULEGROUPNAME = KEY_BASE & "\" & VERSION & "\FirewallRules\FwRuleGroupName"
		End If
	End Property

	Private Property Get KEY_FWRULEDESCRIPTION()
		KEY_FWRULEDESCRIPTION = ""
		
		If getRegValue(KEY_CURRENT_VERSION) = VERSION Then
			KEY_FWRULEDESCRIPTION = KEY_BASE & "\" & VERSION & "\FirewallRules\FwRuleDescription"
		End If
	End Property

	Private Property Get KEY_FWRULELOCALPORTS()
		KEY_FWRULELOCALPORTS = ""
		
		If getRegValue(KEY_CURRENT_VERSION) = VERSION Then
			KEY_FWRULELOCALPORTS = KEY_BASE & "\" & VERSION & "\FirewallRules\FwRuleLocalPorts"
		End If
	End Property

	Private Property Get KEY_FWLOGFILE()
		KEY_FWLOGFILE = ""
		
		If getRegValue(KEY_CURRENT_VERSION) = VERSION Then
			KEY_FWLOGFILE = KEY_BASE & "\" & VERSION & "\FirewallLog\FwLogFile"
		End If
	End Property

	Private Property Get KEY_FWPORT()
		KEY_FWPORT = ""
		
		If getRegValue(KEY_CURRENT_VERSION) = VERSION Then
			KEY_FWPORT = KEY_BASE & "\" & VERSION & "\FirewallLog\FwPort"
		End If
	End Property

	Private Property Get KEY_NOTIFYONTHRESHOLD()
		KEY_NOTIFYONTHRESHOLD = ""
		
		If getRegValue(KEY_CURRENT_VERSION) = VERSION Then
			KEY_NOTIFYONTHRESHOLD = KEY_BASE & "\" & VERSION & "\EMailNotify\NotifyOnThreshold"
		End If
	End Property

	Private Property Get KEY_NOTIFYTHRESHOLD()
		KEY_NOTIFYTHRESHOLD = ""
		
		If getRegValue(KEY_CURRENT_VERSION) = VERSION Then
			KEY_NOTIFYTHRESHOLD = KEY_BASE & "\" & VERSION & "\EMailNotify\NotifyThreshold"
		End If
	End Property

	Private Property Get KEY_NOTIFYEMAILFROM()
		KEY_NOTIFYEMAILFROM = ""
		
		If getRegValue(KEY_CURRENT_VERSION) = VERSION Then
			KEY_NOTIFYEMAILFROM = KEY_BASE & "\" & VERSION & "\EMailNotify\NotifyEmailFrom"
		End If
	End Property

	Private Property Get KEY_NOTIFYEMAILTO()
		KEY_NOTIFYEMAILTO = ""
		
		If getRegValue(KEY_CURRENT_VERSION) = VERSION Then
			KEY_NOTIFYEMAILTO = KEY_BASE & "\" & VERSION & "\EMailNotify\NotifyEmailTo"
		End If
	End Property

	Private Property Get KEY_NOTIFYEMAILSUBJECT()
		KEY_NOTIFYEMAILSUBJECT = ""
		
		If getRegValue(KEY_CURRENT_VERSION) = VERSION Then
			KEY_NOTIFYEMAILSUBJECT = KEY_BASE & "\" & VERSION & "\EMailNotify\NotifyEmailSubject"
		End If
	End Property

	Private Property Get KEY_NOTIFYEMAILMSGTEMPLATE()
		KEY_NOTIFYEMAILMSGTEMPLATE = ""
		
		If getRegValue(KEY_CURRENT_VERSION) = VERSION Then
			KEY_NOTIFYEMAILMSGTEMPLATE = KEY_BASE & "\" & VERSION & "\EMailNotify\NotifyEmailMsgTemplate"
		End If
	End Property

	Private Property Get KEY_NOTIFYSMTPAUTHUSER()
		KEY_NOTIFYSMTPAUTHUSER = ""
		
		If getRegValue(KEY_CURRENT_VERSION) = VERSION Then
			KEY_NOTIFYSMTPAUTHUSER = KEY_BASE & "\" & VERSION & "\EMailNotify\NotifySMTPAuthUser"
		End If
	End Property

	Private Property Get KEY_NOTIFYSMTPAUTHPASS()
		KEY_NOTIFYSMTPAUTHPASS = ""
		
		If getRegValue(KEY_CURRENT_VERSION) = VERSION Then
			KEY_NOTIFYSMTPAUTHPASS = KEY_BASE & "\" & VERSION & "\EMailNotify\NotifySMTPAuthPass"
		End If
	End Property

	Private Property Get KEY_LRLOGPATH()
		KEY_LRLOGPATH = ""
		
		If getRegValue(KEY_CURRENT_VERSION) = VERSION Then
			KEY_LRLOGPATH = KEY_BASE & "\" & VERSION & "\LogRotate\LrLogPath"
		End If
	End Property

	Private Property Get KEY_LRLOGFILENAME()
		KEY_LRLOGFILENAME = ""
		
		If getRegValue(KEY_CURRENT_VERSION) = VERSION Then
			KEY_LRLOGFILENAME = KEY_BASE & "\" & VERSION & "\LogRotate\LrLogFilename"
		End If
	End Property

	Private Property Get KEY_LRMAXARCHIVES()
		KEY_LRMAXARCHIVES = ""
		
		If getRegValue(KEY_CURRENT_VERSION) = VERSION Then
			KEY_LRMAXARCHIVES = KEY_BASE & "\" & VERSION & "\LogRotate\LrMaxArchives"
		End If
	End Property
	
	
	Private Function getRegValue(sKey)
		Dim oShell: Set oShell = CreateObject("WScript.Shell")
		
		If sKey = ""  Then
			getRegValue = ""
		Else
			getRegValue = oShell.RegRead(sKey)
		End If
	End Function
	
	Private Sub setRegValue(sKey, vVal, sType)
		Dim oShell: Set oShell = CreateObject("WScript.Shell")
		
		oShell.RegWrite sKey, vVal, sType
	End Sub
	
	
	Public Property Get CurrentVersion()
		CurrentVersion = getRegValue(KEY_CURRENT_VERSION)
	End Property
	
	
	Public Property Get Threshold()
		Threshold = getRegValue(KEY_THRESHOLD)
	End Property
	
	Public Property Let Threshold(iValue)
		setRegValue KEY_THRESHOLD, iValue, "REG_DWORD"
	End Property
	
	
	Public Property Get FwRuleBaseName()
		FwRuleBaseName = getRegValue(KEY_FWRULEBASENAME)
	End Property
	
	Public Property Let FwRuleBaseName(sValue)
		setRegValue KEY_FWRULEBASENAME, sValue, "REG_SZ"
	End Property
	
	
	Public Property Get FwRuleGroupName()
		FwRuleGroupName = getRegValue(KEY_FWRULEGROUPNAME)
	End Property
	
	Public Property Let FwRuleGroupName(sValue)
		setRegValue KEY_FWRULEGROUPNAME, sValue, "REG_SZ"
	End Property
	
	
	Public Property Get FwRuleDescription()
		FwRuleDescription = getRegValue(KEY_FWRULEDESCRIPTION)
	End Property
	
	Public Property Let FwRuleDescription(sValue)
		setRegValue KEY_FWRULEDESCRIPTION, sValue, "REG_SZ"
	End Property
	
	
	Public Property Get FwRuleLocalPorts()
		FwRuleLocalPorts = getRegValue(KEY_FWRULELOCALPORTS)
	End Property
	
	Public Property Let FwRuleLocalPorts(sValue)
		setRegValue KEY_FWRULELOCALPORTS, sValue, "REG_SZ"
	End Property
	
	
	Public Property Get FwLogFile()
		FwLogFile = getRegValue(KEY_FWLOGFILE)
	End Property
	
	Public Property Let FwLogFile(sValue)
		setRegValue KEY_FWLOGFILE, sValue, "REG_SZ"
	End Property
	
	
	Public Property Get FwPort()
		FwPort = getRegValue(KEY_FWPORT)
	End Property
	
	Public Property Let FwPort(sValue)
		setRegValue KEY_FWPORT, sValue, "REG_SZ"
	End Property
	
	
	Public Property Get NotifyOnThreshold()
		NotifyOnThreshold = getRegValue(KEY_NOTIFYONTHRESHOLD)
	End Property
	
	Public Property Let NotifyOnThreshold(iValue)
		setRegValue KEY_NOTIFYONTHRESHOLD, iValue, "REG_DWORD"
	End Property
	
	
	Public Property Get NotifyThreshold()
		NotifyThreshold = getRegValue(KEY_NOTIFYTHRESHOLD)
	End Property
	
	Public Property Let NotifyThreshold(iValue)
		setRegValue KEY_NOTIFYTHRESHOLD, iValue, "REG_DWORD"
	End Property
	
	
	Public Property Get NotifyEmailFrom()
		NotifyEmailFrom = getRegValue(KEY_NOTIFYEMAILFROM)
	End Property
	
	Public Property Let NotifyEmailFrom(sValue)
		setRegValue KEY_NOTIFYEMAILFROM, sValue, "REG_SZ"
	End Property
	
	
	Public Property Get NotifyEmailTo()
		NotifyEmailTo = getRegValue(KEY_NOTIFYEMAILTO)
	End Property
	
	Public Property Let NotifyEmailTo(sValue)
		setRegValue KEY_NOTIFYEMAILTO, sValue, "REG_SZ"
	End Property
	
	
	Public Property Get NotifyEmailSubject()
		NotifyEmailSubject = getRegValue(KEY_NOTIFYEMAILSUBJECT)
	End Property
	
	Public Property Let NotifyEmailSubject(sValue)
		setRegValue KEY_NOTIFYEMAILSUBJECT, sValue, "REG_SZ"
	End Property
	
	
	Public Property Get NotifyEmailMsgTemplate()
		NotifyEmailMsgTemplate = getRegValue(KEY_NOTIFYEMAILMSGTEMPLATE)
	End Property
	
	Public Property Let NotifyEmailMsgTemplate(sValue)
		setRegValue KEY_NOTIFYEMAILMSGTEMPLATE, sValue, "REG_SZ"
	End Property
	
	
	Public Property Get NotifySMTPAuthUser()
		NotifySMTPAuthUser = getRegValue(KEY_NOTIFYSMTPAUTHUSER)
	End Property
	
	Public Property Let NotifySMTPAuthUser(sValue)
		setRegValue KEY_NOTIFYSMTPAUTHUSER, sValue, "REG_SZ"
	End Property
	
	
	Public Property Get NotifySMTPAuthPass()
		NotifySMTPAuthPass = getRegValue(KEY_NOTIFYSMTPAUTHPASS)
	End Property
	
	Public Property Let NotifySMTPAuthPass(sValue)
		setRegValue KEY_NOTIFYSMTPAUTHPASS, sValue, "REG_SZ"
	End Property
	
	
	Public Property Get LrLogPath()
		LrLogPath = getRegValue(KEY_LRLOGPATH)
	End Property
	
	Public Property Let LrLogPath(sValue)
		setRegValue KEY_LRLOGPATH, sValue, "REG_SZ"
	End Property
	
	
	Public Property Get LrLogFilename()
		LrLogFilename = getRegValue(KEY_LRLOGFILENAME)
	End Property
	
	Public Property Let LrLogFilename(sValue)
		setRegValue KEY_LRLOGFILENAME, sValue, "REG_SZ"
	End Property
	
	
	Public Property Get LrMaxArchives()
		LrMaxArchives = getRegValue(KEY_LRMAXARCHIVES)
	End Property
	
	Public Property Let LrMaxArchives(iValue)
		setRegValue KEY_LRMAXARCHIVES, iValue, "REG_DWORD"
	End Property

End Class 'RegSettings
