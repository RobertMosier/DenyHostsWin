'##############################################################################
'# FILE: SettingsEditor.vbs
'# DATE: 2018.03.19
'#   BY: Robert Mosier
'#       A collection of subroutines and functions for the Settings Editor app.
'# DEPENDENCIES:
'#       RegSettings.vbs
'##############################################################################
Option Explicit


Dim oSettings: Set oSettings = New RegSettings


Function GetFSPath()
	Dim objShell
	Dim objFolder
	
	Set objShell = CreateObject("shell.application")
		Set objFolder = objShell.BrowseForFolder(0, "Select a path:", 0, "")
			If Not objFolder Is Nothing Then
				GetFSPath = objFolder.Self.Path
			End If
		Set objFolder = nothing
	Set objShell = nothing
End Function


Sub BrowseFS(sElement)
	Dim el: Set el = document.getElementById(sElement)
	Dim sPath: sPath = GetFSPath()
	
	If sPath <> "" Then
		el.value = sPath
	End If
End Sub


Sub LoadSettings
	document.getElementById("elCurrentVersion").innerHTML = oSettings.CurrentVersion
	
	document.getElementById("txtThreshold").value = oSettings.Threshold
	document.getElementById("txtFwRuleBaseName").value = oSettings.FwRuleBaseName
	document.getElementById("txtFwRuleGroupName").value = oSettings.FwRuleGroupName
	document.getElementById("areaFwRuleDescription").value = oSettings.FwRuleDescription
	document.getElementById("txtFwRuleLocalPorts").value = oSettings.FwRuleLocalPorts
	document.getElementById("txtFwLogFile").value = oSettings.FwLogFile
	document.getElementById("txtFwPort").value = oSettings.FwPort
	
	If oSettings.NotifyOnThreshold > 0 Then
		document.getElementById("chkNotifyOnThreshold").Checked = True
	Else
		document.getElementById("chkNotifyOnThreshold").Checked = False
	End If
	
	document.getElementById("txtNotifyThreshold").value = oSettings.NotifyThreshold
	document.getElementById("txtNotifyEmailFrom").value = oSettings.NotifyEmailFrom
	document.getElementById("txtNotifyEmailTo").value = oSettings.NotifyEmailTo
	document.getElementById("txtNotifyEmailSubject").value = oSettings.NotifyEmailSubject
	document.getElementById("txtNotifyEmailMsgTemplate").value = oSettings.NotifyEmailMsgTemplate
	document.getElementById("txtNotifySMTPAuthUser").value = oSettings.NotifySMTPAuthUser
	document.getElementById("txtNotifySMTPAuthPass").value = oSettings.NotifySMTPAuthPass
	document.getElementById("txtLrLogPath").value = oSettings.LrLogPath
	document.getElementById("txtLrLogFilename").value = oSettings.LrLogFilename
	document.getElementById("txtLrMaxArchives").value = oSettings.LrMaxArchives
End Sub


Sub SaveSettings
	oSettings.Threshold = document.getElementById("txtThreshold").value
	oSettings.FwRuleBaseName = document.getElementById("txtFwRuleBaseName").value
	oSettings.FwRuleGroupName = document.getElementById("txtFwRuleGroupName").value
	oSettings.FwRuleDescription = document.getElementById("areaFwRuleDescription").value
	oSettings.FwRuleLocalPorts = document.getElementById("txtFwRuleLocalPorts").value
	oSettings.FwLogFile = document.getElementById("txtFwLogFile").value
	oSettings.FwPort = document.getElementById("txtFwPort").value
	
	If document.getElementById("chkNotifyOnThreshold").Checked Then
		oSettings.NotifyOnThreshold = 1
	Else
		oSettings.NotifyOnThreshold = 0
	End If
	
	oSettings.NotifyThreshold = document.getElementById("txtNotifyThreshold").value
	oSettings.NotifyEmailFrom = document.getElementById("txtNotifyEmailFrom").value
	oSettings.NotifyEmailTo = document.getElementById("txtNotifyEmailTo").value
	oSettings.NotifyEmailSubject = document.getElementById("txtNotifyEmailSubject").value
	oSettings.NotifyEmailMsgTemplate = document.getElementById("txtNotifyEmailMsgTemplate").value
	oSettings.NotifySMTPAuthUser = document.getElementById("txtNotifySMTPAuthUser").value
	oSettings.NotifySMTPAuthPass = document.getElementById("txtNotifySMTPAuthPass").value
	oSettings.LrLogPath = document.getElementById("txtLrLogPath").value
	oSettings.LrLogFilename = document.getElementById("txtLrLogFilename").value
	oSettings.LrMaxArchives = document.getElementById("txtLrMaxArchives").value
End Sub


Sub SaveAndClose
	SaveSettings
	self.close()
End Sub


Sub ShowTooltip(sElement)
	document.getElementById(sElement).style.display = "block"
End Sub

Sub HideTooltip(sElement)
	document.getElementById(sElement).style.display = "none"
End Sub


Sub Window_onLoad
	LoadSettings
End Sub


Sub txtThreshold_onFocus
	ShowTooltip "ttThreshold"
End Sub

Sub txtThreshold_onBlur
	HideTooltip "ttThreshold"
End Sub


Sub txtFwRuleBaseName_onFocus
	ShowTooltip "ttFwRuleBaseName"
End Sub

Sub txtFwRuleBaseName_onBlur
	HideTooltip "ttFwRuleBaseName"
End Sub


Sub txtFwRuleGroupName_onFocus
	ShowTooltip "ttFwRuleGroupName"
End Sub

Sub txtFwRuleGroupName_onBlur
	HideTooltip "ttFwRuleGroupName"
End Sub


Sub areaFwRuleDescription_onFocus
	ShowTooltip "ttFwRuleDescription"
End Sub

Sub areaFwRuleDescription_onBlur
	HideTooltip "ttFwRuleDescription"
End Sub


Sub txtFwRuleLocalPorts_onFocus
	ShowTooltip "ttFwRuleLocalPorts"
End Sub

Sub txtFwRuleLocalPorts_onBlur
	HideTooltip "ttFwRuleLocalPorts"
End Sub


Sub txtFwLogFile_onFocus
	ShowTooltip "ttFwLogFile"
End Sub

Sub txtFwLogFile_onBlur
	HideTooltip "ttFwLogFile"
End Sub


Sub txtFwPort_onFocus
	ShowTooltip "ttFwPort"
End Sub

Sub txtFwPort_onBlur
	HideTooltip "ttFwPort"
End Sub


Sub chkNotifyOnThreshold_onFocus
	ShowTooltip "ttNotifyOnThreshold"
End Sub

Sub chkNotifyOnThreshold_onBlur
	HideTooltip "ttNotifyOnThreshold"
End Sub


Sub txtNotifyThreshold_onFocus
	ShowTooltip "ttNotifyThreshold"
End Sub

Sub txtNotifyThreshold_onBlur
	HideTooltip "ttNotifyThreshold"
End Sub


Sub txtNotifyEmailFrom_onFocus
	ShowTooltip "ttNotifyEmailFrom"
End Sub

Sub txtNotifyEmailFrom_onBlur
	HideTooltip "ttNotifyEmailFrom"
End Sub


Sub txtNotifyEmailTo_onFocus
	ShowTooltip "ttNotifyEmailTo"
End Sub

Sub txtNotifyEmailTo_onBlur
	HideTooltip "ttNotifyEmailTo"
End Sub


Sub txtNotifyEmailSubject_onFocus
	ShowTooltip "ttNotifyEmailSubject"
End Sub

Sub txtNotifyEmailSubject_onBlur
	HideTooltip "ttNotifyEmailSubject"
End Sub


Sub txtNotifyEmailMsgTemplate_onFocus
	ShowTooltip "ttNotifyEmailMsgTemplate"
End Sub

Sub txtNotifyEmailMsgTemplate_onBlur
	HideTooltip "ttNotifyEmailMsgTemplate"
End Sub


Sub txtNotifySMTPAuthUser_onFocus
	ShowTooltip "ttNotifySMTPAuthUser"
End Sub

Sub txtNotifySMTPAuthUser_onBlur
	HideTooltip "ttNotifySMTPAuthUser"
End Sub


Sub txtNotifySMTPAuthPass_onFocus
	ShowTooltip "ttNotifySMTPAuthPass"
End Sub

Sub txtNotifySMTPAuthPass_onBlur
	HideTooltip "ttNotifySMTPAuthPass"
End Sub


Sub txtLrLogPath_onFocus
	ShowTooltip "ttLrLogPath"
End Sub

Sub txtLrLogPath_onBlur
	HideTooltip "ttLrLogPath"
End Sub


Sub txtLrLogFilename_onFocus
	ShowTooltip "ttLrLogFilename"
End Sub

Sub txtLrLogFilename_onBlur
	HideTooltip "ttLrLogFilename"
End Sub


Sub txtLrMaxArchives_onFocus
	ShowTooltip "ttLrMaxArchives"
End Sub

Sub txtLrMaxArchives_onBlur
	HideTooltip "ttLrMaxArchives"
End Sub


Sub btnFwLogFile_onClick
	BrowseFS("txtFwLogFile")
End Sub


Sub btnFwLogFile_onClick
	BrowseFS("txtNotifyEmailMsgTemplate")
End Sub


Sub btnReset_onClick
	LoadSettings
End Sub


Sub btnOK_onClick
	SaveAndClose
End Sub


Sub btnCancel_onClick
	self.close()
End Sub


Sub btnApply_onClick
	SaveSettings
End Sub
