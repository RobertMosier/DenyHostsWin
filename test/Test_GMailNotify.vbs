'##############################################################################
'# FILE: Test_GMailNotify.vbs
'# DATE: 2018.03.14
'#   BY: Robert Mosier
'#       Tests the GMailNotify class and demonstrates basic usage.
'# DEPENDENCIES:
'#       HelperLib.vbs
'#       IPv4Addr.vbs
'#       GMailNotify.vbs
'##############################################################################
Option Explicit


Sub Test_GMailNotify()
	Dim gn: Set gn = New GMailNotify
	Dim aIP, oIP, sIP
	
	For Each sIP In Array("8.8.8.8", "8.8.4.4", "4.4.4.4", "4.4.2.2")
		Set oIP = New IPv4Addr
		oIP.SetAddress sIP
		ArrayPush aIP, oIP
	Next
	
	gn.preprocessMessage aIP
	gn.Send
End Sub
