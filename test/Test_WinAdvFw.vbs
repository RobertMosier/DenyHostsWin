'##############################################################################
'# FILE: Test_WinAdvFw.vbs
'# DATE: 2018.03.14
'#   BY: Robert Mosier
'#       Tests the WinAdvFw class and demonstrates basic usage.
'# DEPENDENCIES:
'#       WinAdvFw.vbs
'##############################################################################
Option Explicit


Sub Test_WinAdvFw()
	'Declare & instantiate WinAdvFw
	Dim fw: Set fw = New WinAdvFw

	'Block some IPv4 addresses (creates the firewall rule if not pre-existing)
	fw.BlockIPv4Address "57.0.0.1"
	fw.BlockIPv4Address "58.0.0.1"
	fw.BlockIPv4Address "59.0.0.1"
End Sub