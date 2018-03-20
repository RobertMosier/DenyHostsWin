'##############################################################################
'# FILE: Test_LogRotate.vbs
'# DATE: 2018.03.14
'#   BY: Robert Mosier
'#       Tests the LogRotate class and demonstrates basic usage.
'# DEPENDENCIES:
'#       HelperLib.vbs
'#       LogRotate.vbs
'##############################################################################
Option Explicit


Sub Test_LogRotate()
	Dim lr: Set lr = New LogRotate

	lr.Rotate
End Sub
