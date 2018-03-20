'##############################################################################
'# FILE: HelperLib.vbs
'# DATE: 2018.03.14
'#   BY: Robert Mosier
'#       Helper functions that do not belong to any single class.
'# DEPENDENCIES:
'#       None
'##############################################################################
Option Explicit


'##########################################################################
'Expand environment variables. Useful for user configurable settings.
'Parameters
'	str As String:	A string that may contain environment variables to expand.
'Returns
'	ExpandEnvVars As String: A string that has had all valid environment
'		variables expanded. Otherwise the string will match str.
Function ExpandEnvVars(str)
	Dim wshShell: Set wshShell = CreateObject("WScript.Shell")
	ExpandEnvVars = wshShell.ExpandEnvironmentStrings(str)
End Function 'ExpandEnvVars


'##########################################################################
'Add an element to an array as the last element
'Parameters
'	arr As Array:	An array to add an element to.
'	val As Variant:	A value to add to arr as a new element.
Sub ArrayPush(ByRef arr, val)
	Dim size: size = 0
	
	'Verify arr is actually an array
	If IsArray(arr) Then
		'Resize arr to 1 more element
		size = UBound(arr) + 1
		ReDim Preserve arr(size)
	Else
		'Make arr an array of size 1
		ReDim arr(size)
	End If
	
	'Assign val to new element on arr
	If IsObject(val) Then
		Set arr(size) = val
	Else
		arr(size) = val
	End If
End Sub 'ArrayPush
