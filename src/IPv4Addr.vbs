'##############################################################################
'# FILE: IPv4Addr.vbs
'# DATE: 2018.03.14
'#   BY: Robert Mosier
'#       Manages IPv4 address validity and format.
'# DEPENDENCIES:
'#       None
'##############################################################################
Option Explicit


Class IPv4Addr

	'An array of four integers
	Private octets


	'##########################################################################
	'Create the array of integers for octets.
	Public Sub Class_Initialize()
		octets = Array(0, 0, 0, 0)
	End Sub 'Class_Initialize


	'##########################################################################
	'Set the IPv4 address.
	'Parameters
	'	sIPv4 As String:	A string that represents an IPv4 address. It should
	'		consist of only 4 octets of decimal coded byte values (between 0
	'		and 255) and periods between each octet. There should be no slashes,
	'		subnet mask, or CIDR bit mask. General form is "aaa.bbb.ccc.ddd"
	Public Sub SetAddress(sIPv4)
		Dim tmpOctets: tmpOctets = Split(sIPv4, ".")
		Dim i
		
		For i = 0 to 3
			octets(i) = CInt(tmpOctets(i))
		Next
	End Sub 'SetAddress


	'##########################################################################
	'Get a specific octet from the IPv4 address.
	'Parameters
	'	iIndex As Integer:	An unsigned integer ranging from 0 to 3 specifying
	'		which octet to return.
	'Returns
	'	GetOctet As Integer:	An unsigned integer one byte wide.
	Public Function GetOctet(iIndex)
		GetOctet = octets(iIndex)
	End Function 'GetOctet


	'##########################################################################
	'Get the value of the IPv4 address as a string.
	'Returns
	'	ToString As String:	A string consisting of four decimal coded unsigned
	'	bytes separated by periods.
	Public Function ToString()
		ToString = Join(octets, ".")
	End Function 'ToString


	'##########################################################################
	'Get the value of the IPv4 address as an unsigned integer
	'Returns
	'	ToInteger As Integer:	An unsigned integer representing the value of
	'	the IPv4 address which is four bytes (32 bits) wide.
	Public Function ToInteger()
		ToInteger = (GetOctet(0) * 256^3) + (GetOctet(1) * 256^2) + (GetOctet(2) * 256) + GetOctet(3)
	End Function 'ToInteger


	'##########################################################################
	'Check if a string represents a valid IPv4 address.
	'Parameters
	'	sIPv4 As String:	A string that represents an IPv4 address. It should
	'		consist of only 4 octets of decimal coded byte values (between 0
	'		and 255) and periods between each octet. There should be no slashes,
	'		subnet mask, or CIDR bit mask. General form is "aaa.bbb.ccc.ddd"
	'Returns
	'	ValidIPv4 As Boolean:	True if sIPv4 is valid or False otherwise.
	Public Function ValidIPv4(sIPv4)
		Dim ipAddr: ipAddr = Split(sIPv4, ".")
		'Assume the address is not valid until proven otherwise
		Dim result: result = False
		
		Dim octet
		
		'Check that the IP has 4 octets
		If UBound(ipAddr) = 3 Then
			result = True
			For Each octet in ipAddr 
				octet = Trim(octet)
				
				'Check that the octets are numeric and unsigned byte value
				If NOT isNumeric(octet) Then 
					result = False
					Exit For
				ElseIf octet < 0 OR octet > 255 Then
					result = False
					Exit For
				End If			
			Next
		End If
		
		ValidIPv4 = result
	End Function 'ValidIPv4

End Class 'IPv4Addr
