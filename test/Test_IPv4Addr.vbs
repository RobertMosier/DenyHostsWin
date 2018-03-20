'##############################################################################
'# FILE: Test_IPv4Addr.vbs
'# DATE: 2018.03.14
'#   BY: Robert Mosier
'#       Tests the IPv4Addr class and demonstrates basic usage.
'# DEPENDENCIES:
'#       IPv4Addr.vbs
'##############################################################################
Option Explicit


Dim ipAddrs
Dim addrs: addrs = Array( _
	"172.9.243.43", _
	"192.168.98.34", _
	"208.34.34.12", _
	"10.65.1.56", _
	"59.0.12.24", _
	"172.9.243.44", _
	"192.168.98.50", _
	"98.189.56.91" _
)


Sub Init()
	Dim count: count = UBound(addrs)
	'Dim aIP(count)
	ReDim ipAddrs(count)
	Dim ip, i
	
	For i = 0 to count
		Set ip = New IPv4Addr
		ip.SetAddress addrs(i)
		'Set aIP(i) = ip
		Set ipAddrs(i) = ip
	Next
End Sub


Sub SortAddrs(aIPv4)
	Dim i, j, TempValue

	For i = LBound(aIPv4) to UBound(aIPv4)
		For j = LBound(aIPv4) to UBound(aIPv4) - 1
		
			If aIPv4(j).ToInteger > aIPv4(j + 1).ToInteger Then
				Set TempValue = aIPv4(j + 1)
				Set aIPv4(j + 1) = aIPv4(j)
				Set aIPv4(j) = TempValue
			End If
			
		Next
	Next
End Sub


Sub DisplayAddrs(aIPv4)
	Dim ip
	
	For Each ip in aIPv4
		WScript.Echo ip.ToString
	Next
End Sub


Sub Test_IPv4Addr()
	Init
	SortAddrs ipAddrs
	DisplayAddrs ipAddrs
End Sub
