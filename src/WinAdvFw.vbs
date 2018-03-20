'##############################################################################
'# FILE: WinAdvFw.vbs
'# DATE: 2018.03.14
'#   BY: Robert Mosier
'#       Creates and updates the firewall rules responsible for blocking remote
'#       offenders.
'# DEPENDENCIES:
'#       HelperLib.vbs
'#       The following settings must be defined prior to class creation;
'#           FW_RULE_BASE_NAME
'#           FW_RULE_GROUP_NAME
'#           FW_RULE_DESCRIPTION
'#           FW_RULE_LOCAL_PORTS
'##############################################################################
Option Explicit


Class WinAdvFw

	'The variables used for the settings
	Private fwRuleBaseName, fwRuleGroupName, fwRuleDescription, fwRuleLocalPorts
	
	
	'##########################################################################
	'Initialize the class with user specific settings.
	'See Settings.vbs for more detailed info.
	Public Sub Class_Initialize()
		fwRuleBaseName		= ExpandEnvVars(FW_RULE_BASE_NAME)
		fwRuleGroupName		= ExpandEnvVars(FW_RULE_GROUP_NAME)
		fwRuleDescription	= ExpandEnvVars(FW_RULE_DESCRIPTION)
		fwRuleLocalPorts	= ExpandEnvVars(FW_RULE_LOCAL_PORTS)
	End Sub
	
	
	'##########################################################################
	'These properties are read-only in order to be like private constants and
	'serve as self documentation to their purpose. They are simply nice names
	'in the place of values that make no sense by themselves.
	Private Property Get NET_FW_PROFILE2_DOMAIN
		NET_FW_PROFILE2_DOMAIN = 1
	End Property
	
	Private Property Get NET_FW_PROFILE2_PRIVATE
		NET_FW_PROFILE2_PRIVATE = 2
	End Property
	
	Private Property Get NET_FW_PROFILE2_PUBLIC
		NET_FW_PROFILE2_PUBLIC = 4
	End Property
	
	Private Property Get NET_FW_IP_PROTOCOL_TCP
		NET_FW_IP_PROTOCOL_TCP = 6
	End Property
	
	Private Property Get NET_FW_IP_PROTOCOL_UDP
		NET_FW_IP_PROTOCOL_UDP = 17
	End Property
	
	Private Property Get NET_FW_RULE_DIR_IN
		NET_FW_RULE_DIR_IN = 1
	End Property
	
	Private Property Get NET_FW_ACTION_BLOCK
		NET_FW_ACTION_BLOCK = 0
	End Property
	
	
	'##########################################################################
	'Checks each firewall rule in policy for a matching name.
	'Parameters
	'	policy As Object:	A HNetCfg.FwPolicy2 object to search rules for.
	'	ruleName As String:	A string containing the name of a firewall rule.
	'Returns
	'	FindFwRuleByName As Object:	A HNetCfg.FwRule object if ruleName is found
	'		in policy. If not found Nothing is returned.
	Private Function FindFwRuleByName(policy, ruleName)
		Dim rules: Set rules = policy.Rules
		Dim rule
		
		For Each rule in rules
			If rule.Name = ruleName Then
				Set FindFwRuleByName = rule
				Exit Function
			End If
		Next
		
		Set FindFwRuleByName = Nothing
	End Function
	
	
	'##########################################################################
	'Uses the base name and protocol to determine the full firewall rule name.
	'Parameters
	'	protocol As Integer:	Simply use the two class constants provided,
	'		NET_FW_IP_PROTOCOL_TCP or NET_FW_IP_PROTOCOL_UDP.
	'Returns
	'	GetFwRuleNameByProto As String:	A string is returned which is the full
	'		name of a firewall rule.
	Private Function GetFwRuleNameByProto(protocol)
		Dim name: name = fwRuleBaseName & " - "

		If protocol = NET_FW_IP_PROTOCOL_TCP Then
			name = name & "TCP"
		Else
			name = name & "UDP"
		End If
		
		GetFwRuleNameByProto = name
	End Function
	
	
	'##########################################################################
	'Creates a new firewall rule and inserts it into the policy.
	'Parameters
	'	protocol As Integer:	Simply use the two class constants provided,
	'		NET_FW_IP_PROTOCOL_TCP or NET_FW_IP_PROTOCOL_UDP.
	'Returns
	'	InsertFirewallRule As Object:	A HNetCfg.FwRule that has already been
	'		added to the policy but is returned for convenience since it is
	'		likely to be used after creation.
	Private Function InsertFirewallRule(protocol)
		Dim policy:	Set policy	= CreateObject("HNetCfg.FwPolicy2")
		Dim rule:	Set rule	= CreateObject("HNetCfg.FwRule")
		
		rule.Name			= GetFwRuleNameByProto(protocol)
		rule.Protocol		= protocol
		
		rule.Enabled		= TRUE
		rule.LocalPorts		= fwRuleLocalPorts
		rule.Description	= fwRuleDescription
		rule.Grouping		= fwRuleGroupName
		rule.Direction		= NET_FW_RULE_DIR_IN
		rule.Profiles		= NET_FW_PROFILE2_DOMAIN + NET_FW_PROFILE2_PRIVATE + NET_FW_PROFILE2_PUBLIC
		rule.Action			= NET_FW_ACTION_BLOCK

		policy.Rules.Add rule
		Set InsertFirewallRule = rule
	End Function
	
	
	'##########################################################################
	'Inserts a new remote IPv4 address into an existing firewall rule. If the
	'firewall rule does not already exist, it will be created.
	'Parameters
	'	protocol As Integer:	Simply use the two class constants provided,
	'		NET_FW_IP_PROTOCOL_TCP or NET_FW_IP_PROTOCOL_UDP.
	'	ipAddress As String:	A string containing an IPv4 address. The
	'		address must be a host address only without subnet mask or CIDR bit
	'		mask. Should simply contain 4 octets with decimal coded bytes and
	'		periods between octets.
	Private Sub InsertRemoteAddress(protocol, ipAddress)
		Dim name:		name	= GetFwRuleNameByProto(protocol)
		Dim ipv4:		ipv4	= ipAddress & "/255.255.255.255"
		Dim policy:	Set policy	= CreateObject("HNetCfg.FwPolicy2")
		Dim rule:	Set rule	= FindFwRuleByName(policy, name)
		
		If rule Is Nothing Then
			Set rule = InsertFirewallRule(protocol)
		End If
		
		'If this is the first address to block just write it, otherwise append.
		If Len(rule.RemoteAddresses) > 1 Then 
			If InStr(rule.RemoteAddresses, ipv4) = 0 Then
				rule.RemoteAddresses = rule.RemoteAddresses & "," & ipv4
			End If
		Else
			rule.RemoteAddresses = ipv4
		End If
	End Sub
	
	
	'##########################################################################
	'Inserts a new remote IPv4 address into existing firewall rules blocking
	'both TCP and UDP protocols. If the firewall rules do not already exist,
	'they will be created.
	'Parameters
	'	ipAddress As String:	A string containing an IPv4 address. The
	'		address must be a host address only without subnet mask or CIDR bit
	'		mask. Should simply contain 4 octets with decimal coded bytes and
	'		periods between octets.
	Public Sub BlockIPv4Address(ipAddress)
		InsertRemoteAddress NET_FW_IP_PROTOCOL_TCP, ipAddress
		InsertRemoteAddress NET_FW_IP_PROTOCOL_UDP, ipAddress
	End Sub

End Class 'WinAdvFw
