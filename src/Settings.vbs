'##############################################################################
'# FILE: Settings.vbs
'# DATE: 2018.03.19
'#   BY: Robert Mosier
'#       Groups of global variables used as settings for the various classes.
'# DEPENDENCIES:
'#       None
'##############################################################################
Option Explicit



'##############################################################################
'Class WinAdvFw


'##############################################################################
'This is the name that will show up as a firewall rule in the Windows Advanced
'Firewall. In fact there will be two, one with " - TCP" the other with " - UDP"
'appended to it.
'
'Default is "AUTO-BLACKLISTED IP ADDRESSES"
'
Dim FW_RULE_BASE_NAME


'##############################################################################
'The group name the firewall rule be a member of. This is a convenience thing
'for when manually managing the rules. Otherwise this software makes no use of
'it.
'
'Default is "AUTO-BLACKLIST"
'
Dim FW_RULE_GROUP_NAME


'##############################################################################
'Similar to the group name this is a convenience for when manually managing the
'firewall since this software makes no use of it. You can even set it to an
'empty string.
'
'Default is "Remote IP addresses blocked automatically"
'
Dim FW_RULE_DESCRIPTION


'##############################################################################
'This is the port or ports to block remote addresses from accessing. Acceptable
'values are;
'
'	1) A single port to block offenders from a single service such as RDP only.
'	FW_RULE_LOCAL_PORTS = "3389"
'
'	2) Multiple non-contiguous ports to block offenders from multiple specific
'	services such as FTP, SSH, & HTTP.
'	FW_RULE_LOCAL_PORTS = "21,22,80"
'
'	3) A range of ports to block offenders from multiple contiguous services or
'	even all services. (DEFAULT)
'
'Default is "0-65535"
'
Dim FW_RULE_LOCAL_PORTS



'##############################################################################
'Class WinAdvFw_LogParse


'##############################################################################
'This is the log file to read from. Logging must be enabled. Use this Microsoft
'web page to help enable logging of the Windows Advanced Firewall.
'https://docs.microsoft.com/en-us/windows/security/identity-protection/windows-firewall/configure-the-windows-firewall-log
'We use the 0.pfirewall.log file because we are rotating the log before we
'parse it.
'
'Default is "%windir%\system32\logfiles\firewall\0.pfirewall.log"
'
Dim FW_LOG_FILE


'##############################################################################
'This is the port to find in the log file
'
'Default is "3389" 'RDP port
'
Dim FW_PORT



'##############################################################################
'Application DenyHostsWin


'##############################################################################
'This is an Integer representing how many times a remote host has failed login
'before it is assumed to be hostile and gets blocked in the firewall.
'
'Default is 20
'
Dim THRESHOLD



'##############################################################################
'Class GMailNotify


'##############################################################################
'Should we notify admin by email if the number of blocked IPv4 addresses
'exceeds the quantity defined by NOTIFY_BLOCKED_THRESHOLD? (Use Boolean True or
'False only)
'
'Default is True
'
Dim NOTIFY_ON_THRESHOLD


'##############################################################################
'After how many blocked IPv4 addresses should we notify the admin by email?
'(Use an Unsigned Integer value equal to or greater than zero)
'
'Default is 5
'
Dim NOTIFY_THRESHOLD


'##############################################################################
'The email address to send from which will be a gmail.com address. This can be
'formatted as name and email like, "GMail User" <example@gmail.com>
'
'Default is """DenyHostsWin"" <relay.example@gmail.com>"
'
Dim NOTIFY_EMAIL_FROM


'##############################################################################
'The email address to send to which does not have to be a gmail.com address.
'This can be formatted as name and email like, "Outlook User" <example@outlook.com>
'
'Default is """Administrator"" <admin@example.com>"
'
Dim NOTIFY_EMAIL_TO


'##############################################################################
'The subject line to place within the email sent. This is a great opportunity
'to customize this instance to tell the admin what server this came from.
'
'Default is "ALERT from DenyHostsWin server.domain.tld"
'
Dim NOTIFY_EMAIL_SUBJECT


'##############################################################################
'The path to the message template file.
'
'Format of a message template file is a simple txt file with a few variable
'placeholders which will get preprocessed into values from the script before
'the email is sent.
'Variable Placeholders
'	${BLOCKED_COUNT}	Represents the number of IPv4 addresses blocked.
'	${THRESHOLD}		The value set above in NOTIFY_THRESHOLD.
'	${BLOCKED_ADDRS}	A list of IPv4 addresses blocked one per line.
'
'Default is "%PROGRAMFILES(X86)%\DenyHostsWin\msg-template.txt"
'
Dim NOTIFY_EMAIL_MSG_TEMPLATE


'##############################################################################
'The user name of the gmail account to send from. This is also the email
'address but must be formatted plain and simple.
'
'Default is "relay.example@gmail.com"
'
Dim NOTIFY_SMTP_AUTH_USER


'##############################################################################
'The password for the gmail account to send from. This is case-sensitive.
'
'Default is "Your_smtp_password"
'
Dim NOTIFY_SMTP_AUTH_PASS



'##############################################################################
'Class LogRotate


'##############################################################################
'The path of the directory containing the log file. Do not leave a trailing
'slash.
'
'Default is "%systemroot%\system32\LogFiles\Firewall"
'
Dim LR_LOG_PATH


'##############################################################################
'The file name of the log to be rotated. Do not include the path.
'
'Default is "pfirewall.log"
'
Dim LR_LOG_FILENAME


'##############################################################################
'How many archived log files to we want to keep? Use an unsigned integer here.
'When an archive gets rotated to have a file name prefix equal to this value,
'it will get deleted.
'
'Default is 5
'
Dim LR_MAX_ARCHIVES

