'##############################################################################
'# FILE: GMailNotify.vbs
'# DATE: 2018.03.14
'#   BY: Robert Mosier
'#       Notify an admin via email if specific threshold conditions are met.
'# DEPENDENCIES:
'#       IPv4Addr.vbs
'#       The following settings must be defined prior to class creation;
'#           NOTIFY_ON_THRESHOLD
'#           NOTIFY_THRESHOLD
'#           NOTIFY_EMAIL_FROM
'#           NOTIFY_EMAIL_TO
'#           NOTIFY_EMAIL_SUBJECT
'#           NOTIFY_EMAIL_MSG_TEMPLATE
'#           NOTIFY_SMTP_AUTH_USER
'#           NOTIFY_SMTP_AUTH_PASS
'##############################################################################
Option Explicit


Class GMailNotify

	'EMail message object
	Private emailMsg
	'The email message template
	Private msgTemplate
	'The number of IPv4 addresses blocked
	Private addrsBlocked
	
	
	'##########################################################################
	'Private class constants
	Private Property Get SCHEMA
		SCHEMA = "http://schemas.microsoft.com/cdo/configuration/"
	End Property
	
	
	'##########################################################################
	'Initialize the email message object
	Public Sub Class_Initialize()
		'Send the email using a remote SMTP server
		Const CDO_SEND_USING_PORT	= 2
		'Basic authentication
		Const CDO_BASIC_AUTH		= 1
		
		'Configuration object from the email message object
		Dim conf
		
		'Create the email message object
		Set emailMsg = CreateObject("CDO.Message")
		'Set the configuration object to conf for easy access
		Set conf = emailMsg.Configuration
		
		'Set the remote SMTP server info
		conf.Fields(SCHEMA & "sendusing")		= CDO_SEND_USING_PORT
		conf.Fields(SCHEMA & "smtpserver")		= "smtp.gmail.com"
		conf.Fields(SCHEMA & "smtpserverport")	= 465
		
		'Use basic authentication with SSL
		conf.Fields(SCHEMA & "smtpauthenticate") = CDO_BASIC_AUTH
		conf.Fields(SCHEMA & "smtpusessl")		= True
		
		'Set the credentials
		conf.Fields(SCHEMA & "sendusername")	= NOTIFY_SMTP_AUTH_USER
		conf.Fields(SCHEMA & "sendpassword")	= NOTIFY_SMTP_AUTH_PASS
		
		conf.Fields.Update()
		
		'Message header
		emailMsg.From     = NOTIFY_EMAIL_FROM
		emailMsg.To       = NOTIFY_EMAIL_TO
		emailMsg.Subject  = NOTIFY_EMAIL_SUBJECT
		
		loadMessageTemplate NOTIFY_EMAIL_MSG_TEMPLATE
	End Sub 'Class_Initialize
	
	
	'##########################################################################
	'Load the message template file into the msgTemplate variable.
	'Parameters
	'	sTemplateFile As String:	The path and file name of the message
	'		template file. This should come from NOTIFY_EMAIL_MSG_TEMPLATE.
	Private Sub loadMessageTemplate(sTemplateFile)
		Const ForReading = 1
		Dim fso, f
		
		Set fso = CreateObject("Scripting.FileSystemObject")
		'Open the file for reading
		Set f = fso.OpenTextFile(sTemplateFile, ForReading)
		'read the entire file in
		msgTemplate = f.ReadAll
		
		f.Close
		Set f = Nothing
		Set fso = Nothing
	End Sub 'loadMessageTemplate
	
	
	Public Sub preprocessMessage(aAddrs)
		Dim str, ip, sAddrs
		Dim first: first = True
		
		'Get the count of blocked addresses
		addrsBlocked = UBound(aAddrs) + 1
		
		'Preprocess out the variables into values
		str = Replace(msgTemplate, "${BLOCKED_COUNT}", addrsBlocked)
		str = Replace(str, "${THRESHOLD}", NOTIFY_THRESHOLD)
		
		'Create the list of addresses from an array of IPv4Addr
		For Each ip in aAddrs
			If first Then
				first = False
				sAddrs = ip.ToString
			Else
				sAddrs = sAddrs & vbCrLf & ip.ToString
			End If
		Next
		
		'Preprocess out the BLOCKED_ADDRS variable into the address
		'list we just made
		str = Replace(str, "${BLOCKED_ADDRS}", sAddrs)
		
		'Finally assign the preprocessed message body to the email object
		emailMsg.TextBody = str
	End Sub
	
	
	Public Sub Send()
		If NOTIFY_ON_THRESHOLD And addrsBlocked > NOTIFY_THRESHOLD Then
			emailMsg.Send()
		End If
	End Sub
	
End Class 'GMailNotify
