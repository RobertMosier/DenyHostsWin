                           Welcome to DenyHostsWin
             Simple common sense security for Windows has arrived!

-------------------------------------------------------------------------------

INTRODUCTION

    DenyHostsWin is a scripted solution to basic Windows network security
through strategic cyclical use of the Windows Advanced Firewall and its logs.
DenyHostsWin works by parsing through the Windows Advanced Firewall log and
identifying IP addresses who connect so frequently that they appear suspicious.
If an IP address is found to be suspicious it gets added to deny rules on the
Windows Advanced Firewall. If too many suspicious IP addresses get blocked in
one session the script will send an email to a predefined administrator. Many
of the capabilities are configurable in a settings section of the script.

-------------------------------------------------------------------------------

INSTALLATION

    The process of installing is performed in 4 parts.

Part 1 - Create a GMail Account

    1.  Navigate your browser to https://accounts.google.com/SignUp.
    2.  Fill in the fields necessary to create a new GMail account.
    3.  Save the new user email and password, we will need it.

Part 2 - Install the Scripts

    1.  Run the distributable installer as a local administrator.
    2.  Follow the prompts to completion.

Part 3 - Configure the Settings

    1.  Navigate Windows Explorer to the installation directory at
        %PROGRAMFILES(X86)%\DenyHostsWin\bin.
    2.  Right-click on Settings.bat and choose "Run as administrator".
    3.  Change the settings to suit your environment.
	4.  Save your changes by clicking on Apply or OK.

-------------------------------------------------------------------------------


