; DenyHostsWin-dist-v18_03.nsi
;
; It will install DenyHostsWin into C:\Program Files\DenyHostsWin

;--------------------------------

!define VERSION "18.03"
!define VERSION_PATH "18_03"
!define APP_NAME "DenyHostsWin"

; The name of the installer
Name "${APP_NAME}-dist-v${VERSION_PATH}"

Caption "${APP_NAME} Installer"
Icon "${NSISDIR}\Contrib\Graphics\Icons\win-install.ico"

; The file to write
OutFile "dist\${APP_NAME}-dist-v${VERSION_PATH}.exe"

; The default installation directory
InstallDir $PROGRAMFILES\${APP_NAME}

; License
LicenseText "GNU General Public License v2.0"
LicenseData "LICENSE.txt"

; Registry key to check for directory (so if you install again, it will 
; overwrite the old one automatically)
InstallDirRegKey HKLM "Software\${APP_NAME}" "Install_Dir"

; Request application privileges for Windows Vista
RequestExecutionLevel admin

;--------------------------------

; Pages

Page license
Page instfiles

UninstPage uninstConfirm
UninstPage instfiles

;--------------------------------

; The stuff to install
Section "${APP_NAME} (required)"

  SectionIn RO
  
  ; Write files to install directory
  SetOutPath $INSTDIR
  File "msg-template.txt"
  File "LICENSE.txt"
  File "README.txt"
  
  SetOutPath $INSTDIR\bin
  File "bin\${APP_NAME}.bat"
  File "bin\Settings.bat"
  
  SetOutPath $INSTDIR\src
  File "src\${APP_NAME}.wsf"
  File "src\Settings.hta"
  File "src\${APP_NAME}.vbs"
  File "src\GMailNotify.vbs"
  File "src\HelperLib.vbs"
  File "src\IPv4Addr.vbs"
  File "src\LogRotate.vbs"
  File "src\RegSettings.vbs"
  File "src\SettingsEditor.vbs"
  File "src\WinAdvFw.vbs"
  File "src\WinAdvFw_LogParse.vbs"
  File "src\Settings.vbs"
  
  ; Create the scheduled task
  ;Exec '"$SYSDIR\cscript.exe" "$INSTDIR\ScheduleTask.vbs"'
  ;Exec '"$SYSDIR\schtasks.exe" /Create /RU "NT AUTHORITY\SYSTEM" /RL HIGHEST /SC DAILY /TN ${APP_NAME} /TR "C:\Windows\System32\cscript.exe C:\Program Files (x86)\${APP_NAME}\${APP_NAME}-v${VERSION_PATH}.wsf" /ST 20:00:00 /NP /F'
  Exec '"$SYSDIR\schtasks.exe" /Create /RU "NT AUTHORITY\SYSTEM" /RL HIGHEST /SC DAILY /TN ${APP_NAME} /TR "%PROGRAMFILES(X86)%\${APP_NAME}\bin\${APP_NAME}.bat" /ST 20:00:00 /NP /F'
  
  ; Enable Windows firewall logging for allowed connections
  Exec '"$SYSDIR\netsh.exe" advfirewall set allprofiles logging allowedconnections enable'
  ; Expand the max file size of the firewall log
  Exec '"$SYSDIR\netsh.exe" advfirewall set allprofiles logging maxfilesize 16384'
  
  ; Write default settings into the registry
  WriteRegStr	HKLM "SOFTWARE\${APP_NAME}"								"CurrentVersion"			"${VERSION}"
  WriteRegDWORD	HKLM "SOFTWARE\${APP_NAME}\${VERSION}"					"Threshold"					0x00000014
  WriteRegStr	HKLM "SOFTWARE\${APP_NAME}\${VERSION}\FirewallRules"	"FwRuleBaseName"			"AUTO-BLACKLISTED IP ADDRESSES"
  WriteRegStr	HKLM "SOFTWARE\${APP_NAME}\${VERSION}\FirewallRules"	"FwRuleGroupName"			"AUTO-BLACKLIST"
  WriteRegStr	HKLM "SOFTWARE\${APP_NAME}\${VERSION}\FirewallRules"	"FwRuleDescription"			"Remote IP addresses blocked automatically"
  WriteRegStr	HKLM "SOFTWARE\${APP_NAME}\${VERSION}\FirewallRules"	"FwRuleLocalPorts"			"0-65535"
  WriteRegStr	HKLM "SOFTWARE\${APP_NAME}\${VERSION}\FirewallLog"		"FwLogFile"					"%windir%\system32\logfiles\firewall\0.pfirewall.log"
  WriteRegStr	HKLM "SOFTWARE\${APP_NAME}\${VERSION}\FirewallLog"		"FwPort"					"3389"
  WriteRegDWORD	HKLM "SOFTWARE\${APP_NAME}\${VERSION}\EMailNotify"		"NotifyOnThreshold"			0x00000001
  WriteRegDWORD	HKLM "SOFTWARE\${APP_NAME}\${VERSION}\EMailNotify"		"NotifyThreshold"			0x00000005
  WriteRegStr	HKLM "SOFTWARE\${APP_NAME}\${VERSION}\EMailNotify"		"NotifyEmailFrom"			"$\"${APP_NAME}$\" <relay.example@gmail.com>"
  WriteRegStr	HKLM "SOFTWARE\${APP_NAME}\${VERSION}\EMailNotify"		"NotifyEmailTo"				"$\"Administrator$\" <admin@example.com>"
  WriteRegStr	HKLM "SOFTWARE\${APP_NAME}\${VERSION}\EMailNotify"		"NotifyEmailSubject"		"ALERT from ${APP_NAME} server.domain.tld"
  WriteRegStr	HKLM "SOFTWARE\${APP_NAME}\${VERSION}\EMailNotify"		"NotifyEmailMsgTemplate"	"%PROGRAMFILES(X86)%\DenyHostsWin\msg-template.txt"
  WriteRegStr	HKLM "SOFTWARE\${APP_NAME}\${VERSION}\EMailNotify"		"NotifySMTPAuthUser"		"relay.example@gmail.com"
  WriteRegStr	HKLM "SOFTWARE\${APP_NAME}\${VERSION}\EMailNotify"		"NotifySMTPAuthPass"		"your_smtp_password"
  WriteRegStr	HKLM "SOFTWARE\${APP_NAME}\${VERSION}\LogRotate"		"LrLogPath"					"%systemroot%\system32\LogFiles\Firewall"
  WriteRegStr	HKLM "SOFTWARE\${APP_NAME}\${VERSION}\LogRotate"		"LrLogFilename"				"pfirewall.log"
  WriteRegDWORD	HKLM "SOFTWARE\${APP_NAME}\${VERSION}\LogRotate"		"LrMaxArchives"				0x00000005
  
  ; Write the installation path into the registry
  WriteRegStr HKLM "SOFTWARE\${APP_NAME}" "Install_Dir" "$INSTDIR"
  
  ; Write the uninstall keys for Windows
  WriteRegStr	HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayName"		"${APP_NAME}"
  WriteRegStr	HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "UninstallString"	'"$INSTDIR\uninstall.exe"'
  WriteRegDWORD	HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "NoModify" 1
  WriteRegDWORD	HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "NoRepair" 1
  WriteUninstaller "uninstall.exe"
  
SectionEnd

;--------------------------------

; Uninstaller

UninstallText "This will uninstall ${APP_NAME}."
UninstallIcon "${NSISDIR}\Contrib\Graphics\Icons\nsis1-uninstall.ico"

Section "Uninstall"
  
  ; Remove registry keys
  DeleteRegKey	HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}"
  DeleteRegKey	HKLM "SOFTWARE\${APP_NAME}"
  
  ; Remove files and uninstaller
  Delete "$INSTDIR\msg-template.txt"
  Delete "$INSTDIR\LICENSE.txt"
  Delete "$INSTDIR\README.txt"
  Delete "$INSTDIR\uninstall.exe"

  Delete "$INSTDIR\bin\${APP_NAME}.bat"
  Delete "$INSTDIR\bin\Settings.bat"
  
  Delete "$INSTDIR\src\${APP_NAME}.wsf"
  Delete "$INSTDIR\src\Settings.hta"
  Delete "$INSTDIR\src\${APP_NAME}.vbs"
  Delete "$INSTDIR\src\GMailNotify.vbs"
  Delete "$INSTDIR\src\HelperLib.vbs"
  Delete "$INSTDIR\src\IPv4Addr.vbs"
  Delete "$INSTDIR\src\LogRotate.vbs"
  Delete "$INSTDIR\src\RegSettings.vbs"
  Delete "$INSTDIR\src\SettingsEditor.vbs"
  Delete "$INSTDIR\src\WinAdvFw.vbs"
  Delete "$INSTDIR\src\WinAdvFw_LogParse.vbs"
  Delete "$INSTDIR\src\Settings.vbs"
  
  ; Remove scheduled task
  Exec '"$SYSDIR\schtasks.exe" /Delete /TN ${APP_NAME} /F'
  
  ; Disable Windows firewall logging for allowed connections
  Exec '"$SYSDIR\netsh.exe" advfirewall set allprofiles logging allowedconnections disable'
  ; Return the max file size of the firewall log to default
  Exec '"$SYSDIR\netsh.exe" advfirewall set allprofiles logging maxfilesize 4096'

  ; Remove directories used
  RMDir "$INSTDIR\bin"
  RMDir "$INSTDIR\src"
  RMDir "$INSTDIR"

SectionEnd
