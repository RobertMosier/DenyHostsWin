@ECHO OFF

SET TEST_WSF = "%USERFROFILE%\Documents\projects\DenyHostsWin\test\Test_DenyHostsWin.wsf"

CSCRIPT.EXE /JOB:Test_WinAdvFw %TEST_WSF%

CSCRIPT.EXE /JOB:Test_WinAdvFw_LogParse %TEST_WSF%

CSCRIPT.EXE /JOB:Test_IPv4Addr %TEST_WSF%

CSCRIPT.EXE /JOB:Test_GMailNotify %TEST_WSF%

CSCRIPT.EXE /JOB:Test_LogRotate %TEST_WSF%

CSCRIPT.EXE /JOB:Test_RegSettings %TEST_WSF%

PAUSE
