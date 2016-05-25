# ExcelDnaTest
ExcelDNA


4. Added loggin for the creation of the registry key for the xll
   try capturing any issues with Process monitor. 
	and possibly win debug. 
	https://blogs.msdn.microsoft.com/patricka/2010/11/30/troubleshooting-application-compatibility-issues-tools-tips-and-tricks/
	https://technet.microsoft.com/en-us/sysinternals/bb896645.aspx?f=255&MSPPError=-2147217396
	if that fails too there is windebug
	https://msdn.microsoft.com/en-us/windows/hardware/hh852365


5. Saving dumps
Sometimes, you may want to save dumps for analysis later.  Maybe it’s on a non-development machine or you want to collect several dumps. You can configure Windows 7 and Windows Server 2008 R2 to always generate and save a dump file.

Create a key named: HKLM\Software\Microsoft\Windows\Windows Error Reporting\LocalDumps
Dumps will default to the following directory: %LOCALAPPDATA%\CrashDumps
You can override the default with a DumpFolder value (REG_EXPAND_SZ)
You can also limit the number of dumps saved with a DumpCount value (DWORD)
