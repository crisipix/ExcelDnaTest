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

An XLL is a DLL that exports several procedures that are called by Excel or the Excel Add-in Manager. These procedures are described briefly here and discussed in detail in Add-in Manager and XLL Interface Functions. All of these DLL callbacks start with the prefix xlAuto. Only one of these, the command xlAutoOpen, is required. It is called when the add-in is activated, and it is typically used to register XLL functions and commands with Excel and to do other initialization tasks. The function signatures and example implementations of all of the xlAuto functions are provided in later sections.
Even though xlAutoOpen is the only required one of these callbacks, your add-in may also need to export others depending on its behavior.
Excel 2007 introduced a new data type, XLOPER12, to accommodate larger grids and to support long Unicode strings. XLOPER12 is described later in this topic. Whereas xlAuto functions take or return the old data type XLOPER, new versions of these functions were introduced in Excel 2007 that use XLOPER12 data types. With the exception of xlAutoFree12, which you must sometimes implement to avoid XLOPER12 memory leaks, you can safely omit all the version 12 xlAuto functions, in which case, starting in Excel 2007, Excel calls the XLOPER versions.
xlAutoOpen
Excel calls the xlAutoOpen function whenever the XLL is activated. The add-in will be activated at the start of an Excel session if it was active in the last Excel session that ended normally. The add-in is activated if it is loaded during an Excel session. The add-in can be deactivated and reactivated during an Excel session, and the function is called on reactivation.
You should use xlAutoOpen to register XLL functions and commands, initialize data structures, customize the user interface, and so on.
If your add-in implements and exports the xlAutoRegister function or the xlAutoRegister12 function, Excel might attempt to activate and register a function or command without first calling the xlAutoOpen function. In this case, you should ensure that your add-in is sufficiently initialized for your function or command to work properly. If it is not, you should either fail the attempt to register the function or command, or carry out the necessary initialization.
xlAutoClose
Excel calls the xlAutoClose function whenever the XLL is deactivated. The add-in will be deactivated when an Excel session ends normally. If the user deactivates the add-in during an Excel session, the function is called.
You should use xlAutoClose to unregister functions and commands, release resources, undo customizations, and so on.
https://msdn.microsoft.com/en-us/library/office/bb687861.aspx
