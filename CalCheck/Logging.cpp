// Logging functionality

#include "StdAfx.h"
#include "Logging.h"

BOOL g_bEmitErrorNum = false; // Emit Message EID and Error Number in resulting Error CSV for OffCAT
BOOL g_bVerbose = false; 
BOOL g_bTime = false;  
BOOL g_bCSV = false; //Don't want to make a large CSV file by default - user must request it
BOOL g_bMRMAPI = true; // we will assume it's there, and will spit an error if it's not.
CString g_szAppPath = _T(""); // Global App Path variable so I can use files I need a little easier
BOOL g_bServerMode = false;
BOOL g_bReport = false;
BOOL g_bOutput = false;
CString g_szOutputPath = "";


//Create files - CalCheck.log, CalCheckErr.csv and CALITEMS.CSV
HRESULT CreateLogFiles()
{
	CStdioFile LogFile;
	CString strLogFileName;
	CStdioFile ErrorFile;
	CString strErrorFileName;
	CString strLine;
	CStdioFile CSVFile;
	CString strCSVFileName;
	char szAppPath[MAX_PATH] = "";
	CString strAppDir;
	SYSTEMTIME stTime;
	CString strLocalTime;
	CFileStatus status;
	CFileException e;
	HRESULT hRes = S_OK;

	if(g_szOutputPath.IsEmpty())
	{
		// find where the executable is located and use that path for the log file
		::GetModuleFileNameA(0, szAppPath, sizeof(szAppPath)-1);
		strAppDir = szAppPath;
		strAppDir = strAppDir.Left(strAppDir.ReverseFind('\\'));
		g_szOutputPath = strAppDir;
	}
	else
	{
			strAppDir = g_szOutputPath;
	}
		
	strLogFileName = strAppDir;
	strErrorFileName = strAppDir;
	strCSVFileName = strAppDir;
	strLogFileName += _T("\\CalCheck.log");
	strErrorFileName += _T("\\CalCheckErr.csv");
	strCSVFileName += _T("\\CALITEMS.CSV");

	// if the log file is there (calcheck.log), then remove it so a new one can be created
	if(LogFile.GetStatus(strLogFileName, status))
	{
		try
		{
			LogFile.Remove(strLogFileName);
		}
		catch (CFileException* pEx)
		{
			hRes = pEx->m_lOsError;
			DebugPrint(DBGError,_T("Failed to remove CalCheck.log file. Check that it is not open in another program.\n"));
			//MessageBoxEx(NULL,_T("Could not remove CalCheck.log file. Check that it is not open in another program."),_T("CalCheck"), MB_ICONERROR | MB_OK, 0);
			goto Exit;
		}
	}

	// if the error file is there (calcheckerr.csv), then remove it so a new one can be created
	if(ErrorFile.GetStatus(strErrorFileName, status))
	{
		try
		{
			ErrorFile.Remove(strErrorFileName);
		}
		catch (CFileException* pEx)
		{
			hRes = pEx->m_lOsError;
			DebugPrint(DBGError,_T("Failed to remove CalCheckErr.csv file. Check that it is not open in another program.\n"));
			//MessageBoxEx(NULL,_T("Could not remove CalCheckErr.csv file. Check that it is not open in another program."),_T("CalCheck"), MB_ICONERROR | MB_OK, 0);
			goto Exit;
		}
	}

	// if the CSV file is there (calitems.csv), then remove it so a new one can be created
	if(CSVFile.GetStatus(strCSVFileName, status))
	{
		try
		{
			CSVFile.Remove(strCSVFileName);
		}
		catch (CFileException* pEx)
		{
			hRes = pEx->m_lOsError;
			DebugPrint(DBGError,_T("Failed to remove CALITEMS.CSV file. Check that it is not open in another program.\n"));
			//MessageBoxEx(NULL,_T("Could not remove CALITEMS.CSV file. Check that it is not open in another program."),_T("CalCheck"), MB_ICONERROR | MB_OK, 0);
			goto Exit;
		}
	}

	//create & open the new log file
	if(!LogFile.Open((LPCTSTR)strLogFileName, CFile::modeCreate | CFile::modeNoTruncate | CFile::modeWrite | CFile::shareDenyNone | CFile::typeText, &e))
	{
		hRes = e.m_lOsError;
		DebugPrint(DBGError,_T("Could not open or write to CalCheck.log file.\n"));
		//MessageBoxEx(NULL,_T("Could not open or write to CalCheck.log file."),_T("CalCheck"), MB_ICONERROR | MB_OK, 0);
		goto Exit;
	}

	//create & open the new error file
	if(!ErrorFile.Open((LPCTSTR)strErrorFileName, CFile::modeCreate | CFile::modeNoTruncate | CFile::modeWrite | CFile::shareDenyNone | CFile::typeText, &e))
	{
		hRes = e.m_lOsError;
		DebugPrint(DBGError,_T("Could not open or write to CalCheckErr.csv file.\n"));
		//MessageBoxEx(NULL,_T("Could not open or write to CalCheckErr.csv file."),_T("CalCheck"), MB_ICONERROR | MB_OK, 0);
		goto Exit;
	}

	if(g_bCSV)
	{
		//create & open the new CSV file
		if(!CSVFile.Open((LPCTSTR)strCSVFileName, CFile::modeCreate | CFile::modeNoTruncate | CFile::modeWrite | CFile::shareDenyNone | CFile::typeText, &e))
		{
			hRes = e.m_lOsError;
			DebugPrint(DBGError,_T("Could not open or write to CALITEMS.CSV file.\n"));
			//MessageBoxEx(NULL,_T("Could not open or write to CALITEMS.CSV file."),_T("CalCheck"), MB_ICONERROR | MB_OK, 0);
			goto Exit;
		}
	}

	WriteLogFile(_T("Calendar Checking Tool"), 0);
	WriteLogFile(_T("=======================\r\n"), 0);

	GetLocalTime(&stTime);
	strLocalTime.Format(_T("%02d/%02d/%d %02d:%02d:%02d%s"), 
						stTime.wMonth, 
						stTime.wDay, 
						stTime.wYear, 
						(stTime.wHour <= 12)?stTime.wHour:stTime.wHour-12,  // adjust to non-military time
						stTime.wMinute, 
						stTime.wSecond,
						(stTime.wHour <= 12)?_T("AM"):_T("PM"));

	WriteLogFile(_T("Start: ") + strLocalTime + ("\r\n"), 0);

Exit:
	return hRes;
}

//Remove files - CalCheck.csv and CALITEMS.CSV
void RemoveLogFiles()
{
	CStdioFile LogFile;
	CString strLogFileName;
	CStdioFile ErrorFile;
	CString strErrorFileName;
	CStdioFile CSVFile;
	CString strCSVFileName;
	CString strAppDir;
	CFileStatus status;


	strLogFileName = g_szOutputPath;
	strErrorFileName = g_szOutputPath;
	strCSVFileName = g_szOutputPath;
	strLogFileName += _T("\\CalCheck.log");
	strErrorFileName += _T("\\CalCheckErr.csv");
	strCSVFileName += _T("\\CALITEMS.CSV");

	// if the log file is there, then remove it so a new one can be created
	if(LogFile.GetStatus(strLogFileName, status))
	{
		LogFile.Remove(strLogFileName);
	}

	// if the error file is there, then remove it so a new one can be created
	if(ErrorFile.GetStatus(strErrorFileName, status))
	{
		ErrorFile.Remove(strErrorFileName);
	}

	// if the CSV file is there, then remove it so a new one can be created
	if(CSVFile.GetStatus(strCSVFileName, status))
	{
		CSVFile.Remove(strCSVFileName);
	}
}

HRESULT CreateServerLogFile()
{
	CStdioFile SvrLogFile;
	CString strSvrLogFileName;
	CString strLine;
	char szAppPath[MAX_PATH] = "";
	CString strAppDir;
	SYSTEMTIME stTime;
	CString strLocalTime;
	CFileStatus status;
	CFileException e;
	HRESULT hRes = S_OK;
	
	if(g_szOutputPath.IsEmpty())
	{
		// find where the executable is located and use that path for the log file
		::GetModuleFileNameA(0, szAppPath, sizeof(szAppPath)-1);
		strAppDir = szAppPath;
		strAppDir = strAppDir.Left(strAppDir.ReverseFind('\\'));
		g_szOutputPath = strAppDir;
	}
	else
	{
		strAppDir = g_szOutputPath;
	}

	strSvrLogFileName = strAppDir;
	strSvrLogFileName += _T("\\CalCheckMaster.log");


	// if the log file is there, then remove it so a new one can be created
	if(SvrLogFile.GetStatus(strSvrLogFileName, status))
	{
		try
		{
			SvrLogFile.Remove(strSvrLogFileName);
		}
		catch (CFileException* pEx)
		{
			hRes = pEx->m_lOsError;
			DebugPrint(DBGError,_T("Failed to remove CalCheckMaster.log file. Check that it is not open in another program.\n"));
			//MessageBoxEx(NULL,_T("Could not remove CalCheckMaster.log file. Check that it is not open in another program."),_T("CalCheck"), MB_ICONERROR | MB_OK, 0);
			goto Exit;
		}
	}

	//create & open the new log file
	if(!SvrLogFile.Open((LPCTSTR)strSvrLogFileName, CFile::modeCreate | CFile::modeNoTruncate | CFile::modeWrite | CFile::shareDenyNone | CFile::typeText, &e))
	{
		hRes = e.m_lOsError;
		DebugPrint(DBGError,_T("Could not open or write to CalCheckMaster.log file.\n"));
		//MessageBoxEx(NULL,_T("Could not open or write to CalCheckMaster.log file."),_T("CalCheck"), MB_ICONERROR | MB_OK, 0);
		goto Exit;
	}

	WriteSvrLogFile(_T("Calendar Checking Tool"));
	WriteSvrLogFile(_T("=======================\r\n"));
	WriteSvrLogFile(_T("Multi Mode File - shows flagged items from all the mailboxes from the input file.\r\n"));

	GetLocalTime(&stTime);
	strLocalTime.Format(_T("%02d/%02d/%d %02d:%02d:%02d%s"), 
						stTime.wMonth, 
						stTime.wDay, 
						stTime.wYear, 
						(stTime.wHour <= 12)?stTime.wHour:stTime.wHour-12,  // adjust to non-military time
						stTime.wMinute, 
						stTime.wSecond,
						(stTime.wHour <= 12)?_T("AM"):_T("PM"));

	WriteSvrLogFile(_T("Start: ") + strLocalTime + ("\r\n"));

Exit:
	return hRes;
	
}

//Open/Create and write to a Log File for this app
void WriteLogFile(LPCTSTR strWriteLine, BOOL bSrvLog)
{
	CStdioFile LogFile;
	CString strFileName;
	CString strLine;
	CFileException e;

	strFileName = g_szOutputPath;
	strFileName += _T("\\CalCheck.log");

	// open it up
	if(!LogFile.Open((LPCTSTR)strFileName, CFile::modeCreate | CFile::modeNoTruncate | CFile::modeWrite | CFile::shareDenyNone | CFile::typeText, &e))
	{
		MessageBoxEx(NULL,_T("Could not open or write to CalCheck.log file."),_T("CalCheck"), MB_ICONERROR | MB_OK, 0);
	}

	// put the text in the file and create space for the next line
	LogFile.SeekToEnd();
	LogFile.WriteString(strWriteLine);
	LogFile.Write("\r\n", 2);
	LogFile.Close();

	// if in Server Mode - and this is supposed to be written to the server log, then do it.
	if(g_bMultiMode)
	{
		if(bSrvLog)
		{
			WriteSvrLogFile(strWriteLine);
		}
	}

}

//Open/Create and write to a Error File for this app
void WriteErrorFile(LPCTSTR strWriteLine)
{
	CStdioFile ErrorFile;
	CString strErrorFileName;
	CString strLine;
	CFileException e;

	strErrorFileName = g_szOutputPath;
	strErrorFileName += _T("\\CalCheckErr.csv");

	// open it up
	if(!ErrorFile.Open((LPCTSTR)strErrorFileName, CFile::modeCreate | CFile::modeNoTruncate | CFile::modeWrite | CFile::shareDenyNone | CFile::typeText, &e))
	{
		MessageBoxEx(NULL,_T("Could not open or write to CalCheckErr.csv file."),_T("CalCheck"), MB_ICONERROR | MB_OK, 0);
	}

	// Write to the file
	ErrorFile.SeekToEnd();
	ErrorFile.WriteString(strWriteLine);
	// CSVFile.Write("\r\n", 2);      >>>>> For the CSV output, you must enter \r\n from the calling function when you want it.
	ErrorFile.Close();
}

//Open/Create and write to a CSV File for this app
void WriteCSVFile(LPCTSTR strWriteLine)
{
	if(g_bCSV)
	{
		CStdioFile CSVFile;
		CString strCSVFileName;
		CString strLine;
		CFileException e;

		strCSVFileName = g_szOutputPath;
		strCSVFileName += _T("\\CALITEMS.CSV");

		// open it up
		if(!CSVFile.Open((LPCTSTR)strCSVFileName, CFile::modeCreate | CFile::modeNoTruncate | CFile::modeWrite | CFile::shareDenyNone | CFile::typeText, &e))
		{
			MessageBoxEx(NULL,_T("Could not open or write to CALITEMS.CSV file."),_T("CalCheck"), MB_ICONERROR | MB_OK, 0);
		}

		// Write to the file
		CSVFile.SeekToEnd();
		CSVFile.WriteString(strWriteLine);
		// CSVFile.Write("\r\n", 2);      >>>>> For the CSV output, you must enter \r\n from the calling function when you want it.
		CSVFile.Close();
	} //if

}

//Open and write to the Server Mode Log File 
void WriteSvrLogFile(LPCTSTR strWriteLine)
{
	CStdioFile LogFile;
	CString strFileName;
	CString strLine;
	CFileException e;

	strFileName = g_szOutputPath;
	strFileName += _T("\\CalCheckMaster.log");

	// open it up
	if(!LogFile.Open((LPCTSTR)strFileName, CFile::modeCreate | CFile::modeNoTruncate | CFile::modeWrite | CFile::shareDenyNone | CFile::typeText, &e))
	{
		MessageBoxEx(NULL,_T("Could not open or write to log file."),_T("CalCheck"), MB_ICONERROR | MB_OK, 0);
	}

	// Do time stuff for the log
	// put the text in the file and create space for the next line
	LogFile.SeekToEnd();
	LogFile.WriteString(strWriteLine);
	LogFile.Write("\r\n", 2);
	LogFile.Close();
}

CString GetDNUser(LPTSTR szDN)
{
	CString strUser = (CString)szDN;
	CString strReturn = _T("");

	int iStrLength = strUser.GetLength();
	int iEquals = strUser.ReverseFind('=');
	int iCopy = ((iStrLength - iEquals)-1); // This "should" put us at the right position in the DN to get just the user (CN="user") 

	strReturn = strUser.Right(iCopy);

	return strReturn;		
}

// In server mode the app will append to the filenames.
void AppendFileNames(LPCTSTR strAppend)
{
	CStdioFile LogFile;
	CString strLogFileName;
	CString strLogFileNameNew;
	CStdioFile ErrorFile;
	CString strErrorFileName;
	CString strErrorFileNameNew;
	CString strLine;
	CStdioFile CSVFile;
	CString strCSVFileName;
	CString strCSVFileNameNew;
	CFileStatus status;
	CFileException e;

	strLogFileName = (CString)g_szOutputPath;
	strErrorFileName = (CString)g_szOutputPath;
	strCSVFileName = (CString)g_szOutputPath;
	strLogFileName += _T("\\CalCheck.log");
	strErrorFileName += _T("\\CalCheckErr.csv");
	strCSVFileName += _T("\\CALITEMS.CSV");

	// if the log file (CalCheck.log) is there, then rename/append it with the passed in string
	if(LogFile.GetStatus(strLogFileName, status))
	{		
		strLogFileNameNew = strLogFileName.Left(strLogFileName.ReverseFind('.'));
		strLogFileNameNew += ("_" + (CString)strAppend + ".log");

		// if the renamed log file is already there (from a past run), then remove it so a new one can be created
		if(LogFile.GetStatus(strLogFileNameNew, status))
		{
			LogFile.Remove(strLogFileNameNew);
		}

		LogFile.Rename(strLogFileName, strLogFileNameNew);
	}

	// if the error file (CalCheckErr.csv) is there, then rename/append it with the passed in string
	if(ErrorFile.GetStatus(strErrorFileName, status))
	{		
		strErrorFileNameNew = strErrorFileName.Left(strErrorFileName.ReverseFind('.'));
		strErrorFileNameNew += ("_" + (CString)strAppend + ".csv");

		// if the renamed error file is already there (from a past run), then remove it so a new one can be created
		if(ErrorFile.GetStatus(strErrorFileNameNew, status))
		{
			ErrorFile.Remove(strErrorFileNameNew);
		}

		ErrorFile.Rename(strErrorFileName, strErrorFileNameNew);
	}

	// if the CSV file (CALITEMS.CSV) is there, then rename/append it with the passed in string
	if(CSVFile.GetStatus(strCSVFileName, status))
	{
		strCSVFileNameNew = strCSVFileName.Left(strCSVFileName.ReverseFind('.'));
		strCSVFileNameNew += ("_" + (CString)strAppend + ".CSV");
		
		// if the renamed CSV file is already there (from a past run), then remove it so a new one can be created
		if(CSVFile.GetStatus(strCSVFileNameNew, status))
		{
			CSVFile.Remove(strCSVFileNameNew);
		}

		CSVFile.Rename(strCSVFileName, strCSVFileNameNew);
	}
}

// Convert dispidPropDefStream prop to its "Smart View" Text
BOOL ConvertPropDefStream(LPCTSTR strStream, CString* strRetError)
{
	CStdioFile PDSFile;
	CStdioFile RETFile;
	CString strPDSFileName;
	CString strRETFileName;
	CString strError = _T("");
	CFileStatus status;
	CFileException ex;
	HRESULT hr = S_OK;

	strPDSFileName = (CString)g_szOutputPath;
	strPDSFileName += _T("\\PDS.txt");
	strRETFileName = (CString)g_szOutputPath;
	strRETFileName += _T("\\PDSret.txt");

	// if these files exist, then need to delete them first...
	if (PDSFile.GetStatus(strPDSFileName, status))
		PDSFile.Remove(strPDSFileName);

	if (RETFile.GetStatus(strRETFileName, status))
		RETFile.Remove(strRETFileName);


	//create & open a new PDS blob file
	if(!PDSFile.Open((LPCTSTR)strPDSFileName, CFile::modeCreate | CFile::modeNoTruncate | CFile::modeWrite | CFile::shareDenyNone | CFile::typeText, &ex))
	{
		strError = _T("ERROR: Could not open or write PropDefStream property to local file. Check that you are able to write to the folder CalCheck is located in.");
		*strRetError = strError;
		return false;
	}

	// put the blob in the file and save/close it.
	PDSFile.SeekToEnd();
	PDSFile.WriteString(strStream);
	PDSFile.Close();

	// Now run MRMAPI to parse/convert the blob in the file
	CString strSwitches = " -p 12 -i \"" + strPDSFileName + "\" -o \"" + strRETFileName + "\"";
	hr = RunMRMAPI(strSwitches);

	if(g_bMRMAPI) // if MrMAPI is there, then need to check for other problems that may have occurred running it
	{
		if(hr != S_OK)
		{
			CString strMMError;
			strMMError.Format(_T("%X"), hr);
			strError = _T("ERROR: ") + strMMError + (" when parsing the PropDefStream data.");
			*strRetError = strError;
			return false; // problem with running MrMAPI
		}
		if(!RETFile.GetStatus(strRETFileName, status))
		{
			strError = _T("ERROR: The converted PropDefStream data file was not created. Check for the correct MAPI DLLs and that the user has correct permissions / context to write files here.");
			*strRetError = strError;
			return false; // MrMAPI didn't create the file
		}
	}

	// now that we're done with this iteration of the Input file, remove it so a new one can be created later
	if(PDSFile.GetStatus(strPDSFileName, status))
	{
		PDSFile.Remove(strPDSFileName);
	}

	return true;
}

// Convert the dispidApptRecur prop to its "Smart View" text
BOOL ConvertApptRecurBlob(LPCTSTR strBlob, CString* strRetError)
{
	CStdioFile ARFile;
	CStdioFile CVTFile;
	CString strARFileName;
	CString strCVTFileName;
	CString strRecurErr = _T("");
	CFileStatus status;
	CFileException e;
	HRESULT hr = S_OK;
	
	strARFileName = (CString)g_szOutputPath;
	strCVTFileName = (CString)g_szOutputPath;
	strARFileName += _T("\\ApptRecur.txt");  //temp name for the hex blob text file
	strCVTFileName += _T("\\arcvt.txt");     // temp name for the converted appt recurrence blob

	// Just in case the AR Blob file is there - go ahead and remove it. It should already be gone though...
	if(ARFile.GetStatus(strARFileName, status))
	{
		ARFile.Remove(strARFileName);
	}

	// if the Converted blob file is there, then remove it so a new one can be created
	if(CVTFile.GetStatus(strCVTFileName, status))
	{
		CVTFile.Remove(strCVTFileName);
	}

	//create & open a new AR blob file
	if(!ARFile.Open((LPCTSTR)strARFileName, CFile::modeCreate | CFile::modeNoTruncate | CFile::modeWrite | CFile::shareDenyNone | CFile::typeText, &e))
	{
		strRecurErr = _T("ERROR: Could not open or write recurrence property to local file. Check that you are able to write to the folder CalCheck is located in.");
		*strRetError = strRecurErr;
		return false;
	}

	// put the blob in the file and save/close it.
	ARFile.SeekToEnd();
	ARFile.WriteString(strBlob);
	ARFile.Close();

	// Now run MRMAPI to parse/convert the blob in the file
	CString strSwitches = " -p 2 -i \"" + strARFileName + "\" -o \"" + strCVTFileName + "\"";
	hr = RunMRMAPI(strSwitches);

	if(g_bMRMAPI) // if MrMAPI is there, then need to check for other problems that may have occurred running it
	{
		if(hr != S_OK)
		{
			CString strError;
			strError.Format(_T("%X"), hr);
			strRecurErr = _T("ERROR: ") + strError + (" when parsing the Recurrence data.");
			*strRetError = strRecurErr;
			return false; // problem with running MrMAPI
		}

		if(!CVTFile.GetStatus(strCVTFileName, status))
		{
			strRecurErr = _T("ERROR: The converted Recurrence data file was not created. Check for the correct MAPI DLLs and that the user has correct permissions / context to write files here.");
			*strRetError = strRecurErr;
			return false; // MrMAPI didn't create the file
		}
	}

	// now that we're done with this iteration of the AR blob file, remove it so a new one can be created later
	if(ARFile.GetStatus(strARFileName, status))
	{
		ARFile.Remove(strARFileName);
	}

	return true;
}

HRESULT RunMRMAPI(LPCTSTR strSwitches)
{
	// Run MRMAPI /? to see the list of switches that can be used

	STARTUPINFO si;
	PROCESS_INFORMATION pi;
	CStdioFile EXEFile;
	CFileStatus status;
	CString strEXE;
	CString strQuotedEXE;
	DWORD dwExitCode = 0;
	HRESULT hr = 0;

	strEXE = (CString)g_szAppPath;
	strQuotedEXE = "\"" + (CString)g_szAppPath;
	strEXE += _T("\\MrMAPI.exe");
	strQuotedEXE += _T("\\MrMAPI.exe\"");

	// If it's there, then run it. 
	if(EXEFile.GetStatus(strEXE, status))
	{
		

		CString strCmdline = strQuotedEXE + CString(strSwitches); 
		LPWSTR pwszCmdline = (LPWSTR)(LPCTSTR)strCmdline; // have to convert to LPWSTR from CString...

		ZeroMemory( &si, sizeof(si) );
		si.cb = sizeof(si);
		ZeroMemory( &pi, sizeof(pi) );

		if(CreateProcess( NULL, // No module name (use command line)
					   pwszCmdline, // Command line - contains the entire command to run
					   NULL, // Process handle not inheritable
					   NULL, // Thread handle not inheritable
					   FALSE, // Set handle inheritance to FALSE
					   0, // No creation flags
					   NULL, // Use parent's environment block
					   NULL, // Use parent's starting directory 
					   &si, // Pointer to STARTUPINFO structure
					   &pi  // Pointer to PROCESS_INFORMATION structure
					 ) != false)
		{
			dwExitCode = WaitForSingleObject(pi.hProcess, 1000); // need to wait for it to finish...
		}
		else
		{
			hr = GetLastError();
		}

		CloseHandle(pi.hProcess);
		CloseHandle(pi.hThread);

	}
	else // the EXE is not there, so error out and don't try this again
	{
		hr = 0x2; // File Not Found error
		g_bMRMAPI = false;
		WriteLogFile(_T("WARNING: MrMAPI.exe not in the same directory with CalCheck.exe. Recurrence data will not be parsed."), 0);
		WriteLogFile(_T("         Please place MrMAPI.exe in the same directory with CalCheck.exe to get additional data."), 0);
	}

	return hr;
}

// For removing the ApptRecur Blob and Converted ApptRecur files when the tool has finished
void RemoveTempFiles()
{
	CStdioFile ARFile;
	CString strARFileName;
	CStdioFile CVTFile;
	CString strCVTFileName;
	CFileStatus status;
	CStdioFile PDSFile;
	CString strPDSFileName;
	CStdioFile RETFile;
	CString strRETFileName;


	strARFileName = (CString)g_szOutputPath;
	strCVTFileName = (CString)g_szOutputPath;
	strARFileName += _T("\\ApptRecur.txt");  //temp name for the hex blob text file
	strCVTFileName += _T("\\arcvt.txt");     // temp name for the converted appt recurrence blob
	strPDSFileName = (CString)g_szOutputPath;
	strPDSFileName += _T("\\PDS.txt");
	strRETFileName = (CString)g_szOutputPath;
	strRETFileName += _T("\\PDSret.txt");


	// if the log file is there, then remove it 
	if(ARFile.GetStatus(strARFileName, status))
	{
		ARFile.Remove(strARFileName);
	}

	// if the CSV file is there, then remove it 
	if(CVTFile.GetStatus(strCVTFileName, status))
	{
		CVTFile.Remove(strCVTFileName);
	}

	if(PDSFile.GetStatus(strPDSFileName, status))
	{
		PDSFile.Remove(strPDSFileName);
	}

	if(RETFile.GetStatus(strRETFileName, status))
	{
		RETFile.Remove(strRETFileName);
	}


}

void DisplayError(
				  HRESULT hRes, 
				  LPCTSTR szErrorMsg, 
				  LPCSTR szFile,
				  int iLine,
				  LPCSTR szFunction)
{
	TCHAR szErrString[4096];
	HRESULT myHRes = NULL;

	LPTSTR szErrFormatString = NULL;
	if (FAILED(hRes))
	{
		szErrFormatString = _T("ERROR: %s, Code: 0x%08X, Function %hs, File %hs, Line %d\n");
	}
	else
	{
		szErrFormatString = _T("WARNING: %s, Code: 0x%08X, Function %hs, File %hs, Line %d\n");
	}
	myHRes = StringCchPrintf(
		szErrString,
		CCH(szErrString),
		szErrFormatString,
		szErrorMsg,
		hRes,
		szFunction,
		szFile,
		iLine);

	if (g_bVerbose)
		_tprintf(szErrString);
}

BOOL WarnHResFn(HRESULT hRes,LPCSTR szFunction,LPCTSTR szErrorMsg,LPCSTR szFile,int iLine)
{
	if (S_OK != hRes)
	{
		DisplayError(hRes,szErrorMsg,szFile,iLine,szFunction);
	}
	return SUCCEEDED(hRes);
}

void __cdecl DebugPrint(ULONG ulDbgLvl, LPCTSTR szMsg,...)
{
	if (DBGVerbose == ulDbgLvl && !g_bVerbose) return;
	if (DBGTime == ulDbgLvl && !g_bTime) return;
	HRESULT hRes = S_OK;
	
	va_list argList = NULL;
	va_start(argList, szMsg);

	if (DBGError == ulDbgLvl)
	{
		_tprintf(_T("ERROR: "));
	}
	else if (DBGHelp != ulDbgLvl) 
	{
		_tprintf(_T("INFO: "));
	}

	// print the time first - this is INFO
	if (DBGTime == ulDbgLvl)
	{
		SYSTEMTIME	stLocalTime;
		GetLocalTime(&stLocalTime); // Changed to "GetLocalTime" so we will get local instead of GMT (GetSystemTime)

		_tprintf(_T("%02d/%02d/%4d %02d:%02d:%02d%s "),
			stLocalTime.wMonth,
			stLocalTime.wDay,
			stLocalTime.wYear,
			(stLocalTime.wHour <= 12)?stLocalTime.wHour:stLocalTime.wHour-12,
			stLocalTime.wMinute,
			stLocalTime.wSecond,
			(stLocalTime.wHour <= 12)?_T("AM"):_T("PM"));
	}
	
	if (argList)
	{
		TCHAR szDebugString[4096];
		hRes = StringCchVPrintf(szDebugString, CCH(szDebugString), szMsg, argList);
		_tprintf(szDebugString);
	}
	else
		_tprintf(szMsg);
	va_end(argList);

}

BOOL CheckStringProp(LPSPropValue lpProp, ULONG ulPropType)
{
	if (PT_STRING8 != ulPropType && PT_UNICODE != ulPropType)
	{
		DebugPrint(DBGVerbose,_T("CheckStringProp: Called with invalid ulPropType of 0x%X\n"),ulPropType);
		return false;
	}
	if (!lpProp)
	{
		DebugPrint(DBGVerbose,_T("CheckStringProp: lpProp is NULL\n"));
		return false;
	}

	if (PT_ERROR == PROP_TYPE(lpProp->ulPropTag))
	{
		DebugPrint(DBGVerbose,_T("CheckStringProp: lpProp->ulPropTag is of type PT_ERROR\n"));
		return false;
	}

	if (ulPropType != PROP_TYPE(lpProp->ulPropTag))
	{
		DebugPrint(DBGVerbose,_T("CheckStringProp: lpProp->ulPropTag is not of type 0x%X\n"),ulPropType);
		return false;
	}

	if (NULL == lpProp->Value.LPSZ)
	{
		DebugPrint(DBGVerbose,_T("CheckStringProp: lpProp->Value.LPSZ is NULL\n"));
		return false;
	}

	if (PT_STRING8 == ulPropType && NULL == lpProp->Value.lpszA[0])
	{
		DebugPrint(DBGVerbose,_T("CheckStringProp: lpProp->Value.lpszA[0] is NULL\n"));
		return false;
	}

	if (PT_UNICODE == ulPropType && NULL == lpProp->Value.lpszW[0])
	{
		DebugPrint(DBGVerbose,_T("CheckStringProp: lpProp->Value.lpszW[0] is NULL\n"));
		return false;
	}

	return true;
}