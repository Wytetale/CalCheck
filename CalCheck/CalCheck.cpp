// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
// Copyright 2011 Microsoft Corporation.  All Rights Reserved.
//
//  MODULE:   CalCheck.cpp
//
//  PURPOSE:  Start the CalCheck App
//
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "CalCheck.h"
#include "Defaults.h"
#include "MapiStoreFunctions.h"
#include "ProcessCalendar.h"
#include "Logging.h"

/////////////////////////////////////////////////////////////////////////////
// The one and only application object

HRESULT LoadMAPIVer(ULONG ulMAPIVer);
HRESULT FindAndLoadMAPI(ULONG *ulVer);
HRESULT LoopMailboxes(LPMAPISESSION lpMAPISession);
HRESULT ProcessMailbox(LPTSTR szDN, LPMDB lpMDB, LPMAPISESSION lpMAPISession);
HRESULT GetProxyAddrs(LPTSTR szDN, LPMAPISESSION lpMAPISession);
BOOL ReadConfigFile();
BOOL ReadMBXListFile(LPSTR lpszFileName);
VOID SelectTest(CString strTest);

MYOPTIONS ProgOpts;
BOOL g_bClearLogs = true;
CString g_MBXDisplayName = _T("");
CString g_MBXLegDN = _T("");
CString g_UserDisplayName = _T("");
CString g_UserLegDN = _T("");
CStringList AddrList; // For Proxy Addresses
TESTLIST CalCheckTests; // Read in from CalCheck.cfg
CStringList MBXInfoList;
BOOL g_bMultiMode = false;
BOOL g_bOtherMBXMode = false;
BOOL g_bInput = false;
CString g_szInputPath = _T("");

const WCHAR WszMAPISystemDrivePath[] = L"%s%s%s";
const WCHAR szMAPISystemDrivePath[] =  L"%hs%hs%ws";
static const WCHAR WszOlMAPI32DLL[] =  L"olmapi32.dll";
static const WCHAR WszMSMAPI32DLL[] =  L"msmapi32.dll";
static const WCHAR WszMapi32[] =       L"mapi32.dll";


void DisplayUsage()
{
	DebugPrint(DBGHelp,_T("Usage:\n"));
	DebugPrint(DBGHelp,_T("   You can edit the CalCheck.cfg file to turn specific tests on or off.\n"));
	DebugPrint(DBGHelp,_T("\n"));
	DebugPrint(DBGHelp,_T("   CalCheck [-P <profilename>] [-L <path\\file>] [-M <LegacyExchangeDN>] [-N <Display Name>] [-O <path>] [-C <Version>] [-A] [-F] [-R] [-V]\n"));
	DebugPrint(DBGHelp,_T("   CalCheck -?\n"));
	DebugPrint(DBGHelp,_T("\n"));
	DebugPrint(DBGHelp,_T("   -P   <Profile name> (if absent, will prompt for profile)\n"));
	DebugPrint(DBGHelp,_T("   -L   <List File> (file including Name and LegDN) of mailbox(es) to check\n"));
	DebugPrint(DBGHelp,_T("   -M   <Mailbox DN> Used with -N (if specified, only process the mailbox specified.)\n"));
	DebugPrint(DBGHelp,_T("   -N   <DisplayName> Used with -M (if specified, only process the mailbox specified)\n"));
	DebugPrint(DBGHelp,_T("   -O   <Output Path> (path to place output files - default is the current directory)\n"));
	DebugPrint(DBGHelp,_T("   -C   <Version> Click To Run scenario with Office 2013 - load a specific MAPI version\n"));
	DebugPrint(DBGHelp,_T("   -A   All calendar items output to CALITEMS.CSV\n"));
	DebugPrint(DBGHelp,_T("   -F   Create CalCheck Folder and move flagged error items there\n"));
	DebugPrint(DBGHelp,_T("   -R   Put a Report message in the Inbox with the CalCheck.csv file\n"));
	DebugPrint(DBGHelp,_T("   -V   Verbose output to the command window\n"));  // << Combined Timing and Verbose together
	DebugPrint(DBGHelp,_T("   -?   Print this message\n"));
	DebugPrint(DBGHelp,_T("\n"));
	DebugPrint(DBGHelp,_T("There will be resulting CalCheck.log and CalCheckErr.csv files that show\n"));
	DebugPrint(DBGHelp,_T("potential problems and/or items to fix or remove as well as processing information.\n"));
	DebugPrint(DBGHelp,_T("\n"));
	DebugPrint(DBGHelp,_T("\n"));
	DebugPrint(DBGHelp,_T("Examples\n"));
	DebugPrint(DBGHelp,_T("========\n"));
	DebugPrint(DBGHelp,_T("\n"));
	DebugPrint(DBGHelp,_T("   Default - Prompt for a profile, and process the mailbox in that profile:\n"));
	DebugPrint(DBGHelp,_T("      CalCheck\n"));
	DebugPrint(DBGHelp,_T("\n"));
	DebugPrint(DBGHelp,_T("   Process just the mailbox in MyProfile:\n"));
	DebugPrint(DBGHelp,_T("      CalCheck -P MyProfile\n"));
	DebugPrint(DBGHelp,_T("\n"));
	DebugPrint(DBGHelp,_T("   Process List of mailboxes in \"C:\\Directory\\List.txt\":\n"));
	DebugPrint(DBGHelp,_T("      CalCheck -L \"C:\\Directory\\List.txt\" \n"));
	DebugPrint(DBGHelp,_T("      \"List.txt\" needs to be in the format of Get-Mailbox | fl output that includes\n"));
	DebugPrint(DBGHelp,_T("      the LegacyExchangeDN and Name for each mailbox, like: \n"));
	DebugPrint(DBGHelp,_T("         Name             : Display Name\n"));
	DebugPrint(DBGHelp,_T("         LegacyExchangeDN : /o=ORG/ou=AdminGroup/cn=Recipients/cn=mailbox\n"));
	DebugPrint(DBGHelp,_T("\n"));
	DebugPrint(DBGHelp,_T("   Process a mailbox this user has full rights to:\n"));
	DebugPrint(DBGHelp,_T("      CalCheck -M <LegacyExchangeDN of the mailbox> -N <Display Name of the mailbox>\n"));
	DebugPrint(DBGHelp,_T("\n"));
	DebugPrint(DBGHelp,_T("   Process a mailbox and move error items to the CalCheck folder in the mailbox,\n"));
	DebugPrint(DBGHelp,_T("   and place a report message in the Inbox:\n"));
	DebugPrint(DBGHelp,_T("      CalCheck -F -R\n"));
	DebugPrint(DBGHelp,_T("\n"));
	DebugPrint(DBGHelp,_T("   Process a mailbox based on a specific profile and version of MAPI for Click To Run scenario:\n"));
	DebugPrint(DBGHelp,_T("      CalCheck -C <Outlook version - like 2007, 2010, 2013> -P MyProfile\n"));
	DebugPrint(DBGHelp,_T("\n"));
	DebugPrint(DBGHelp,_T("   Print this message\n"));
	DebugPrint(DBGHelp,_T("      CalCheck -?\n"));
	DebugPrint(DBGHelp,_T("\n==================================================================\n"));
	DebugPrint(DBGHelp,_T("\nPrivacy Note:\n"));
	DebugPrint(DBGHelp,_T("The data files produced by CalCheck can contain PII such as e-mail addresses.\n")); 
	DebugPrint(DBGHelp,_T("Please delete these files from your system after analysis and/or supplying them to Microsoft for analysis.\n"));
	DebugPrint(DBGHelp,_T("For more information on Microsoft's privacy standards and practices, please go to http://privacy.microsoft.com/en-us/default.mspx. \n"));
	DebugPrint(DBGHelp,_T("\n"));
}

// scans an arg and returns the string or hex number that it represents
void GetArg(char* szArgIn, char** lpszArgOut,ULONG* ulArgOut)
{
	ULONG ulArg = NULL;
	char* szArg = NULL;
	LPSTR szEndPtr = NULL;
	ulArg = strtoul(szArgIn,&szEndPtr,16);

	// if szEndPtr is pointing to something other than NULL, this must be a string
	if (!szEndPtr || *szEndPtr)
	{
		ulArg = NULL;
		szArg = szArgIn;
	}

	if (lpszArgOut) *lpszArgOut = szArg;
	if (ulArgOut)   *ulArgOut   = ulArg;
}

BOOL ParseArgs(int argc, char * argv[], MYOPTIONS * pRunOpts)
{
	if (!pRunOpts) return false;

	CString strNoOpt;
	
	// Clear our options list
	ZeroMemory(pRunOpts,sizeof(MYOPTIONS));

	// If no args at all - then just bring up the profile dialog
	if (1 == argc)
	{
		pRunOpts->lpszProfileName = "";
	}
	
	for (int i = 1; i < argc; i++)
	{
		switch (argv[i][0])
		{
		case '-':
			{
				if (0 == argv[i][1])
				{
					// Bad argument - get out of here
					return false;
				}
				switch (tolower(argv[i][1]))
				{
				case 'v':
					g_bVerbose = true;
					g_bTime = true;
					break;
				case 'e':
					g_bEmitErrorNum = true;
					break;
				case 'i':
					if (i+1 < argc)
					{
						g_szInputPath = argv[i+1];
						i++;
						g_bInput = true;
					}
					else return false;
					break;
				case 'a':
					g_bCSV = true;
					break;
				case 'f':
					g_bFolder = true;
					break;
				case 'r':
					g_bReport = true;
					break;
				case 'l':
					if (i+1 < argc)
					{
						pRunOpts->lpszFileName = argv[i+1];
						i++;
					}
					else return false;
					break;
				case 'p':
					if (i+1 <= argc)
					{
						// If nothing after -P then just bring up the dialog - no profile was specified
						if (i+1 == argc)
						{
							pRunOpts->lpszProfileName = "";
						}
						else
						{
							pRunOpts->lpszProfileName = argv[i+1];
						}
						i++;
					}
					else return false;
					break;
				case 'm':
					if (i+1 < argc)
					{
						pRunOpts->lpszMailboxDN = argv[i+1];
						i++;
					}
					else return false;
					break;
				case 'n':
					if (i+1 < argc)
					{
						pRunOpts->lpszDisplayName = argv[i+1];
						i++;
					}
					else return false;
					break;
				case 'c':
					if (i+1 < argc)
					{
						pRunOpts->lpszMAPIVer = argv[i+1];
						i++;
					}
					else return false;
					break;
				case 'o':
					if (i+1 < argc)
					{
						g_szOutputPath = argv[i+1];
						i++;
						g_bOutput = true;
					}
					else return false;
					break;
				case '?':
				case 'h':
				default:
					// display help
					return FALSE;
					break;
				}
			}
			break;
		
		default:
			return false;
			break;
		}
	}

	// Didn't fail - return true
	return true;
}

void main(int argc, char* argv[])
{
	char szAppPath[MAX_PATH] = "";
	CString strAppDir = "";
	BOOL bReadConfig = true;
	LPTSTR	lpszFileName = NULL;
	LPTSTR	lpszProfileName = NULL;
	LPTSTR  lpszMailboxDN = NULL;
	LPTSTR	lpszDisplayName = NULL;
	LPTSTR  lpszMAPIVer = NULL;
	CString strHR = _T("");
	ULONG	ulMAPIVer = 0;
	HRESULT hRes = S_OK;
	ULONG	ulCodePage = 0;


	// find where the executable is located and use that path for the Application
	::GetModuleFileNameA(0, szAppPath, sizeof(szAppPath)-1);
	strAppDir = szAppPath;
	strAppDir = strAppDir.Left(strAppDir.ReverseFind('\\'));
	g_szAppPath = strAppDir;
	
	DebugPrint(DBGGeneric,_T("Calendar Checking Tool\n"));
	DebugPrint(DBGGeneric,_T("=======================\n"));
	DebugPrint(DBGGeneric,_T("Checks Outlook calendars and items, and reports problems.\n"));
	DebugPrint(DBGGeneric,_T("For help, use the -? command switch.\n"));
	DebugPrint(DBGGeneric,_T("\n"));

	if (!ParseArgs(argc, argv, &ProgOpts))
	{
		DisplayUsage();
		return;
	}

	if (ProgOpts.lpszMailboxDN)
	{
		if (!ProgOpts.lpszDisplayName)
		{
			DebugPrint(DBGError,_T("You must use the -M and -N switches together.\n"));
			DebugPrint(DBGGeneric,_T("\n"));
			DisplayUsage();
			return;
		}
	}

	if (ProgOpts.lpszDisplayName)
	{
		if (!ProgOpts.lpszMailboxDN)
		{
			DebugPrint(DBGError,_T("You must use the -M and -N switches together.\n"));
			DebugPrint(DBGGeneric,_T("\n"));
			DisplayUsage();
			return;
		}
	}

#ifdef _UNICODE
	if (ProgOpts.lpszFileName) AnsiToUnicode(ProgOpts.lpszFileName,&lpszFileName);
	if (ProgOpts.lpszProfileName) AnsiToUnicode(ProgOpts.lpszProfileName,&lpszProfileName);
	if (ProgOpts.lpszMailboxDN) AnsiToUnicode(ProgOpts.lpszMailboxDN,&lpszMailboxDN);
	if (ProgOpts.lpszDisplayName) AnsiToUnicode(ProgOpts.lpszDisplayName,&lpszDisplayName);
	if (ProgOpts.lpszMAPIVer) AnsiToUnicode(ProgOpts.lpszMAPIVer,&lpszMAPIVer);
#else
	lpszFileName = ProgOpts.lpszFileName;
	lpszProfileName = ProgOpts.lpszProfileName;
	lpszMailboxDN = ProgOpts.lpszMailboxDN;
	lpszDisplayName = ProgOpts.lpszDisplayName;
	lpszMAPIVer = ProgOpts.lpszMAPIVer;
#endif

	if (ProgOpts.lpszMAPIVer) // load a specific version of MAPI for C2R scenarios.
	{
		CString strMAPIVer = lpszMAPIVer;

		//ensure version string is valid
		if (!(strMAPIVer.Compare(_T("2007"))) || !(strMAPIVer.Compare(_T("12"))))
		{
			ulMAPIVer = 12;
		}

		if (!(strMAPIVer.Compare(_T("2010"))) || !(strMAPIVer.Compare(_T("14"))))
		{
			ulMAPIVer = 14;
		}

		if (!(strMAPIVer.Compare(_T("2013"))) || !(strMAPIVer.Compare(_T("15"))))
		{
			ulMAPIVer = 15;
		}

		if (!(strMAPIVer.Compare(_T("2016"))) || !(strMAPIVer.Compare(_T("16"))))
		{
			ulMAPIVer = 16;
		}

		if (ulMAPIVer == 0)
		{
			DebugPrint(DBGError,_T("You must specify a valid and supported Outlook version when using the -C switch.\n"));
			DebugPrint(DBGError,_T("Outlook 2013 Click To Run is only supported with Outlook 2007 or 2010 installed on the machine.\n"));
			DebugPrint(DBGGeneric,_T("\n"));
			DisplayUsage();
			return;
		}

		// Now load the specific MAPI version
		hRes = LoadMAPIVer(ulMAPIVer);
		/*if (FAILED(hRes))
		{
			DebugPrint(DBGError,_T("Failed to load the correct version of MAPI. Check the install. Error: 0x%08X\n"), hRes);
			goto Exit;
		}*/

	}// ProgOpts.lpszMAPIVer
	else
	{
		hRes = FindAndLoadMAPI(&ulMAPIVer); // find the installed build of Outlook/MAPI and load it
		/*if (FAILED(hRes))
		{
			DebugPrint(DBGError,_T("Failed to load the correct version of MAPI. Check the install. Error: 0x%08X\n"), hRes);
			goto Exit;
		}*/
	}


	// Create Output Directory if specified
	if (g_bOutput)
	{
		if (!(CreateDirectory(g_szOutputPath, NULL)))
		{
			hRes = GetLastError();
			if (hRes == ERROR_ALREADY_EXISTS)
			{
				hRes = S_OK; // the directory is there so it can be used
			}
			else
			{			
				DebugPrint(DBGError,_T("Failed to create output directory. Please check and try again. Error: 0x%08X\n"), hRes);
				DebugPrint(DBGGeneric,_T("\n"));
				DisplayUsage();
				goto Exit;
			}
		}
	}

	// Create new log/output files
	hRes = CreateLogFiles();
	if (!hRes == S_OK)
	{
		DebugPrint(DBGError,_T("Failed creating required output files. Please close any opened files and try again. Error: 0x%08X\n"), hRes);
		goto Exit;
	}

	if (g_bOutput)
	{
		DebugPrint(DBGVerbose,_T("Output directory is set to " + g_szOutputPath + "\n"));
		DebugPrint(DBGVerbose,_T("\n"));
	}
	
	DebugPrint(DBGTime,_T("Begin process\n"));
	DebugPrint(DBGVerbose,_T("\n"));

	// Read in CalCheck.cfg to get the list of tests to perform / report
	bReadConfig = ReadConfigFile();
	if(!bReadConfig)
	{
		DebugPrint(DBGGeneric,_T("Could not find or access the CalCheck.cfg file.\n"));
		DebugPrint(DBGGeneric,_T("Make sure it exists and can be accessed in the directory with CalCheck.exe and try again.\n"));
		goto Exit;
	}


	LPMAPISESSION lpMAPISession = NULL;

	WC_H(MAPIInitialize(NULL));

	hRes = MAPILogonEx(0, (LPTSTR) ProgOpts.lpszProfileName, NULL,
		MAPI_LOGON_UI | MAPI_NEW_SESSION |  MAPI_EXPLICIT_PROFILE,
		&lpMAPISession);
	if (FAILED(hRes))
	{
		if (MAPI_E_USER_CANCEL == hRes)
		{
			DebugPrint(DBGGeneric,_T("User cancelled the process.\n"));
			WriteLogFile(_T("User cancelled the process."), 0);
			hRes = S_OK;
		}
		else
		{
			strHR.Format(_T("%08X"), hRes);
			DebugPrint(DBGError,_T("MAPILogonEx returned error 0x%08X\n"), hRes);
			DebugPrint(DBGError,_T("Check access to the mailbox.\n"));
			WriteLogFile(_T("ERROR: MAPILogonEx failed with error 0x" + strHR + ". Check access to the mailbox."), 0);
			g_bClearLogs = false;
		}
	}
	
	if (lpMAPISession)
	{
		LPMDB			lpPrimaryMDB = NULL;
		LPMDB			lpMDB = NULL;
		CString			szUserDN;
		CString			szDisplayName;
		ULONG			cbEntryID = 0;
		LPENTRYID		lpEntryID = NULL;
		LPMAPIPROP		lpIdentity = NULL;
		LPSPropValue	lpMailboxName = NULL;
		LPSPropValue	lpDisplayName = NULL;
		ULONG			ulObjType = NULL;

		// need to get the current user's (the one that is running this) DN and display name
		WC_H(OpenMessageStoreGUID(lpMAPISession, 
								  pbExchangeProviderPrimaryUserGuid, 
								  &lpMDB));
		if (lpMDB)
		{
			hRes = lpMAPISession->QueryIdentity(&cbEntryID,	&lpEntryID);

			if (SUCCEEDED(hRes) && lpEntryID)
			{
				hRes = lpMAPISession->OpenEntry(cbEntryID,
												lpEntryID,
												NULL,
												NULL,
												&ulObjType,
												(LPUNKNOWN*)&lpIdentity);

				if (SUCCEEDED(hRes) && lpIdentity)
				{			
					hRes = HrGetOneProp(lpIdentity, PR_EMAIL_ADDRESS, &lpMailboxName);
					if (SUCCEEDED(hRes) && lpMailboxName)
					{
						szUserDN = lpMailboxName->Value.LPSZ;
					}
					else szUserDN = _T("Unknown");

					hRes = HrGetOneProp(lpIdentity, PR_DISPLAY_NAME, &lpDisplayName);
					if(SUCCEEDED(hRes) && lpDisplayName)
					{
						szDisplayName = lpDisplayName->Value.LPSZ;
					}
					else szDisplayName = _T("Unknown");

					// set these globals for checking other things later...
					g_UserLegDN = szUserDN;
					g_UserDisplayName = szDisplayName;
				}
			}
		}

		if (ProgOpts.lpszFileName)
		{
			g_bMultiMode = true;
			BOOL bSuccess = true;
			bSuccess = ReadMBXListFile(ProgOpts.lpszFileName);			
			if(bSuccess)
			{
				WC_H(LoopMailboxes(lpMAPISession));
			}
			else
			{
				DebugPrint(DBGError,_T("Failed to read Mailbox List file. Please check the file, ensure you entered the correct path and name, and try again.\n"));
				DebugPrint(DBGGeneric,_T("\n"));
				DisplayUsage();
				goto Exit;
			}
		}
		else if (ProgOpts.lpszMailboxDN)
		{
			g_bOtherMBXMode = true;
			LPSERVICEADMIN	pSvcAdmin = NULL;
			LPPROFSECT		pGlobalProfSect = NULL;
			LPSPropValue	lpServerName	= NULL;
			WCHAR			*szUnicodeMsgStoreDN = NULL;

			g_MBXLegDN = (CString)lpszMailboxDN;
			g_MBXDisplayName = (CString)lpszDisplayName;

			// Do this single mailbox
			WC_H(OpenMessageStoreGUID(lpMAPISession, pbExchangeProviderPrimaryUserGuid, &lpPrimaryMDB));
			if (lpPrimaryMDB)
				WC_H(lpMAPISession->AdminServices(0,&pSvcAdmin));
			if (pSvcAdmin)
				WC_H(pSvcAdmin->OpenProfileSection((LPMAPIUID)pbGlobalProfileSectionGuid, NULL, 0,	&pGlobalProfSect));
			if (pGlobalProfSect)
					WC_H(HrGetOneProp(pGlobalProfSect,PR_PROFILE_HOME_SERVER,&lpServerName));
			if (lpPrimaryMDB && lpServerName && CheckStringProp(lpServerName,PT_STRING8))
			{				
				AnsiToUnicode(lpServerName->Value.lpszA, &szUnicodeMsgStoreDN);

				OpenOtherUsersMailbox(lpMAPISession,
									lpPrimaryMDB,
									szUnicodeMsgStoreDN,
									lpszMailboxDN,
									false, // don't need admin priv - don't assert it
									&lpMDB);

				if (lpMDB)
				{
					WC_H(ProcessMailbox(lpszMailboxDN, lpMDB, lpMAPISession));
				}
				else
				{
					DebugPrint(DBGError,_T("Could not connect. Ensure you are connecting Outlook via MAPI to an Exchange server.\n"));
					WriteLogFile(_T("ERROR: Could not connect. Ensure you are connecting Outlook via MAPI to an Exchange server."), 0);
				}
			}	
			
			MAPIFreeBuffer(lpServerName);
			if (pGlobalProfSect) pGlobalProfSect->Release();
			if (pSvcAdmin) pSvcAdmin->Release();
		}
		else
		{
			if(lpMDB)
			{
				// Just do primary Exchange mailbox in the profile
				WC_H(ProcessMailbox(lpMailboxName->Value.LPSZ, lpMDB, lpMAPISession));
			}
			else
			{
				DebugPrint(DBGError,_T("Could not connect. Ensure you are connecting Outlook via MAPI to an Exchange server.\n"));
				WriteLogFile(_T("ERROR: Could not connect. Ensure you are connecting Outlook via MAPI to an Exchange server."), 0);
			}
		}

		MAPIFreeBuffer(lpMailboxName);
		MAPIFreeBuffer(lpDisplayName);
		if (lpIdentity) 
		{
			lpIdentity->Release();
		}
		MAPIFreeBuffer(lpEntryID);
		if (lpPrimaryMDB) 
		{
			lpPrimaryMDB->Release();
		}
		if (lpMDB) 
		{
			lpMDB->Release();
		}
		lpMAPISession->Logoff(0, MAPI_LOGOFF_UI, 0);
		lpMAPISession->Release();
	}
	
	DebugPrint(DBGVerbose,_T("\n"));
	DebugPrint(DBGTime,_T("End process\n"));
	DebugPrint(DBGVerbose,_T("\n"));
	
	// remove temporary text files now that we're done
	RemoveTempFiles();

	if (g_bOutput) // report the correct directory below if a different one was specified
	{
		g_szAppPath = g_szOutputPath;
	}

	// if the File option was specified there could be many logs
	if (ProgOpts.lpszFileName)
	{
		//remove extra CalCheck logs if it all goes well - otherwise there may be an error to look at in the log
		if(g_bClearLogs)
			RemoveLogFiles();
		
		DebugPrint(DBGGeneric,_T("\n"));
		DebugPrint(DBGGeneric,_T("Please check the resulting files in " + g_szAppPath + " for more information.\n"));
		if(g_bFolder) DebugPrint(DBGGeneric,_T("Flagged error items were placed in the CalCheck folder in Outlook for each user.\n"));
	}
	else
	{
		DebugPrint(DBGGeneric,_T("\n"));
		DebugPrint(DBGGeneric,_T("Please check the resulting files in " + g_szAppPath + " for more information.\n"));
		if(g_bFolder) DebugPrint(DBGGeneric,_T("Flagged error items were placed in the CalCheck folder in Outlook.\n"));
	}

Exit:
		
	DebugPrint(DBGGeneric,_T("\n==========================================\n"));
	DebugPrint(DBGHelp,_T(   "\nPrivacy Note:\n"));
	DebugPrint(DBGHelp,_T("The data files produced by CalCheck can contain PII such as e-mail addresses.\n")); 
	DebugPrint(DBGHelp,_T("Please delete these files from your system after analysis and/or supplying them to Microsoft for analysis.\n"));
	DebugPrint(DBGHelp,_T("For more information on Microsoft's privacy standards and practices, please go to http://privacy.microsoft.com/en-us/default.mspx. \n"));
	
		
	MAPIUninitialize(); // If this hangs, look for the leaking MAPI objects...
	
	return;
}

// Loop through the mailboxes listed in file passed in by user
HRESULT LoopMailboxes(LPMAPISESSION lpMAPISession)
{
	DebugPrint(DBGVerbose,_T("\n"));
	DebugPrint(DBGTime,_T("Begin LoopMailboxes.\n"));
	DebugPrint(DBGVerbose,_T("\n"));
	
	LPSERVICEADMIN	pSvcAdmin = NULL;
	LPPROFSECT		pGlobalProfSect = NULL;
	LPSPropValue	lpServerName	= NULL;
	HRESULT			hRes = S_OK;
	CString			strHR = _T("");	
	CString			strProbItemsSvr, strErrorsSvr, strWarningsSvr;
	CStringList		rgstrProbMBXs;
	POSITION		posProb;
	POSITION		posList;
	LPMDB			lpPrimaryMDB	= NULL;
	LPMDB			lpUserMDB		= NULL;
	LPTSTR			lpszServerName  = _T("");
	LPTSTR			lpszMailboxDN   = _T("");
	LPTSTR			lpszDisplayName = _T("");
	WCHAR			*szUnicodeMsgStoreName = NULL;
	SYSTEMTIME		stTime;
	CString			strLocalTime;

	hRes = CreateServerLogFile();
	if (!hRes == S_OK)
	{			
		DebugPrint(DBGError,_T("Failed creating required output file CalCheckMaster.log. Please close any opened files and try again. Error: 0x%08X\n"), hRes);
		goto Exit;
	}

	DebugPrint(DBGGeneric,_T("Running multi-mailbox mode. \n"));
	WriteSvrLogFile(_T("Running multi-mailbox mode.\r\n"));

	// Need to get the User's Exchange Server Name (the user who is running this)
	WC_H(OpenMessageStoreGUID(lpMAPISession, pbExchangeProviderPrimaryUserGuid, &lpPrimaryMDB));
	if (lpPrimaryMDB)
		WC_H(lpMAPISession->AdminServices(0,&pSvcAdmin));
	if (pSvcAdmin)
		WC_H(pSvcAdmin->OpenProfileSection((LPMAPIUID)pbGlobalProfileSectionGuid, NULL, 0,	&pGlobalProfSect));
	if (pGlobalProfSect)
		WC_H(HrGetOneProp(pGlobalProfSect,PR_PROFILE_HOME_SERVER,&lpServerName));
	if (lpPrimaryMDB && lpServerName && CheckStringProp(lpServerName,PT_STRING8))
	{
		AnsiToUnicode(lpServerName->Value.lpszA,&szUnicodeMsgStoreName);
		lpszServerName = szUnicodeMsgStoreName;
	}

	// Need to get each mailbox from the in-memory CStringList and try to process it.
	if (!MBXInfoList.IsEmpty())
	{
		CString strReadLine = _T("");
		CString strDispName = _T("");
		CString strMailboxDN = _T("");
		int iStrLen = 0;
		int iColon = 0;
		int iCopy = 0;
		int iElements = 0;
		int iChar = 0;
		int iMBXInfoNeeded = 2;

		for (posList = MBXInfoList.GetHeadPosition(); posList != NULL; )
		{
			strReadLine = MBXInfoList.GetNext(posList);

			if (strReadLine.Find(_T("name")) != -1)
			{
				iChar = strReadLine.Find('n', 0);
				if (iChar == 0)
				{
					iStrLen = strReadLine.GetLength();
					iColon = strReadLine.Find(':');
					iCopy = ((iStrLen - iColon)-2);
					if (lpszDisplayName == _T(""))
					{
						strDispName = strReadLine.Right(iCopy);
						lpszDisplayName = strDispName.GetBuffer(0);
						iElements++;
					}
					else
					{
						DebugPrint(DBGGeneric,_T("Mailbox List File Format Error.\n"));
						return 1;
					}
				}
			}

			if (strReadLine.Find(_T("legacyexchangedn")) != -1)
			{
				iChar = strReadLine.Find('l', 0);
				if (iChar == 0)
				{
					iStrLen = strReadLine.GetLength();
					iColon = strReadLine.Find(':');
					iCopy = ((iStrLen - iColon)-2);
					if (lpszMailboxDN == _T(""))
					{
						strMailboxDN = strReadLine.Right(iCopy);
						lpszMailboxDN = strMailboxDN.GetBuffer(0);
						iElements++;
					}
					else
					{
						DebugPrint(DBGGeneric,_T("Mailbox List File Format Error.\n"));
						return 1;
					}
				}
			}
			if (iElements == iMBXInfoNeeded) // we need two elements populated above before trying to process
			{
				WriteLogFile(_T("Opening mailbox: " + (CString)lpszDisplayName), 1);
				WriteLogFile(_T("DN: " + (CString)lpszMailboxDN + "/r/n"), 1);
				
				// We should have the display name, legDN, and server name now, so process this mailbox
				WC_H(OpenOtherUsersMailbox(lpMAPISession,
											lpPrimaryMDB,
											lpszServerName,
											lpszMailboxDN,
											false, // don't need admin priv - don't assert it
											&lpUserMDB));

				g_MBXDisplayName = (CString)lpszDisplayName;
				g_MBXLegDN = (CString)lpszMailboxDN;

				if (lpUserMDB)
				{
					WC_H(ProcessMailbox(lpszMailboxDN,lpUserMDB,lpMAPISession));
					ULONG lpulFlags = LOGOFF_NO_WAIT;
					WC_H(lpUserMDB->StoreLogoff(&lpulFlags));
					lpUserMDB->Release();
					lpUserMDB = NULL;

					if(g_bFoundProb)
					{
						//if probs were found, then set the global back to false for the next mailbox
						//and add this one to the list
						g_bFoundProb = false;
						rgstrProbMBXs.AddHead((CString)lpszMailboxDN);
					}
				}
				else
				{
					CString strUser;

					DebugPrint(DBGGeneric,_T("\n"));
					DebugPrint(DBGError,_T("Opening a mailbox returned error 0x%08X\n"),hRes);
					DebugPrint(DBGError,_T("Display Name: %s\n"), lpszDisplayName);
					DebugPrint(DBGError,_T("LegacyExchangeDN: %s\n"), lpszMailboxDN);
					DebugPrint(DBGError,_T("Check the mailbox name and DN, and that the store is mounted and you have permissions.\n"));
					DebugPrint(DBGError,_T("Continuing...\n"));
					DebugPrint(DBGError,_T("==============\n"));

					strHR.Format(_T("%08X"), hRes);
					WriteLogFile(_T("Failed to open this mailbox. Error: ") + strHR, 1);
					WriteLogFile(_T("Display Name: ") + (CString)lpszDisplayName, 1);
					WriteLogFile(_T("LegacyExchangeDN: ") + (CString)lpszMailboxDN, 1);
					WriteLogFile(_T("Check the mailbox name and DN, and that the store is mounted and you have permissions.\r\n"), 1);

					// If an error opening a mailbox, we won't process, but we still need to have a separate log file
					strUser += _T("_");
					strUser += GetDNUser(lpszMailboxDN);

					AppendFileNames((LPCTSTR)strUser);

					CreateLogFiles();
				}

				iElements = 0;
				lpszDisplayName = _T("");
				lpszMailboxDN = _T("");
			}

		} // for Entire list of mailboxes to go through
	}
	else
	{
		// something bad happened and there's no mailbox info list to go through.
		WriteLogFile(_T("No mailbox information list - check to ensure the input list file is correct and try again."), 1);
	}
	
	GetLocalTime(&stTime);
	strLocalTime.Format(_T("%02d/%02d/%d %02d:%02d:%02d%s"), 
							stTime.wMonth, 
							stTime.wDay, 
							stTime.wYear, 
							(stTime.wHour <= 12)?stTime.wHour:stTime.wHour-12,  // adjust to non-military time
							stTime.wMinute, 
							stTime.wSecond,
							(stTime.wHour <= 12)?_T("AM"):_T("PM"));

	WriteSvrLogFile(_T("Finish: ") + strLocalTime + ("\r\n"));

	strProbItemsSvr.Format(_T("%d"), g_ulProbItemsSvr);
	strErrorsSvr.Format(_T("%d"), g_ulErrorsSvr);
	strWarningsSvr.Format(_T("%d"), g_ulWarningsSvr);

	// Do some overall reporting...
	WriteSvrLogFile(_T("CalCheck Multi Mailbox Mode Results:"));
	WriteSvrLogFile(_T("Total Problem Items: ") + strProbItemsSvr);
	WriteSvrLogFile(_T("Total Errors found: ") + strErrorsSvr);
	WriteSvrLogFile(_T("Total Warnings found: ") + strWarningsSvr + (" \n"));

	if(!rgstrProbMBXs.IsEmpty())
	{
		WriteSvrLogFile(_T("Mailboxes reporting problems:"));
		for(posProb = rgstrProbMBXs.GetHeadPosition(); posProb != NULL; )
		{
			WriteSvrLogFile(rgstrProbMBXs.GetNext(posProb));
		}

	}

	lpPrimaryMDB->Release();

	DebugPrint(DBGVerbose,_T("\n"));
	DebugPrint(DBGTime,_T("End LoopServer\n"));
	DebugPrint(DBGVerbose,_T("\n"));

Exit:

	return hRes;
}

HRESULT ProcessMailbox(LPTSTR szDN, LPMDB lpMDB, LPMAPISESSION lpMAPISession)
{
	HRESULT				hRes = S_OK;
	CString				strUser = _T("");
	CString				strProbItems = _T("");
	CString				strErrors = _T("");
	CString				strWarnings = _T("");
	CString				strTotalItems = _T("");
	CString				strRecurItems = _T("");
	CString				strHRes = _T("");
	LPMAPIFOLDER		lpCalendarFolder = NULL;
	SYSTEMTIME			stTime;
	CString				strLocalTime;

	DebugPrint(DBGVerbose,_T("\n"));
	DebugPrint(DBGTime,_T("Begin ProcessMailbox\n"));
	DebugPrint(DBGVerbose,_T("\n"));

	if (g_MBXDisplayName == _T("")) // If regular mode, user is the mailbox
	{
		g_MBXDisplayName = g_UserDisplayName;
	}

	DebugPrint(DBGVerbose,_T("Opening mailbox: %s\n"), g_MBXDisplayName);
	DebugPrint(DBGVerbose,_T("%s\n"),szDN);

	if(!g_bMultiMode) // Server Mode writes this elsewhere...
	{
		WriteLogFile(_T("Opening mailbox: " + g_MBXDisplayName), 0);
		WriteLogFile(_T("DN: " + (CString)szDN + "\r\n"), 0);
	}

	// We want to get all the proxy addresses for this user - and this is the best place to do it...
	hRes = GetProxyAddrs(szDN, lpMAPISession);

	// Open the Calendar Folder
	hRes = OpenDefaultFolder(DEFAULT_CALENDAR, lpMDB, NULL, &lpCalendarFolder);
	if (FAILED(hRes))
	{
		strHRes.Format(_T("%08X"), hRes);
		DebugPrint(DBGGeneric,_T("\n"));
		DebugPrint(DBGGeneric,_T("CalCheck failed to locate default Calendar for %s - skipping.\n"),szDN);
		DebugPrint(DBGGeneric,_T("Error: 0x%08X\n"), hRes); 
		WriteLogFile(_T("Error " + strHRes + " opening default Calendar."), 1);
		WriteLogFile(_T("This could be expected for a system mailbox, otherwise check things like permissions and entry ids."), 1);
	}

	if (SUCCEEDED(hRes) && lpCalendarFolder)
	{
		// BIG ProcessCalendar Function since we seem to have it and it looks right
		hRes = ProcessCalendar(szDN, lpMDB, lpCalendarFolder);

		if (FAILED(hRes)) 
		{
			strHRes.Format(_T("%08X"), hRes);
			DebugPrint(DBGGeneric,_T("CalCheck failed to process calendar for %s\n"),szDN);
			DebugPrint(DBGGeneric,_T("Error: 0x%08X\n"), hRes);
			WriteLogFile(_T("Error " + strHRes + " processing calendar for " + (CString)szDN), 1);
		}
	}

	// clear the list of proxy addresses - for server mode before opening another mailbox
	if(!AddrList.IsEmpty())
	{
		AddrList.RemoveAll();
	}

	//Clean up and log the end of the process
	if (lpCalendarFolder) 
	{
		lpCalendarFolder->Release();
	}

	if(g_ulProbItems > 0)
	{
		strProbItems.Format(_T("%d"), g_ulProbItems);
		strErrors.Format(_T("%d"), g_ulErrors);
		strWarnings.Format(_T("%d"), g_ulWarnings);
		strTotalItems.Format(_T("%d"), g_ulItems);
		strRecurItems.Format(_T("%d"), g_ulRecurItems);
		
		WriteLogFile(_T("CalCheck Results:  (Errors and Warnings may add up to more than the number of problem items when there are multiple problems found in one item)"), 0);
		WriteLogFile(_T("================="), 0);
		WriteLogFile(_T("Problem Items: ") + strProbItems, 0);
		WriteLogFile(_T("Errors found: ") + strErrors, 0);
		WriteLogFile(_T("Warnings found: ") + strWarnings + (" \n"), 0);
		WriteLogFile(_T("Total Calendar Items: ") + strTotalItems, 0);
		WriteLogFile(_T("Number of Recurring Items: ") + strRecurItems + (" \n"), 0);
	}
	else
	{
		WriteLogFile(_T("No known errors or problems were found.\r\n"), 0);
	}

	GetLocalTime(&stTime);
	strLocalTime.Format(_T("%02d/%02d/%d %02d:%02d:%02d%s"), 
							stTime.wMonth, 
							stTime.wDay, 
							stTime.wYear, 
							(stTime.wHour <= 12)?stTime.wHour:stTime.wHour-12,  // adjust to non-military time
							stTime.wMinute, 
							stTime.wSecond,
							(stTime.wHour <= 12)?_T("AM"):_T("PM"));

	WriteLogFile(_T("Finish: ") + strLocalTime + ("\r\n"), 0);

	WriteLogFile(_T("Finished processing mailbox " + (CString)szDN), 0);

	if(g_bMultiMode)
	{
		if(!g_bFoundProb) // if no problems were found with any items, then let the user know that
		{
			WriteLogFile(_T("No known errors or problems were found.\r\n"), 0);
		}

		WriteSvrLogFile(_T("Finished processing mailbox " + (CString)szDN));

		// continue server mode counts - reset happens a little further down in the code...
		g_ulProbItemsSvr += g_ulProbItems;
		g_ulErrorsSvr += g_ulErrors;
		g_ulWarningsSvr += g_ulWarnings;
	}

	if(g_bReport)
	{
		// if not server mode OR if server mode and probs were detected then do it
		if(!g_bMultiMode || (g_bMultiMode && g_ulProbItems > 0))
		{
			// Create a message and place it in the Inbox with log as attachment
			hRes = CreateReportMsg(lpMDB);
		}
	}

	// set item counts back to 0 before going to next Calendar
	g_ulProbItems = 0;
	g_ulErrors = 0;
	g_ulWarnings = 0;
	g_ulItems = 0;
	g_ulRecurItems = 0;


	// if doing multiple mailboxes, then make separate files for each mailbox
	if(ProgOpts.lpszFileName)
	{
		
		strUser += _T("_");
		strUser += GetDNUser(szDN);

		AppendFileNames((LPCTSTR)strUser);

		CreateLogFiles();
	}

	DebugPrint(DBGVerbose,_T("\n"));
	DebugPrint(DBGTime,_T("End ProcessMailbox\n"));
	DebugPrint(DBGVerbose,_T("\n"));

	return hRes;
}

// Get Proxy Addresses for this user
HRESULT GetProxyAddrs(LPTSTR szDN, LPMAPISESSION lpMAPISession)
{
	HRESULT				hRes = S_OK;
	LPADRBOOK			lpAB = NULL;
	LPADRLIST			lpAddrList = NULL;
	CString				strAddress;
	ULONG				iProp = 0;
	ULONG				i = 0;
	ULONG				cValues = 0;
	CString				strHR = _T("");

	if(g_MBXLegDN.IsEmpty())
	{
		g_MBXLegDN = szDN;
	}
	
	AddrList.AddHead(g_MBXLegDN); // need to put the DN in the list - so doing that now...

	hRes = lpMAPISession->OpenAddressBook(NULL,NULL,NULL, &lpAB);

	if(SUCCEEDED(hRes) && lpAB)
	{
		hRes = MAPIAllocateBuffer(CbNewSRowSet(1),(LPVOID*) &lpAddrList);

		if(SUCCEEDED(hRes) && lpAddrList)
		{
			ZeroMemory ( lpAddrList, CbNewSRowSet(1));
			hRes = MAPIAllocateBuffer(7 * sizeof(SPropValue), (LPVOID FAR *)&(lpAddrList->aEntries[0].rgPropVals));

			if(SUCCEEDED(hRes))
			{
				ZeroMemory ( lpAddrList -> aEntries[0].rgPropVals, 7 * sizeof(SPropValue));
				
				lpAddrList->cEntries = 1;
				lpAddrList->aEntries[0].cValues = 2L;
				
				lpAddrList->aEntries[0].rgPropVals[0].ulPropTag = PR_DISPLAY_NAME;
				lpAddrList->aEntries[0].rgPropVals[0].Value.lpszW = (LPWSTR)(LPCTSTR)g_MBXDisplayName;
				lpAddrList->aEntries[0].rgPropVals[1].ulPropTag = PR_ADDRTYPE_W;
				lpAddrList->aEntries[0].rgPropVals[1].Value.lpszW = _T("EX");

				hRes = lpAB->ResolveName(NULL,MAPI_UNICODE,NULL,lpAddrList);

				if(SUCCEEDED(hRes))
				{
					for(i = 0; i < 7; ++i)
					{
						ULONG				cbEntryID = 0;
						LPENTRYID			lpEntryID = NULL;
						ULONG				ulObjType = NULL;
						LPMAPIPROP			lpAddrs = NULL;
						LPSPropValue		lpAddrsProp = NULL;

						if(lpAddrList->aEntries[0].rgPropVals[i].ulPropTag == PR_ENTRYID)
						{
							cbEntryID = lpAddrList->aEntries[0].rgPropVals[i].Value.bin.cb;
							lpEntryID = (LPENTRYID)lpAddrList->aEntries[0].rgPropVals[i].Value.bin.lpb;
								
							if(g_bMultiMode || g_bOtherMBXMode) // if running Calcheck in server mode or other MBX mode - check the AB in Online Mode
							{
								hRes = S_OK;
									if (lpAddrs) (lpAddrs)->Release();
									lpAddrs = NULL;
									hRes = lpAB->OpenEntry(cbEntryID,
														   lpEntryID,
														   NULL, 
														   MAPI_BEST_ACCESS | MAPI_NO_CACHE,
														   &ulObjType,
														   (LPUNKNOWN*)&lpAddrs);
							}
							else
							{
								hRes = lpAB->OpenEntry(cbEntryID,
													   lpEntryID,
													   NULL,
													   MAPI_BEST_ACCESS | MAPI_CACHE_ONLY,
													   &ulObjType,
													   (LPUNKNOWN*)&lpAddrs);

								if(FAILED(hRes)) // if Cached Mode fails then try Online mode...
								{
									hRes = S_OK;
									if (lpAddrs) (lpAddrs)->Release();
									lpAddrs = NULL;
									hRes = lpAB->OpenEntry(cbEntryID,
														   lpEntryID,
														   NULL, 
														   MAPI_BEST_ACCESS | MAPI_NO_CACHE,
														   &ulObjType,
														   (LPUNKNOWN*)&lpAddrs);
								}
							}

							if(SUCCEEDED(hRes) && lpAddrs)
							{
								hRes = HrGetOneProp(lpAddrs, PR_EMS_AB_PROXY_ADDRESSES, &lpAddrsProp);

								if(SUCCEEDED(hRes) && lpAddrsProp)
								{
									WriteLogFile(_T("Proxy Addresses for ") + g_MBXDisplayName + ":", 0);
									
									cValues = lpAddrsProp->Value.MVi.cValues;

									for(iProp = 0; iProp < cValues; iProp++)
									{
										strAddress = lpAddrsProp->Value.MVszW.lppszW[iProp];
										WriteLogFile(strAddress, 0);
										AddrList.AddHead(strAddress);
									}								
								}
								else
								{
									strHR.Format(_T("%08X"), hRes);
									WriteLogFile(_T("Error " + strHR + " retrieving proxy addresses for ") + g_MBXDisplayName, 0);
									WriteLogFile(_T("Ensure connectivity and permissions to access this address book entry."), 0);
								}
							}
							else
							{
								strHR.Format(_T("%08X"), hRes);
								WriteLogFile(_T("Error " + strHR + " opening the address book entry for ") + g_MBXDisplayName, 0);	
								WriteLogFile(_T("Ensure connectivity and permissions to access this address book entry."), 0);
							}
						}

						MAPIFreeBuffer(lpAddrsProp);
						if(lpAddrs)
						{
							(lpAddrs)->Release();
						}

					} //for loop
				}
				else if (hRes == MAPI_E_AMBIGUOUS_RECIP)
				{
					strHR.Format(_T("%08X"), hRes);
					WriteLogFile(_T("Error " + strHR + "resolving the name in the address book for ") + g_MBXDisplayName, 0);
					WriteLogFile(_T("There are multiple users in the address book that match, so could not get proxy addresses."), 0);
				}
				else
				{
					strHR.Format(_T("%08X"), hRes);
					WriteLogFile(_T("Error " + strHR + "resolving the name in the address book for ") + g_MBXDisplayName, 0);
					WriteLogFile(_T("If this is a system or hidden mailbox, this is expected, otherwise ensure connectivity and permissions to access the address book."), 0);
				}
			}
		}
	}

	if (FAILED(hRes))
	{
		WriteLogFile(_T("Proxy addresses are used for the Attendee becomes Organizer test."), 0);
		WriteLogFile(_T("Failure to get proxy addresses only affects this one test, so check results carefully."), 0);
	}

	if (lpAB)
	{
		lpAB->Release();
	}
	if (lpAddrList)
	{
		MAPIFreeBuffer(lpAddrList);
	}

	WriteLogFile(_T("\r\n"), 0);

	return hRes;
}

// Read in Text File containing Names and DNs from 
BOOL ReadMBXListFile(LPSTR	lpszFileName)
{
	/* Based on output from Get-Mailbox command where output looks like:
	Name             : test01
	LegacyExchangeDN : /o=Reprohouse/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=test01
	*/
	CStdioFile MBXFile;
	CString strMBXFile = _T("");
	CFileStatus status;
	CFileException e;
	CString strReadLine = _T("");
	CString strWriteLine = _T("");
	BOOL bReturn = 1;

	strMBXFile = lpszFileName;

	if (strMBXFile.Find(':') == -1) // if no path was given, pre-pend the application path
	{
		CString strTemp = g_szAppPath;
		strTemp += _T("\\") + strMBXFile;

		strMBXFile = strTemp;
	}

	if(MBXFile.GetStatus(strMBXFile, status))
	{
		//open it
		if(MBXFile.Open((LPCTSTR)strMBXFile, CFile::modeRead | CFile::typeText))
		{
			ULONG ulLength = MBXFile.GetLength();
			ULONG ulLen2 = ulLength / 2;
			CString strBuffer;					
					 

			MBXFile.Read(strBuffer.GetBuffer(ulLength), ulLength);
			LPCWSTR lpszTemp1 = strBuffer.GetBuffer(0);
			LPCWSTR lpszTemp2 = lpszTemp1;
			LPBYTE bEdit = (LPBYTE)lpszTemp2;

			BYTE b1 = (BYTE)*lpszTemp1;
			if(b1 == 0xff) // This is a UNICODE File if it has this in the first byte
			{
				// first go to the end of the text and NULL terminate it
				lpszTemp2 += ulLen2;
				b1 = (BYTE)*lpszTemp2;
				if(b1 != 0x00)
				{
					bEdit = (LPBYTE)lpszTemp2;
					*bEdit = 0x00;
					bEdit += 1;
					*bEdit = 0x00;
				}
				// now set the beginning of the text to the actual beginning of the text
				for (int i = 0; i < ulLength; i++)
				{
					b1 = (BYTE)*lpszTemp1;

					if(b1 == 0xff || b1 == 0xfe || b1 == 0xef || b1 == 0x00 || b1 == 0x0d || b1 == 0x0a)
					{
						lpszTemp1 += 1;
						b1 = (BYTE)*lpszTemp1;
					}
					else // should be at the actual unicode text now - set that as the CString
					{
						//strBuffer.Empty();
						//strBuffer.SetString(lpszTemp1, ulLength);

						strBuffer.Empty();
						strBuffer = lpszTemp1;

						lpszTemp1 = _T("");
						lpszTemp2 = _T("");

						break;
					}
				}

				int iPos1 = 0;
				int iPos2 = 0;						
						
				strBuffer.Replace(_T("\r\n\r\n"), _T("\r\n")); // remove double crlf's... that messes things up below...

				ULONG ulBuf = strBuffer.GetLength();

				while (ulBuf > 0)
				{
					iPos1 = strBuffer.Find(_T("\r\n"), iPos1); 

					if(iPos1 != -1)
					{
						strReadLine = strBuffer.Left(iPos1);
						strBuffer = strBuffer.Right(ulBuf-iPos1-2);
						ulBuf = strBuffer.GetLength();
						strReadLine.MakeLower();
						MBXInfoList.AddHead(strReadLine);
						iPos1 = 0;
					}
					else // no crlf should mean we're at the end of the buffer
					{
						strReadLine = strBuffer;
						strReadLine.MakeLower();
						MBXInfoList.AddHead(strReadLine);
						ulBuf = 0;
					}				
				}
			}
			else
			{
				// it's an ANSI file and we'll process that
				MBXFile.SeekToBegin();
				while(MBXFile.ReadString(strReadLine))
				{
					if (strReadLine.Find(':') != -1)
					{
						strReadLine.MakeLower();
						MBXInfoList.AddHead(strReadLine);
					}
				}
			}
		}
		else
		{
			DebugPrint(DBGGeneric,_T("Could not open file: \"%s\" Check access to the file and directory.\n"), lpszFileName);
			return false;
		}
	}
	else
	{
		DebugPrint(DBGGeneric,_T("Could not find or open the file: \"%s\" \n"), lpszFileName);
		return false;
	}

	return true;
}

CString rgstrTests[51] = {"organizeraddress",
							"organizername",
							"senderaddress",
							"sendername",
							"nosubject",			
							"messageclass",			
							"conflictitems",		
							"recuritemlimit",		
							"itemsize10",
							"itemsize25",
							"itemsize50",
							"delivertime",
							"attachcount",			
							"recurringprop",
							"starttimemin",
							"starttimex",			
							"starttime1995",		
							"starttime2025",		
							"starttimemax",
							"endtimemin",
							"endtimex",			
							"endtime1995",			
							"endtime2025",			
							"endtimemax",
							"recurstartmin",			
							"recurstart1995",		
							"recurstart2025",		
							"recurstartmax",
							"recurendmin",			
							"recurend1995",		
							"recurend2025",		
							"recurendmax",			
							"exceptionbounds",		
							"exceptiondata",		
							"duplicates",			
							"attendtoorganizer",	
							"dupglobalobjectids",	
							"noglobalobjectids",
							"rtaddresstype",
							"rtaddress",
							"rtdisplayname",
							"rtduplicates",
							"rtorganizeraddress",
							"rtorganizerisorganizer",
							"timezonedefrecur",
							"timezonedefstart",
							"propdefstream",
							"holidayitems",
							"birthdayitems",
							"pastitems",
							"warningiserror"
						  };

int cTests = 51;  // Be sure to update this if you add any strings to the array above


// Read in CalCheck.cfg to an in-memory string array. Return true if successful.
BOOL ReadConfigFile()
{
	CStdioFile CFGFile;
	CString strCFGFile = _T("");
	CFileStatus status;
	CFileException e;
	CString strReadLine = _T("");
	int iComment = 0; // where "'" is located
	int iTestFound = 0;	
	BOOL bReturn = 1;

	// OffCAT might be putting the CalCheck files in a Read-only location, 
	// so if using the Input switch then the CFG file can be modified for additional use by OffCAT
	if(g_bInput)
		strCFGFile = g_szInputPath;
	else
		strCFGFile = g_szAppPath;
	strCFGFile += _T("\\CalCheck.cfg");

	// set all the tests to true - they will run unless set to false or 0 in the config file - except the WarningIsError test.
	CalCheckTests.bAttachCount = true;
	CalCheckTests.bAttendToOrganizer = true;
	CalCheckTests.bBirthdayItems = true;
	CalCheckTests.bConflictItems = true;
	CalCheckTests.bDupGlobalObjectIDs = true;
	CalCheckTests.bDuplicates = true;
	CalCheckTests.bEndTime1995 = true;
	CalCheckTests.bEndTime2025 = true;
	CalCheckTests.bEndTimeMax = true;
	CalCheckTests.bEndTimeMin = true;
	CalCheckTests.bEndTimeX = true;
	CalCheckTests.bExceptionBounds = true;
	CalCheckTests.bExceptionData = true;
	CalCheckTests.bHolidayItems = true; 
	CalCheckTests.bItemSize10 = true;
	CalCheckTests.bItemSize25 = true;
	CalCheckTests.bItemSize50 = true;
	CalCheckTests.bDeliverTime = true;
	CalCheckTests.bMessageClass = true;
	CalCheckTests.bNoGlobalObjectIDs = true;
	CalCheckTests.bNoSubject = true;
	CalCheckTests.bOrganizerAddress = true;
	CalCheckTests.bOrganizerName = true;
	CalCheckTests.bPastItems = true;
	CalCheckTests.bRecurEnd1995 = true;
	CalCheckTests.bRecurEnd2025 = true;
	CalCheckTests.bRecurEndMax = true;
	CalCheckTests.bRecurEndMin = true;
	CalCheckTests.bRecurItemLimit = true;
	CalCheckTests.bRecurringProp = true;
	CalCheckTests.bRecurStart1995 = true;
	CalCheckTests.bRecurStart2025 = true;
	CalCheckTests.bRecurStartMax = true;
	CalCheckTests.bRecurStartMin = true;
	CalCheckTests.bRTAddress = true;
	CalCheckTests.bRTAddressType = true;
	CalCheckTests.bRTDisplayName = true;
	CalCheckTests.bRTDuplicates = true;
	CalCheckTests.bRTOrganizerAddress = true;
	CalCheckTests.bRTOrganizerIsOrganizer = true;
	CalCheckTests.bSenderAddress = true;
	CalCheckTests.bSenderName = true;
	CalCheckTests.bStartTime1995 = true;
	CalCheckTests.bStartTime2025 = true;
	CalCheckTests.bStartTimeMax = true;
	CalCheckTests.bStartTimeMin = true;
	CalCheckTests.bStartTimeX = true;
	CalCheckTests.bTZDefRecur = true;
	CalCheckTests.bTZDefStart = true;
	CalCheckTests.bPropDefStream = true;
	CalCheckTests.bWarningIsError = false; // false by default

	// if the file is there...
	if(CFGFile.GetStatus(strCFGFile, status))
	{
		//open it
		if(CFGFile.Open((LPCTSTR)strCFGFile, CFile::modeRead | CFile::typeText))
		{
			// get the settings out of the file
			CFGFile.SeekToBegin();
			while(CFGFile.ReadString(strReadLine))
			{
				if (strReadLine.IsEmpty()) // skip empty lines
					continue;

				iComment = strReadLine.Find('\'', 0);
				if (iComment == 0) // skip lines marked with "'" at the beginning
					continue;

				strReadLine.MakeLower();
				for (int i = 0; i < cTests; i++)
				{
					if(strReadLine.Find(rgstrTests[i]) != -1)
					{
						SelectTest(strReadLine);	
					}
				}
			}
		}
		else
		{
			DebugPrint(DBGGeneric,_T("Could not open CalCheck.cfg. Check access to the directory where CalCheck.exe is located: \"%s\"\n"), g_szAppPath);
			bReturn = 0; 
		}
	}
	else
	{
		// Missing file means run all tests	
		DebugPrint(DBGGeneric,_T("Missing CalCheck.cfg - which is used to configure tests to run. Make sure the file is located in the same directory with CalCheck.exe: \"%s\"\n"), g_szAppPath);
		DebugPrint(DBGGeneric,_T("Continuing with CalCheck running all tests.\n"));
		WriteLogFile(_T("Missing CalCheck.cfg - which is used to configure tests to run."), 1);
		WriteLogFile(_T("Make sure the file is located in the same directory with CalCheck.exe:  " + g_szAppPath), 1);
		WriteLogFile(_T("Continuing with CalCheck running all tests.\n"), 1);
	}

	return bReturn;
}

// If new tests are added - they need to be added here as well in order to be able to turn them on/off
VOID SelectTest(CString strTest)
{
	// Find the test - then set it to false if that's how it's set in the config file
	if(strTest.Find(_T("organizeraddress")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bOrganizerAddress = false;
	}

	if(strTest.Find(_T("organizername")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bOrganizerName = false;
	}

	if(strTest.Find(_T("senderaddress")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bSenderAddress = false;		
	}

	if(strTest.Find(_T("sendername")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bSenderName = false;		
	}

	if(strTest.Find(_T("nosubject")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bNoSubject = false;		
	}

	if(strTest.Find(_T("messageclass")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bMessageClass = false;		
	}

	if(strTest.Find(_T("conflictitems")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bConflictItems = false;		
	}

	if(strTest.Find(_T("recuritemlimit")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bRecurItemLimit = false;		
	}

	if(strTest.Find(_T("itemsize10")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bItemSize10 = false;		
	}

	if(strTest.Find(_T("itemsize25")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bItemSize25 = false;		
	}

	if(strTest.Find(_T("itemsize50")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bItemSize50 = false;		
	}

	if (strTest.Find(_T("delivertime")) != -1)
	{
		if (strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bDeliverTime = false;
	}

	if(strTest.Find(_T("attachcount")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bAttachCount = false;		
	}

	if(strTest.Find(_T("recurringprop")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bRecurringProp = false;		
	}

	if(strTest.Find(_T("starttimemin")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bStartTimeMin = false;		
	}

	if(strTest.Find(_T("starttimex")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bStartTimeX = false;		
	}

	if(strTest.Find(_T("starttime1995")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bStartTime1995 = false;		
	}

	if(strTest.Find(_T("starttime2025")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bStartTime2025 = false;		
	}

	if(strTest.Find(_T("starttimemax")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bStartTimeMax = false;		
	}

	if(strTest.Find(_T("endtimemin")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bEndTimeMin = false;		
	}

	if(strTest.Find(_T("endtimex")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bEndTimeX = false;		
	}

	if(strTest.Find(_T("endtime1995")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bEndTime1995 = false;		
	}

	if(strTest.Find(_T("endtime2025")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bEndTime2025 = false;		
	}

	if(strTest.Find(_T("endtimemax")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bEndTimeMax = false;		
	}

	if(strTest.Find(_T("recurstartmin")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bRecurStartMin = false;		
	}

	if(strTest.Find(_T("recurstart1995")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bRecurStart1995 = false;		
	}

	if(strTest.Find(_T("recurstart2025")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bRecurStart2025 = false;		
	}

	if(strTest.Find(_T("recurstartmax")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bRecurStartMax = false;		
	}

	if(strTest.Find(_T("recurendmin")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bRecurEndMin = false;		
	}

	if(strTest.Find(_T("recurend1995")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bRecurEnd1995 = false;		
	}

	if(strTest.Find(_T("recurend2025")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bRecurEnd2025 = false;		
	}

	if(strTest.Find(_T("recurendmax")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bRecurEndMax = false;		
	}

	if(strTest.Find(_T("exceptionbounds")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bExceptionBounds = false;		
	}

	if(strTest.Find(_T("exceptiondata")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bExceptionData = false;		
	}

	if(strTest.Find(_T("duplicates")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bDuplicates = false;		
	}

	if(strTest.Find(_T("attendtoorganizer")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bAttendToOrganizer = false;		
	}

	if(strTest.Find(_T("dupglobalobjectids")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bDupGlobalObjectIDs = false;		
	}

	if(strTest.Find(_T("noglobalobjectids")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bNoGlobalObjectIDs = false;		
	}

	if(strTest.Find(_T("rtaddresstype")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bRTAddressType = false;		
	}

	if(strTest.Find(_T("rtaddress")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bRTAddress = false;		
	}

	if(strTest.Find(_T("rtdisplayname")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bRTDisplayName = false;		
	}

	if(strTest.Find(_T("rtduplicates")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bRTDuplicates = false;		
	}

	if(strTest.Find(_T("rtorganizeraddress")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bRTOrganizerAddress = false;		
	}

	if(strTest.Find(_T("rtorganizerisorganizer")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bRTOrganizerIsOrganizer = false;		
	}

	if(strTest.Find(_T("holidayitems")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bHolidayItems = false;		
	}

	if(strTest.Find(_T("birthdayitems")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bBirthdayItems = false;		
	}

	if(strTest.Find(_T("timezonedefrecur")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bTZDefRecur = false;		
	}

	if(strTest.Find(_T("timezonedefstart")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bTZDefStart = false;		
	}

	if(strTest.Find(_T("propdefstream")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bPropDefStream = false;		
	}

	if(strTest.Find(_T("pastitems")) != -1)
	{
		if(strTest.Find(_T("=false")) != -1 || strTest.Find(_T("=0")) != -1)
			CalCheckTests.bPastItems = false;		
	}

	if(strTest.Find(_T("warningiserror")) != -1)
	{
		if(strTest.Find(_T("=true")) != -1 || strTest.Find(_T("=1")) != -1)
			CalCheckTests.bWarningIsError = true;		
	}

}


HRESULT LoadMAPIVer(ULONG ulMAPIVer)
{
	HRESULT hRes = E_FAIL;
	HMODULE hMod = NULL;
	LPTSTR lpszTempPath = NULL;
	LPWSTR szPath = NULL;
    TCHAR pszaOutlookQualifiedComponents[][MAX_PATH] = {
        TEXT("{BC174BAD-2F53-4855-A1D5-0D575C19B1EA}"), // Outlook 2003
		TEXT("{24AAE126-0911-478F-A019-07B875EB9996}"), // Outlook 2007
		TEXT("{1E77DE88-BCAB-4C37-B9E5-073AF52DFD7A}"), // Outlook 2010
		TEXT("{E83B4360-C208-4325-9504-0D23003A74A5}"), // Outlook 2013
		TEXT("{5812C571-53F0-4467-BEFA-0A4F47A9437C}"), // Outlook 2016
    };
    int nOutlookQualifiedComponents = _countof(pszaOutlookQualifiedComponents);
	int iVer = 0;
    DWORD dwValueBuf = 0;
    UINT ret = 0;
	BOOL b64 = FALSE;

	if (ulMAPIVer == 16)
	{
		iVer = 4;
	}

	if (ulMAPIVer == 15)
	{
		iVer = 3;
	}

	if (ulMAPIVer == 14)
	{
		iVer = 2;
	}

	if (ulMAPIVer == 12)
	{
		iVer = 1;
	}

	if (ulMAPIVer == 11)
	{
		iVer = 0;
	}

	// check for 64-bit first...
	ret = MsiProvideQualifiedComponent(pszaOutlookQualifiedComponents[iVer],
									   _T("outlook.x64.exe"),
									   (DWORD) INSTALLMODE_DEFAULT,
									   NULL,
									   &dwValueBuf);
	
	if (ret != ERROR_SUCCESS)
	{
		ret = ERROR_SUCCESS;
		ret = MsiProvideQualifiedComponent(pszaOutlookQualifiedComponents[iVer],
										   _T("outlook.exe"),
									       (DWORD) INSTALLMODE_DEFAULT,
									       NULL,
									       &dwValueBuf);
	}
	else
	{
		b64 = TRUE;
	}

	// get the path we need if we found this version of Outlook.
	if (ret == ERROR_SUCCESS)
	{
		dwValueBuf += 1;
		lpszTempPath = new TCHAR[dwValueBuf];

		if (lpszTempPath != NULL)
		{
			ret = MsiProvideQualifiedComponent(pszaOutlookQualifiedComponents[iVer],
												_T("outlook.x64.exe"),
												(DWORD) INSTALLMODE_DEFAULT,
												lpszTempPath,
												&dwValueBuf);

			if (ret != ERROR_SUCCESS)
			{
				ret = ERROR_SUCCESS;
				ret = MsiProvideQualifiedComponent(pszaOutlookQualifiedComponents[iVer],
												   _T("outlook.exe"),
									               (DWORD) INSTALLMODE_DEFAULT,
												   lpszTempPath,
												   &dwValueBuf);
			}													
		}
	}
	else // the MAPI version the user passed in wasn't found
	{
		 hRes = S_OK; // set S_OK and we'll just use whatever the default MAPI is
		 goto Exit;
	}

	// Now get the MAPI path
	if (lpszTempPath)
	{
		ret = ERROR_SUCCESS;
		TCHAR szDrive[_MAX_DRIVE] = {0};
		TCHAR szOutlookPath[MAX_PATH] = {0};

		ret = _tsplitpath_s(lpszTempPath, szDrive, _MAX_DRIVE, szOutlookPath, MAX_PATH, NULL, NULL, NULL, NULL);

		if (ret == ERROR_SUCCESS)
		{
			szPath = new WCHAR[MAX_PATH];
			if (szPath)
			{
#ifdef UNICODE
					swprintf_s(szPath, MAX_PATH, WszMAPISystemDrivePath, szDrive, szOutlookPath, WszOlMAPI32DLL);
#else
					swprintf_s(szPath, MAX_PATH, szMAPISystemDrivePath, szDrive, szOutlookPath, WszOlMAPI32DLL);
#endif
			}
		}
	}

	if (szPath)
	{
		hMod = LoadLibraryW(szPath);
	}


Exit:
	if (lpszTempPath != NULL)
		delete[] lpszTempPath;

	if (szPath != NULL)
		delete[] szPath;

	return hRes;
}

HRESULT FindAndLoadMAPI(ULONG *ulVer)
{
	HRESULT hRes = E_FAIL;
	HMODULE hMod = NULL;
	LPTSTR lpszTempPath = NULL;
	LPWSTR szPath = NULL;
    TCHAR pszaOutlookQualifiedComponents[][MAX_PATH] = {
        TEXT("{BC174BAD-2F53-4855-A1D5-0D575C19B1EA}"), // Outlook 2003
		TEXT("{24AAE126-0911-478F-A019-07B875EB9996}"), // Outlook 2007
		TEXT("{1E77DE88-BCAB-4C37-B9E5-073AF52DFD7A}"), // Outlook 2010
		TEXT("{E83B4360-C208-4325-9504-0D23003A74A5}"), // Outlook 2013
		TEXT("{5812C571-53F0-4467-BEFA-0A4F47A9437C}"), // Outlook 2016
    };
    int iOQC = _countof(pszaOutlookQualifiedComponents);
	int iVer = 4;
    DWORD dwValueBuf = 0;
    UINT ret = 0;
	BOOL b64 = FALSE;
	BOOL bFound = FALSE;

	for (;iVer >= 0; iVer--) // Find out which one exists - start with latest and work down
	{
		// check for 64-bit first...
		ret = MsiProvideQualifiedComponent(pszaOutlookQualifiedComponents[iVer],
										   _T("outlook.x64.exe"),
										   (DWORD) INSTALLMODE_DEFAULT,
										   NULL,
										   &dwValueBuf);
	
		if (ret != ERROR_SUCCESS)
		{
			ret = ERROR_SUCCESS;
			ret = MsiProvideQualifiedComponent(pszaOutlookQualifiedComponents[iVer],
											   _T("outlook.exe"),
											   (DWORD) INSTALLMODE_DEFAULT,
											   NULL,
											   &dwValueBuf);
			if ( ret == ERROR_SUCCESS)
			{
				bFound = TRUE;
				break;
			}
		}
		else
		{
			b64 = TRUE;
			bFound = TRUE;
			break;
		}
	}

	// get the path we need if we found a version of Outlook/MAPI.
	if (ret == ERROR_SUCCESS)
	{
		dwValueBuf += 1;
		lpszTempPath = new TCHAR[dwValueBuf];

		if (lpszTempPath != NULL)
		{
			ret = MsiProvideQualifiedComponent(pszaOutlookQualifiedComponents[iVer],
												_T("outlook.x64.exe"),
												(DWORD) INSTALLMODE_DEFAULT,
												lpszTempPath,
												&dwValueBuf);

			if (ret != ERROR_SUCCESS)
			{
				ret = ERROR_SUCCESS;
				ret = MsiProvideQualifiedComponent(pszaOutlookQualifiedComponents[iVer],
													_T("outlook.exe"),
													(DWORD) INSTALLMODE_DEFAULT,
													lpszTempPath,
													&dwValueBuf);
			}													
		}
	}
	else // the MAPI version wasn't found
	{
		 hRes = S_OK; // set S_OK and we'll just use whatever the default MAPI is
		 goto Exit;
	}

	if (iVer == 4)
	{
		*ulVer = 16;
		DebugPrint(DBGVerbose, _T("Loading 2016 version of MAPI. \n"));
	}

	else if (iVer == 3)
	{
		*ulVer = 15;
		DebugPrint(DBGVerbose,_T("Loading 2013 version of MAPI. \n"));
	}

	else if (iVer == 2)
	{
		*ulVer = 14;
		DebugPrint(DBGVerbose,_T("Loading 2010 version of MAPI. \n"));
	}

	else if (iVer == 1)
	{
		*ulVer = 12;
		DebugPrint(DBGVerbose,_T("Loading 2007 version of MAPI. \n"));
	}

	else if (iVer == 0)
	{
		*ulVer = 11;
		DebugPrint(DBGVerbose,_T("Loading 2003 version of MAPI. \n"));
	}
	else
	{
		DebugPrint(DBGVerbose,_T("Loading default MAPI version. \n"));
	}

	// Now get the MAPI path
	if (lpszTempPath)
	{
		ret = ERROR_SUCCESS;
		TCHAR szDrive[_MAX_DRIVE] = {0};
		TCHAR szOutlookPath[MAX_PATH] = {0};

		ret = _tsplitpath_s(lpszTempPath, szDrive, _MAX_DRIVE, szOutlookPath, MAX_PATH, NULL, NULL, NULL, NULL);

		if (ret == ERROR_SUCCESS)
		{
			szPath = new WCHAR[MAX_PATH];
			if (szPath)
			{
#ifdef UNICODE
					swprintf_s(szPath, MAX_PATH, WszMAPISystemDrivePath, szDrive, szOutlookPath, WszOlMAPI32DLL);
#else
					swprintf_s(szPath, MAX_PATH, szMAPISystemDrivePath, szDrive, szOutlookPath, WszOlMAPI32DLL);
#endif
			}
		}
	}

	if (szPath)
	{
		hMod = LoadLibraryW(szPath);
	}


Exit:
	if (lpszTempPath != NULL)
		delete[] lpszTempPath;

	if (szPath != NULL)
		delete[] szPath;

	return hRes;
}







