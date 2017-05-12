// ProcessCalendar.cpp - the guts of this app...

#include "stdafx.h"
#include "ProcessCalendar.h"
#include "Logging.h"
#include "Defaults.h"


FILETIME g_ftNow;
ULONG g_ulItems = 0;
ULONG g_ulRecurItems = 0;
BOOL g_bFoundProb = false; // for when a problem is found
BOOL g_bFolder = false; // for the "-f" switch
BOOL g_bMoveItem = false; // set it if we need to move an item
int g_ulErrors = 0;
int g_ulErrorsSvr = 0;
int g_ulWarnings = 0;
int g_ulWarningsSvr = 0;
int g_ulProbItems = 0;
int g_ulProbItemsSvr = 0;
CStringList DupList;
CStringList GOIDDupList;
CString g_strSentRepresenting = "";
CString g_strRecipOrganizer = "";

// MsgClass check strings - update if new common strings are found
CString rgstrMsgClassChk[5] = {"IPM.Appointment", 
							   "IPM.Appointment.Live Meeting Request", 
							   "IPM.Appointment.Location", 
							   "IPM.Appointment.MeetingPlace",  // Cisco MeetingPlace
							   "IPM.Appointment.MP" // Cisco MeetingPlace
						      };

ULONG cMsgClassChk = 5;  // Be sure to update this if you add any strings to the array above


// ApptRecur data checked for in arcvt.txt
CString rgstrApptRecurMain[9] = {"StartDate:",
								 "EndDate:",
								 "EndType:",
								 "ExceptionCount:",
								 "DeletedInstanceCount:",
								 "ModifiedInstanceCount:",
								 "OccurrenceCount:",
								 "StartTimeOffset:",
								 "EndTimeOffset:"
								};

ULONG cApptRecurMain = 9; // Be sure to update this if you add any strings to the array above

// EndType from the recurrence blob
CString rgstrEndTypes[3] = {"IDC_RCEV_PAT_ERB_END",
							"IDC_RCEV_PAT_ERB_AFTERNOCCUR",
							"IDC_RCEV_PAT_ERB_NOEND"
						   };

ULONG cEndTypes = 3; 

CString rgstrBDayText[5] = {"birthday",
							"b-day",
							"anniversaire",
							"geburtstag",
							"cumpleaños"
						   };
ULONG cBDayText = 5;

void printBool(BOOL bBool)
{
	if (bBool) _tprintf(_T("true"));
	else _tprintf(_T("false"));
}

void printFileTime(FILETIME fileTime)
{
	SYSTEMTIME SysTime = {0};

	FileTimeToSystemTime(&fileTime,&SysTime);
	
	_tprintf(_T("%02d/%02d/%4d %02d:%02d:%02d%s"),
		SysTime.wMonth,
		SysTime.wDay,
		SysTime.wYear,
		(SysTime.wHour <= 12)?SysTime.wHour:SysTime.wHour-12,
		SysTime.wMinute,
		SysTime.wSecond,
		(SysTime.wHour <= 12)?_T("AM"):_T("PM"));
/*	printf(" = ");
	
	printf("Low: 0x%08X High: 0x%08X",lpFileTime->dwLowDateTime,lpFileTime->dwHighDateTime);*/
}

// String-ize a Boolean
CString formatBool(BOOL bFormat)
{
	if(bFormat) return(_T("true"));
	else return(_T("false"));
}

//String-ize a FILETIME structure
CString formatFiletime(FILETIME ftFormat)
{
	SYSTEMTIME stTime;
	CString strFormattedTime;

	FileTimeToSystemTime(&ftFormat,&stTime);

	if(stTime.wHour == 0)
		stTime.wHour = 12;

	strFormattedTime.Format(_T("%02d/%02d/%d %02d:%02d:%02d%s"), 
							stTime.wMonth, 
							stTime.wDay, 
							stTime.wYear, 
							(stTime.wHour <= 12)?stTime.wHour:stTime.wHour-12,  // adjust to non-military time
							stTime.wMinute, 
							stTime.wSecond,
							(stTime.wHour <= 12)?_T("AM"):_T("PM"));

	return strFormattedTime;
}

//If more tags are added to the array below, appropriate constants need to be added to the enum.
// for ACL data
enum {
		ePR_MEMBER_ENTRYID, 
		ePR_MEMBER_RIGHTS,  
		ePR_MEMBER_ID, 
		ePR_MEMBER_NAME, 
		NUM_ACL_COLS
	 };

static SizedSPropTagArray(NUM_ACL_COLS, rgACLPropTags) = {
															NUM_ACL_COLS,
															{
																PR_MEMBER_ENTRYID,  // Unique across directory.
																PR_MEMBER_RIGHTS,  
																PR_MEMBER_ID,       // Unique within ACL table. 
																PR_MEMBER_NAME,     // Display name.
															}
														 };



//If more tags are added to the array below, appropriate constants need to be added to the enum.
// for Calendar Items
enum {
		tagSubject,
		tagSentRepresenting,
		tagSentByAddr,
		tagSentByAddrType,
		tagSenderName,
		tagSenderAddr,
		tagMsgClass,
		tagLastModTime,
		tagLastModName,
		tagEID, 
		tagSize,
		tagDeliverTime,
		tagHasAttach,
		tagMsgStatus,
		tagCreateTime,
		tagRecur,
		tagRecurType,
		tagStart, 
		tagEnd, 
		tagStateFlags,
		tagLocation,
		tagTZDesc,
		tagAllDayEvt,
		tagApptRecur,
		tagPidLidRecur,
		tagGlobalObjId,
		tagCleanGlobalObjId,
		tagApptAuxFlags,
		tagIsException,
		tagKeywords,
		tagTZStruct,
		tagTZDefStart,
		tagTZDefEnd,
		tagTZDefRecur,
		tagPropDefStream,
		NUM_ITEMPROP_COLS
	};

//These tags are what we want to check and/or use for finding and reporting "bad" items
static SizedSPropTagArray(NUM_ITEMPROP_COLS,sptCols) = {	NUM_ITEMPROP_COLS,
															PR_SUBJECT_A,//tagSubject
															PR_SENT_REPRESENTING_NAME_A,//tagSentRepresenting
															PR_SENT_REPRESENTING_EMAIL_ADDRESS_A, //tagSentByAddr
															PR_SENT_REPRESENTING_ADDRTYPE_A, //tagSentByAddrType
															PR_SENDER_NAME_A, //tagSenderName
															PR_SENDER_EMAIL_ADDRESS_A, //tagSenderAddr
															PR_MESSAGE_CLASS,//tagMsgClass
															PR_LAST_MODIFICATION_TIME, //tagLastModTime
															PR_LAST_MODIFIER_NAME, //tagLastModName
															PR_ENTRYID,//tagEID
															PR_MESSAGE_SIZE,//tagSize
															PR_MESSAGE_DELIVERY_TIME, //tagDeliverTime
															PR_HASATTACH,//tagHasAttach
															PR_MSG_STATUS,//tagMsgStatus
															PR_CREATION_TIME, //tagCreateTime
															PR_NULL, //tagRecur
															PR_NULL, //tagRecurType
															PR_NULL, //tagStart
															PR_NULL, //tagEnd
															PR_NULL, //tagStateFlags
															PR_NULL, //tagLocation
															PR_NULL, //tagTZDesc
															PR_NULL, //tagAllDayEvt
															PR_NULL, //tagApptRecur
															PR_NULL, //tagPidLidRecur
															PR_NULL, //tagGlobalObjId
															PR_NULL, //tagCleanGlobalObjId
															PR_NULL, //tagApptAuxFlags
															PR_NULL, //tagIsException
															PR_NULL, //tagKeywords
															PR_NULL, //tagTZStruct
															PR_NULL, //tagTZDefStart
															PR_NULL, //tagTZDefEnd
															PR_NULL, //tagTZDefRecur
															PR_NULL  //tagPropDefStream
														};


HRESULT ProcessCalendar(LPTSTR szDN, LPMDB lpMDB, LPMAPIFOLDER lpFolder)
{
	if (!lpFolder) return MAPI_E_INVALID_PARAMETER;

	HRESULT			hRes = S_OK;
	LPMAPITABLE		lpContents = NULL;
	LPMAPIFOLDER	lpCalCheckFld = NULL;
	LPSRowSet		lpCalendarRows = NULL;
	LPSPropTagArray lpNamedPropTags = NULL; // for Named Props
	LPSPropTagArray	lpPropTags = NULL;  // for all props
	LPSPropValue	pv = NULL;
	CString			strCalendar = _T("Calendar");
	CString			strFolder = _T("");
	CString			strLogging = _T("");
	SYSTEMTIME		stSysTime;
	FILETIME		ftFileTime;
	TIME_ZONE_INFORMATION TZInfo;
	CString			strLocalTZ = _T("");
	CString			strHRes = _T("");
	LPMAPIFOLDER    lpOnlineCalendar = NULL;
	LPMAPITABLE		lpOnlineContents = NULL;
	LPSRowSet		lpOnlineRows = NULL;
	ULONG			ulOnlineCount = 0;
	CString			strOnlineCount = _T("");
	CString			strCount = _T("");
	BOOL			bDoCountCheck = true;

	// create/update a global "now" time for checks during processing
	GetLocalTime(&stSysTime);
	SystemTimeToFileTime(&stSysTime, &ftFileTime);
	g_ftNow = ftFileTime;

	//get and log the local time zone
	GetTimeZoneInformation(&TZInfo);
	strLocalTZ = TZInfo.StandardName;
	WriteLogFile(_T("Local time zone: " + strLocalTZ), 0);

	// Get the name of the folder and log it
	hRes = HrGetOneProp(lpFolder, PR_DISPLAY_NAME, &pv);
	if (SUCCEEDED(hRes) && pv)
	{
		strFolder = pv->Value.lpszW;
		// removing check for "Calendar" since that is only the English name... (feedback from website)
		// let the user know we successfully opened the folder
		WriteLogFile(_T("Successfully opened the ") + strFolder + _T(" folder."), 0);

	}
	else
	{
		WriteLogFile(_T("Error retrieving folder display name. Check the folder property PR_DISPLAY_NAME for the Calendar folder before continuing."), 0);
		goto Exit;
	}
	

	// if -F then report to the user
	if (g_bFolder)
	{
		WriteLogFile(_T("CalCheck will move flagged error items to the CalCheck folder in the Outlook folder list."), 0);
		DebugPrint(DBGGeneric,_T("CalCheck will move flagged error items to the CalCheck folder in the Outlook folder list.\n"));
		WriteLogFile(_T("If there are no problems - then no folder will be created."), 0);
		DebugPrint(DBGGeneric,_T("If there are no problems - then no folder will be created.\n"));
	}
	
	DebugPrint(DBGVerbose,_T("\n"));
	DebugPrint(DBGTime,_T("Begin ProcessCalendar\n"));
	DebugPrint(DBGGeneric,_T("\n"));
	DebugPrint(DBGGeneric,_T("Processing calendar for %s\n"),g_MBXDisplayName);

	WriteLogFile(_T("Processing calendar for " + g_MBXDisplayName) + "\r\n", 0);

	DebugPrint(DBGVerbose,_T("Checking delegate & permission information.\n"));

	// Get Delegate names from the localFreeBusy message and write results to the LOG
	hRes = GetDelegates(lpMDB);

	// Get ACL/Permission data for the Calendar and write results out to the LOG
	hRes = GetACLs(lpFolder);
	
	DebugPrint(DBGVerbose,_T("Completed delegate & permission check.\n"));
	
#define NUMNAMEDPROPS 20  // Be sure to update this if you add any named props
	MAPINAMEID  rgnmid[NUMNAMEDPROPS];
	LPMAPINAMEID rgpnmid[NUMNAMEDPROPS];
	
	// Add Named Props that need to be checked / used here
	rgnmid[0].lpguid = (LPGUID)&PSETID_Appointment;
	rgnmid[0].ulKind = MNID_ID;
	rgnmid[0].Kind.lID = dispidRecurring;
	rgpnmid[0] = &rgnmid[0];

	rgnmid[1].lpguid = (LPGUID)&PSETID_Appointment;
	rgnmid[1].ulKind = MNID_ID;
	rgnmid[1].Kind.lID = dispidRecurType;
	rgpnmid[1] = &rgnmid[1];

	rgnmid[2].lpguid = (LPGUID)&PSETID_Appointment;
	rgnmid[2].ulKind = MNID_ID;
	rgnmid[2].Kind.lID = dispidApptStartWhole;
	rgpnmid[2] = &rgnmid[2];

	rgnmid[3].lpguid = (LPGUID)&PSETID_Appointment;
	rgnmid[3].ulKind = MNID_ID;
	rgnmid[3].Kind.lID = dispidApptEndWhole;
	rgpnmid[3] = &rgnmid[3];

	rgnmid[4].lpguid = (LPGUID)&PSETID_Appointment;
	rgnmid[4].ulKind = MNID_ID;
	rgnmid[4].Kind.lID = dispidApptStateFlags;
	rgpnmid[4] = &rgnmid[4];

	rgnmid[5].lpguid = (LPGUID)&PSETID_Appointment;
	rgnmid[5].ulKind = MNID_ID;
	rgnmid[5].Kind.lID = dispidLocation;
	rgpnmid[5] = &rgnmid[5];

	rgnmid[6].lpguid = (LPGUID)&PSETID_Appointment;
	rgnmid[6].ulKind = MNID_ID;
	rgnmid[6].Kind.lID = dispidTimeZoneDesc;
	rgpnmid[6] = &rgnmid[6];

	rgnmid[7].lpguid = (LPGUID)&PSETID_Appointment;
	rgnmid[7].ulKind = MNID_ID;
	rgnmid[7].Kind.lID = dispidApptSubType; //tagAllDayEvt
	rgpnmid[7] = &rgnmid[7];

	rgnmid[8].lpguid = (LPGUID)&PSETID_Appointment;
	rgnmid[8].ulKind = MNID_ID;
	rgnmid[8].Kind.lID = dispidApptRecur; //RecurBlob NamedProp
	rgpnmid[8] = &rgnmid[8];

	rgnmid[9].lpguid = (LPGUID)&PSETID_Meeting;
	rgnmid[9].ulKind = MNID_ID;
	rgnmid[9].Kind.lID = PidLidIsRecurring;
	rgpnmid[9] = &rgnmid[9];

	rgnmid[10].lpguid = (LPGUID)&PSETID_Meeting;
	rgnmid[10].ulKind = MNID_ID;
	rgnmid[10].Kind.lID = PidLidGlobalObjectId;
	rgpnmid[10] = &rgnmid[10];

	rgnmid[11].lpguid = (LPGUID)&PSETID_Meeting;
	rgnmid[11].ulKind = MNID_ID;
	rgnmid[11].Kind.lID = PidLidCleanGlobalObjectId;
	rgpnmid[11] = &rgnmid[11];

	rgnmid[12].lpguid = (LPGUID)&PSETID_Appointment;
	rgnmid[12].ulKind = MNID_ID;
	rgnmid[12].Kind.lID = dispidApptAuxFlags;
	rgpnmid[12] = &rgnmid[12];

	rgnmid[13].lpguid = (LPGUID)&PSETID_Meeting;
	rgnmid[13].ulKind = MNID_ID;
	rgnmid[13].Kind.lID = PidLidIsException;
	rgpnmid[13] = &rgnmid[13];

	rgnmid[14].lpguid = (LPGUID)&PS_PUBLIC_STRINGS;
	rgnmid[14].ulKind = MNID_STRING;
	rgnmid[14].Kind.lpwstrName = L"Keywords";
	rgpnmid[14] = &rgnmid[14];

	rgnmid[15].lpguid = (LPGUID)&PSETID_Appointment;
	rgnmid[15].ulKind = MNID_ID;
	rgnmid[15].Kind.lID = dispidTimeZoneStruct;
	rgpnmid[15] = &rgnmid[15];

	rgnmid[16].lpguid = (LPGUID)&PSETID_Appointment;
	rgnmid[16].ulKind = MNID_ID;
	rgnmid[16].Kind.lID = dispidApptTZDefStartDisplay;
	rgpnmid[16] = &rgnmid[16];

	rgnmid[17].lpguid = (LPGUID)&PSETID_Appointment;
	rgnmid[17].ulKind = MNID_ID;
	rgnmid[17].Kind.lID = dispidApptTZDefEndDisplay;
	rgpnmid[17] = &rgnmid[17];

	rgnmid[18].lpguid = (LPGUID)&PSETID_Appointment;
	rgnmid[18].ulKind = MNID_ID;
	rgnmid[18].Kind.lID = dispidApptTZDefRecur;
	rgpnmid[18] = &rgnmid[18];

	rgnmid[19].lpguid = (LPGUID)&PSETID_Common;
	rgnmid[19].ulKind = MNID_ID;
	rgnmid[19].Kind.lID = dispidPropDefStream;
	rgpnmid[19] = &rgnmid[19];

	// Get the IDs for these named props from the store
	hRes = lpFolder->GetIDsFromNames(
		NUMNAMEDPROPS, 
		rgpnmid, 
		NULL, 
		&lpNamedPropTags);

	if (SUCCEEDED(hRes) && lpNamedPropTags)
	{
		// add new named props here as well if added/changed above
		sptCols.aulPropTag[tagRecur] = CHANGE_PROP_TYPE(lpNamedPropTags->aulPropTag[0],PT_BOOLEAN);
		sptCols.aulPropTag[tagRecurType] = CHANGE_PROP_TYPE(lpNamedPropTags->aulPropTag[1],PT_LONG);
		sptCols.aulPropTag[tagStart] = CHANGE_PROP_TYPE(lpNamedPropTags->aulPropTag[2],PT_SYSTIME);
		sptCols.aulPropTag[tagEnd] = CHANGE_PROP_TYPE(lpNamedPropTags->aulPropTag[3],PT_SYSTIME);
		sptCols.aulPropTag[tagStateFlags] = CHANGE_PROP_TYPE(lpNamedPropTags->aulPropTag[4],PT_LONG);
		sptCols.aulPropTag[tagLocation] = CHANGE_PROP_TYPE(lpNamedPropTags->aulPropTag[5],PT_STRING8);
		sptCols.aulPropTag[tagTZDesc] = CHANGE_PROP_TYPE(lpNamedPropTags->aulPropTag[6],PT_STRING8);
		sptCols.aulPropTag[tagAllDayEvt] = CHANGE_PROP_TYPE(lpNamedPropTags->aulPropTag[7],PT_BOOLEAN);
		sptCols.aulPropTag[tagApptRecur] = CHANGE_PROP_TYPE(lpNamedPropTags->aulPropTag[8],PT_BINARY);
		sptCols.aulPropTag[tagPidLidRecur] = CHANGE_PROP_TYPE(lpNamedPropTags->aulPropTag[9],PT_BOOLEAN);
		sptCols.aulPropTag[tagGlobalObjId] = CHANGE_PROP_TYPE(lpNamedPropTags->aulPropTag[10],PT_BINARY);
		sptCols.aulPropTag[tagCleanGlobalObjId] = CHANGE_PROP_TYPE(lpNamedPropTags->aulPropTag[11],PT_BINARY);
		sptCols.aulPropTag[tagApptAuxFlags] = CHANGE_PROP_TYPE(lpNamedPropTags->aulPropTag[12],PT_LONG);
		sptCols.aulPropTag[tagIsException] = CHANGE_PROP_TYPE(lpNamedPropTags->aulPropTag[13],PT_BOOLEAN);
		sptCols.aulPropTag[tagKeywords] = CHANGE_PROP_TYPE(lpNamedPropTags->aulPropTag[14],PT_MV_UNICODE);
		sptCols.aulPropTag[tagTZStruct] = CHANGE_PROP_TYPE(lpNamedPropTags->aulPropTag[15],PT_BINARY);
		sptCols.aulPropTag[tagTZDefStart] = CHANGE_PROP_TYPE(lpNamedPropTags->aulPropTag[16],PT_BINARY);
		sptCols.aulPropTag[tagTZDefEnd] = CHANGE_PROP_TYPE(lpNamedPropTags->aulPropTag[17],PT_BINARY);
		sptCols.aulPropTag[tagTZDefRecur] = CHANGE_PROP_TYPE(lpNamedPropTags->aulPropTag[18],PT_BINARY);
		sptCols.aulPropTag[tagPropDefStream] = CHANGE_PROP_TYPE(lpNamedPropTags->aulPropTag[19],PT_BINARY);
	}

	//Put all the props that we want to retreive into the PTagArray 
	lpPropTags = (LPSPropTagArray)&sptCols;

	// Get the count of items in the Online Store Calendar to compare with the local copy. If this is an online profile, then these will be the same anyway...
	hRes = OpenDefaultFolder(DEFAULT_CALENDAR, lpMDB, MAPI_BEST_ACCESS | MAPI_NO_CACHE, &lpOnlineCalendar);
	if (SUCCEEDED(hRes) && lpOnlineCalendar)
	{
		hRes = lpOnlineCalendar->GetContentsTable(0, &lpOnlineContents);
		if (SUCCEEDED(hRes) && lpOnlineContents)
		{
			hRes = lpOnlineContents->GetRowCount(0, &ulOnlineCount);
			if (SUCCEEDED(hRes))
			{
				strOnlineCount.Format(_T("%u"), ulOnlineCount);
			}
			else
			{
				ulOnlineCount = 0;
				bDoCountCheck = false;

			}
			lpOnlineContents->Release();
		}
		else
		{
			ulOnlineCount = 0;
			bDoCountCheck = false;
		}
		lpOnlineCalendar->Release();
	}
	else
	{
		ulOnlineCount = 0;
		bDoCountCheck = false;
	}

	// Get the Contents of the Calendar folder
	hRes = lpFolder->GetContentsTable(0, &lpContents);
	if (SUCCEEDED(hRes) && lpContents)
	{
		// Get total count of items 
		hRes = HrQueryAllRows(lpContents, NULL, NULL, NULL, 0, &lpCalendarRows);
		if (SUCCEEDED(hRes))
		{
			g_ulItems = lpCalendarRows->cRows; // set the global for use later...
			strLogging.Format(_T("%u"), lpCalendarRows->cRows);

			// now check online count vs cached count and report if different
			if (bDoCountCheck)
			{
				if(g_ulItems != ulOnlineCount)
				{
					WriteLogFile(_T("Count of Online and Cached items do not match. \r\n"), 0);
					WriteLogFile(_T("Online count: " + strOnlineCount + "   Cached count: " + strLogging + " \r\n"), 0);
					WriteLogFile(_T("This indicates a syncing problem. Check the Sync Issues folder, and try clearing offline items and then syncing again.\r\n"), 0);
					DebugPrint(DBGGeneric,_T("Count of Online and Cached items do not match.\n" ));
					DebugPrint(DBGGeneric,_T("Online count: %s   Cached count: %s \n" ),strOnlineCount, strLogging);
					DebugPrint(DBGGeneric,_T("This indicates a syncing problem. Check the Sync Issues folder, and try clearing offline items and then syncing again.\n" ));
				}
			}
			else
			{
				WriteLogFile(_T("Couldn't retreive Online item count for comparison with Cached item count. Continuing... \r\n"), 0);
			}

			if (lpCalendarRows->cRows == 1)
			{
				WriteLogFile(_T("Found " + strLogging + " item in the Calendar. \r\n"), 0);
				DebugPrint(DBGGeneric,_T("Found %s item in this Calendar.\n" ),strLogging);				
			}
			else if (lpCalendarRows->cRows == 0)
			{
				WriteLogFile(_T("Found " + strLogging + " items in the Calendar - skipping. \r\n"), 0);
				DebugPrint(DBGGeneric,_T("Found %s items in this Calendar - skipping.\n" ),strLogging);

				goto Exit;
			}
			else
			{
				WriteLogFile(_T("Found " + strLogging + " items in the Calendar. Processing... \r\n"), 0);
				DebugPrint(DBGGeneric,_T("Found %s items in this Calendar.\n" ),strLogging);
			}
			
			//clean up
			strLogging = _T("");
		}
		else
		{
			goto Exit;
		}

		hRes = lpContents->SetColumns(lpPropTags, 0);
		if (SUCCEEDED(hRes))
		{
			// try to sort the Calendar items by PR_CREATION_TIME so we get oldest to newest items - needed for Duplicate GOID test.
			LPSSortOrderSet lpSortOrder = NULL;
			hRes = MAPIAllocateBuffer(CbNewSSortOrderSet(1),(LPVOID*)&lpSortOrder);

			if (SUCCEEDED(hRes))
			{
				lpSortOrder->cSorts = 1;
				lpSortOrder->cCategories = 0;
				lpSortOrder->cExpanded = 0;
				lpSortOrder->aSort[0].ulPropTag = PR_CREATION_TIME;
				lpSortOrder->aSort[0].ulOrder = TABLE_SORT_ASCEND;
				
				hRes = lpContents->SortTable(lpSortOrder, TBL_ASYNC);	
			}			
			if (FAILED(hRes))
			{
				strHRes.Format(_T("%08X"), hRes);
				WriteLogFile(_T("Failed to sort table by item creation time. Duplicate Global Object ID test may give skewed results.\r\n"), 0);
				WriteLogFile(_T("Error: ") + strHRes, 0);
				DebugPrint(DBGGeneric,_T("Failed to sort table by item creation time. Duplicate Global Object ID test may give skewed results. Error 0x%08X\r\n"),hRes);
				hRes = S_OK;
			}

			hRes = lpContents->SeekRow(BOOKMARK_BEGINNING, 0, NULL);
			if (SUCCEEDED(hRes))
			{
				if(g_bCSV)
				{
					CString strUserName = g_MBXDisplayName;
					StringFix(&strUserName);
					// Output Headers for data in CALITEMS.CSV
					WriteCSVFile(_T("Calendar For :  " + strUserName + "  :  " + (CString)szDN + "\r"));
					WriteCSVFile(_T("Subject, Location, Size, Start Date, Start Time, End Date, End Time, Msg Class, Last Modifier, Last Mod Date, Last Mod Time, Num Recips, Recurring, Attachments, All Day Event, "));
					WriteCSVFile(_T("Organizer Name, Organizer Address, Sender Name, Sender Address, Appt State Flags, Recur Start Date, Recur End Date, Recur End Type, Num Exceptions, Deleted Instances, "));
					WriteCSVFile(_T("Modified Instances, Timezone Struct, Flagged by CalCheck\r"));						  
				}

				// output headers for data in CalCheckErr.csv
				if (g_bEmitErrorNum)
					WriteErrorFile(_T("Subject, Location, Start Time, End Time, Recurring, Organizer, Is Past Item, Errors and Warnings, Other Item Subject, Other Item Start, Other Item End, EntryID, ErrorNum\r"));
				else
					WriteErrorFile(_T("Subject, Location, Start Time, End Time, Recurring, Organizer, Is Past Item, Errors and Warnings, Other Item Subject, Other Item Start, Other Item End\r"));
	
				// start progress counter in command window
				DebugPrint(DBGGeneric,_T("Processing items: "));

				LPSRowSet pRows = NULL;
				for (;;)
				{
					// Update command window with progress - happens every 32 items
					DebugPrint(DBGHelp,_T("."));
									
					// Do 32 at a time...
					hRes = lpContents->QueryRows(50, 0, &pRows);
					if (FAILED(hRes)) break;
					if (!pRows || !pRows->cRows) break;

					ULONG i = 0;
					for (i = 0 ; i < pRows->cRows ; i++)
					{
						// Process the rows / items - checking for problems
						ProcessRow(szDN, lpFolder, &pRows->aRow[i]);

						if (g_bFolder && g_bMoveItem) // if folder option specified and moving an error item 
						{
							if(!lpCalCheckFld) // if the CalCheck folder isn't there yet then create it
							{
								hRes = CreateCalCheckFolder(lpMDB, &lpCalCheckFld);

								// if we couldn't create the folder, log it and continue
								if (!SUCCEEDED(hRes) || !lpCalCheckFld)
								{
									CString		strHR = _T("");
									strHR.Format(_T("%08X"), hRes);
									WriteLogFile(_T("Could not create the CalCheck folder. Error 0x") + strHR, 0);
									DebugPrint(DBGGeneric,_T("Creating CalCheck folder returned error 0x%08X\n"),hRes);
								}
							}

							LPENTRYLIST lpEIDs = NULL;
							WC_H(MAPIAllocateBuffer(sizeof(ENTRYLIST),(LPVOID*) &lpEIDs));

							if(lpEIDs)
							{
								lpEIDs->cValues = 1;
								lpEIDs->lpbin = NULL;

								WC_H(MAPIAllocateMore((ULONG)sizeof(SBinary), lpEIDs, (LPVOID*) &lpEIDs->lpbin));

								if(lpEIDs->lpbin)
								{
									lpEIDs->lpbin[0].cb = pRows->aRow[i].lpProps[tagEID].Value.bin.cb;
									lpEIDs->lpbin[0].lpb = pRows->aRow[i].lpProps[tagEID].Value.bin.lpb;
								}
							}

							// move this item to the CalCheck folder
							hRes = lpFolder->CopyMessages(lpEIDs, &IID_IMAPIFolder, lpCalCheckFld, NULL, NULL, MESSAGE_MOVE);

							if(FAILED(hRes)) // log failure to move the item...
							{
								CString strError;
								strError.Format(_T("%08X"), hRes);
								WriteLogFile(_T("Could not move this item to the CalCheck folder. Error: ") + strError, 0);
							}

							g_bMoveItem = false; // set the move item global back to false for further processing
						}

					}
					FreeProws(pRows);
					pRows = NULL;
				}

				// end the progress counter
				DebugPrint(DBGHelp,_T("\n"));
				DebugPrint(DBGGeneric,_T("Done!\n"));

				// For server mode - need to clear the duplicate list before doing another Calendar folder
				if(!DupList.IsEmpty())
				{
					DupList.RemoveAll();
				}

				// clear duplicate GOID item list too...
				if(!GOIDDupList.IsEmpty())
				{
					GOIDDupList.RemoveAll();
				}

			} // if Row

		} // if Columns

	} // if Contents

Exit: 

	DebugPrint(DBGVerbose,_T("\n"));
	DebugPrint(DBGTime,_T("End ProcessCalendar\n"));
	DebugPrint(DBGVerbose,_T("\n"));

	if (lpContents) lpContents->Release();

	return hRes;
}

// Look at several properties on each item and see if a problem exists...
HRESULT ProcessRow(LPTSTR szDN, LPMAPIFOLDER lpFolder, LPSRow lpsRow)
{
	HRESULT hRes = S_OK;

	CString strError = _T("");
	CString strSubject = _T("");
	CString strLocation = _T("");
	CString strSentRepresenting = _T("");
	CString strSentByAddr = _T("");
	CString strSentByAddrType = _T("");
	CString strMsgClass = _T("");
	CString strLastModName = _T("");
	CString strApptRecurBin = _T("");
	CString strSenderName = _T("");
	CString strSenderAddr = _T("");
	CString strDN = szDN;
	BOOL bRecur = false;
	BOOL bPidLidRecur = false;
	BOOL bIsException = false;
	ULONG ulRecurType = 0;
	FILETIME ftLastMod = {0};
	FILETIME ftStart = {0};
	FILETIME ftEnd = {0};
	FILETIME ftDeliver = { 0 };
	ULONG ulSize = 0;
	BOOL bHasAttach = false;
	ULONG ulNumRecips = 0;
	ULONG ulApptStateFlags = 0;
	BOOL bIsOrganizer = false;
	BOOL bIsAllDayEvt = false;
	BOOL bLogProps = false;
	ULONG ulFt = 0;
	CString strRecurStart = _T("");
	CString strRecurEnd = _T("");
	CString strEndType = _T("");
	CString strNumOccur = _T("");
	CString strNumExceptions = _T("");
	CString strNumOccurDeleted = _T("");
	CString strNumOccurModified = _T("");
	FILETIME ftRecurStart = {0};
	FILETIME ftRecurEnd = {0};
	ULONG ulFTCompare = 0;
	ULONG ulMsgStatus = 0;
	CString strGlobalObjId = _T("");
	CString strCleanGlobalObjId = _T("");
	CString strGOIDupSubject = _T("");
	ULONG ulApptAuxFlags = 0;
	ULONG ulNumAttach = 0;
	ULONG ulPastItem = 0;
	ULONG ulPastRecurItem = 0;
	CString strTZStruct = _T("");
	CString strTZDefRecur = _T("");
	CString strRptSubject = _T(""); // for reporting dup GOID
	CString strRptStart = _T("");  // for reporting dup GOID
	CString strRptEnd = _T("");  // for reporting dup GOID
	BOOL bBdayText = false;
	BOOL bErrFound = false;
	CString strEntryID = _T(""); // Output for OffCAT
	CString strErrorNum = _T(""); // Output for OffCAT
	
	strDN.MakeLower();
	

	//Get duplicate check props first and then do tests on them later...
	//An error on other tests will drop to the duplicate check, so we need these props up front

	if (PR_SUBJECT_A == lpsRow->lpProps[tagSubject].ulPropTag)
	{
		strSubject = lpsRow->lpProps[tagSubject].Value.lpszA;
		
		// See if this is a "Birthday" set in the calendar
		if (!(strSubject.IsEmpty()))
		{
			bBdayText = CheckSubject(strSubject);
		}
	}

	if (PR_SENT_REPRESENTING_NAME_A == lpsRow->lpProps[tagSentRepresenting].ulPropTag)
	{
		strSentRepresenting = lpsRow->lpProps[tagSentRepresenting].Value.lpszA;
	}

	// dispidLocation
	if (PT_STRING8 == PROP_TYPE(lpsRow->lpProps[tagLocation].ulPropTag))
	{
		strLocation = lpsRow->lpProps[tagLocation].Value.lpszA;
	}

	// dispidRecurring
	if (PT_BOOLEAN == PROP_TYPE(lpsRow->lpProps[tagRecur].ulPropTag))
	{
		bRecur = lpsRow->lpProps[tagRecur].Value.b;
		if (CalCheckTests.bRecurItemLimit)
		{
			if(bRecur)
				g_ulRecurItems++; // if it's recurring, then increment this global count.
		}
	}

	// dispidApptStartWhole
	if (PT_SYSTIME == PROP_TYPE(lpsRow->lpProps[tagStart].ulPropTag))
	{
		ftStart = lpsRow->lpProps[tagStart].Value.ft;
	}

	// dispidApptEndWhole
	if (PT_SYSTIME == PROP_TYPE(lpsRow->lpProps[tagEnd].ulPropTag))
	{
		ftEnd = lpsRow->lpProps[tagEnd].Value.ft;
	}

	// End getting Duplicate Item check props


	if (PR_ENTRYID == lpsRow->lpProps[tagEID].ulPropTag)
	{
		ULONG cbBinVal = lpsRow->lpProps[tagEID].Value.bin.cb;
		LPBYTE lpBinVal = lpsRow->lpProps[tagEID].Value.bin.lpb;

		strEntryID = Binary2String(cbBinVal, lpBinVal); //string-ize the blob as hex text for potential output
	}

	if (!(PR_MESSAGE_DELIVERY_TIME == lpsRow->lpProps[tagDeliverTime].ulPropTag))
	{
		if (CalCheckTests.bDeliverTime)
		{
			strError += _T("ERROR: Missing required prop PR_MESSAGE_DELIVERY_TIME. This can cause Calendar sync problems.  ");
			strErrorNum += _T("0059 ");

			bLogProps = true;
			g_bMoveItem = true;
			g_ulErrors++;
			//goto Exit;
		}
	}
	else // not sure if the prop would be needed if it's there - but setting it in case needed at some point in the future...
	{
		ftDeliver = lpsRow->lpProps[tagDeliverTime].Value.ft;
	}

	// if this is a meeting conflict item and the folder switch was used - then move it to the CalCheck folder 
	// and then go to the next item
	if (CalCheckTests.bConflictItems)
	{
		if (PR_MSG_STATUS == lpsRow->lpProps[tagMsgStatus].ulPropTag)
		{
			ulMsgStatus = lpsRow->lpProps[tagMsgStatus].Value.l;

			if (ulMsgStatus & MSGSTATUS_IN_CONFLICT)
			{
				if(g_bFolder)
				{
					strError += _T("INFO: Moving conflict item to CalCheck folder.  ");
					strErrorNum += _T("0062 ");
				
					bLogProps = true;
					g_bMoveItem = true;
					goto Exit;
				}
			}
		}
	}

	// "Keywords"
	if (PT_MV_UNICODE == PROP_TYPE(lpsRow->lpProps[tagKeywords].ulPropTag))
	{
		// Look for "Holiday" - if so then ignore it.
		ULONG cKeyVals = lpsRow->lpProps[tagKeywords].Value.MVi.cValues;
		for(int i = 0; i < cKeyVals; i++)
		{
			CString strKeyVal = lpsRow->lpProps[tagKeywords].Value.MVszW.lppszW[i];

			if (!CalCheckTests.bHolidayItems)
			{
				if (strKeyVal.Compare(_T("Holiday")) == 0)
				{
					// Yeah - this is a Holiday Item - so skip it.
					goto Exit;
				}
			}
		}		
	}

	// check dispidApptEndWhole - in case we want to skip reporting old items
	if (PT_SYSTIME == PROP_TYPE(lpsRow->lpProps[tagEnd].ulPropTag))
	{
		// ftEnd = lpsRow->lpProps[tagEnd].Value.ft;   <<<<<< Set up above with Dup Check props
		// if this is NOT a recurring item and ended in the past - we'll skip it if set to skip past items
		if(!bRecur)
		{
			ulPastItem = CompareFileTime(&ftEnd, &g_ftNow);
			if (ulPastItem == -1)
			{
				if (!CalCheckTests.bPastItems)
				{
					goto Exit;
				}
			}
		}
	}

	//dispidApptSubType  >> Is this an All Day Event?
	if (PT_BOOLEAN == PROP_TYPE(lpsRow->lpProps[tagAllDayEvt].ulPropTag))
	{
		bIsAllDayEvt = lpsRow->lpProps[tagAllDayEvt].Value.b;

		if (bIsAllDayEvt && bBdayText) // All Day, and has "Birthday" in the subject
		{
			// This is a Birthday someone put in the calendar - so skip it unless test is turned on
			if(!CalCheckTests.bBirthdayItems)
			{
				goto Exit;
			}
		}
	}

	// dispidRecurring (checking if it's NOT there)
	if (!(PT_BOOLEAN == PROP_TYPE(lpsRow->lpProps[tagRecur].ulPropTag)))
	{
		if (CalCheckTests.bRecurringProp)
		{
			strError += _T("ERROR: Missing required prop dispidRecurring. See KB 827432 https://support.microsoft.com/en-us/kb/827432  ");
			strErrorNum += _T("0026 ");
		
			bLogProps = true;
			g_bMoveItem = true;
			g_ulErrors++;
			//goto Exit;
		}
	}

	//dispidTimeZoneStruct	
	if (PT_BINARY == PROP_TYPE(lpsRow->lpProps[tagTZStruct].ulPropTag))
	{
		ULONG cbBinVal = lpsRow->lpProps[tagTZStruct].Value.bin.cb;
		LPBYTE lpBinVal = lpsRow->lpProps[tagTZStruct].Value.bin.lpb;

		strTZStruct = Binary2String(cbBinVal, lpBinVal); //string-ize the blob as hex text
		// REMOVED THIS FOR NOW UNTIL I CAN KNOW HOW TO TRULY DETECT THIS PROBLEM - IF IT CAN BE DONE AT ALL.
		// Didn't find a way to detect the Mac Outlook problem here. So the trick will be to use the "-A" switch and find the items using the 
		// dispidTimeZoneStruct property in the resulting CSV file...
		//
		/*if (strTZStruct.Left(24) == "000000000000000000000000")
		{
			WriteLogFile(_T("ERROR: Timezone bias is set to zero / GMT. This may have happened via interaction between Exchange and Mac Outlook."), 1);
							
			bLogProps = true;
			g_bMoveItem = true;
			g_ulErrors++;
		}*/
		
	}

	//dispdApptTZDefRecur
	if (PT_BINARY == PROP_TYPE(lpsRow->lpProps[tagTZDefRecur].ulPropTag))
	{
		ULONG cbBinVal = lpsRow->lpProps[tagTZDefRecur].Value.bin.cb;
		LPBYTE lpBinVal = lpsRow->lpProps[tagTZDefRecur].Value.bin.lpb;

		if (CalCheckTests.bTZDefRecur)
		{
			if (bRecur) // if recurring and this prop got munged then opening single instances of the meeting will error out.
			{
				TZDEFINITION* ptzDefRecur = BinToTZDEFINITION(cbBinVal, lpBinVal);

				if (ptzDefRecur != NULL)
				{
					strTZDefRecur = ptzDefRecur->pwszKeyName;
					if (strTZDefRecur.IsEmpty())
					{
						strError += _T("ERROR: Missing required data in prop dispidApptTZDefRecur. Please check this item.  ");
						strErrorNum += _T("0055 ");

						bLogProps = true;
						g_bMoveItem = true;
						g_ulErrors++;
					}
				}
				else
				{
					strError += _T("ERROR: Missing required data in prop dispidApptTZDefRecur. Please check this item.  ");
					strErrorNum += _T("0055 ");

					bLogProps = true;
					g_bMoveItem = true;
					g_ulErrors++;
				}
			}
		}
	}

	// dispidApptTZDefStartDisplay (testing if it is NOT there.. that's bad and caused a crashing problem for a customer)
	if (!(PT_BINARY == PROP_TYPE(lpsRow->lpProps[tagTZDefStart].ulPropTag)))
	{
		if (CalCheckTests.bTZDefStart)
		{
			if (bRecur) // looks like NOT having this on non-recurring items is okay...
			{
				strError += _T("ERROR: Missing required prop dispidApptTZDefStartDisplay. Please check this item.  ");
				strErrorNum += _T("0027 ");
		
				bLogProps = true;
				g_bMoveItem = true;
				g_ulErrors++;
				//goto Exit;
			}
		}
	}

	// dispidPropDefStream
	if (PT_BINARY == PROP_TYPE(lpsRow->lpProps[tagPropDefStream].ulPropTag))
	{
		if (CalCheckTests.bPropDefStream)  // right now there's only one test in the PropDefStream - might need to be changed if more tests come up.
		{
			BOOL bLogError = false;
			CString strRetError = _T("");
			CString strPDSBlob = _T("");

			if(g_bMRMAPI) // we will do this if MRMAPI is where it needs to be.
			{
				// if it's there - then string-ize the blob and convert it using MRMAPI 
				ULONG cbBinVal = lpsRow->lpProps[tagPropDefStream].Value.bin.cb;
				LPBYTE lpBinVal = lpsRow->lpProps[tagPropDefStream].Value.bin.lpb;
				BOOL bConverted = true;

				strPDSBlob = Binary2String(cbBinVal, lpBinVal); //string-ize the blob as hex text for conversion

				if(!strPDSBlob.IsEmpty()) // check to see if it's empty - if so then skip the conversion and throw an error out to the user.
				{
					bConverted = ConvertPropDefStream(strPDSBlob, &strRetError);

					if (bConverted)
					{
						bLogError = GetPropDefStreamData(&strError,
														 &strErrorNum);
						if (bLogError)
						{
							//strError += strRetError;
							//strErrorNum += strErrorNum;
							bLogProps = true;
							g_bMoveItem = true;
							g_ulErrors++;
						}
					}
					else
					{
						if(g_bMRMAPI) // if MrMAPI is present, then it's not FileNotFound - it's something else
						{
							strError += strRetError;
							strErrorNum += _T("0057 "); // this could be one of three errors - all having to do with the prop not happening when calling MrMapi.
														//  From Logging.cpp in ConvertPropDefStream
							bLogProps = true;
							g_bMoveItem = true;
							g_ulErrors++;
						}
					}
				}
			}
		}
	}

	// dispidApptRecur
	if (PT_BINARY == PROP_TYPE(lpsRow->lpProps[tagApptRecur].ulPropTag))
	{
		BOOL bLogError = false;
		CString strRetError = _T("");

		if(g_bMRMAPI) // we will do this if MRMAPI is where it needs to be.
		{
			// if it's there - then string-ize the blob and convert it using MRMAPI 
			ULONG cbBinVal = lpsRow->lpProps[tagApptRecur].Value.bin.cb;
			LPBYTE lpBinVal = lpsRow->lpProps[tagApptRecur].Value.bin.lpb;
			BOOL bConverted = true;
		
			strApptRecurBin = Binary2String(cbBinVal, lpBinVal); //string-ize the blob as hex text
		
			if(!strApptRecurBin.IsEmpty()) // check to see if it's empty - if so then skip the conversion and throw an error out to the user.
			{			
				bConverted = ConvertApptRecurBlob(strApptRecurBin, &strRetError); // Save the blob off to ApptRecur.txt and run MRMAPI

				if(bConverted)
				{
					// Pull some preliminary data - this may be used in the future for knowing whether to pull other additional data
					bLogError = GetMainApptRecurData(&strRecurStart,
													 &ftRecurStart,
													 &strRecurEnd,
													 &ftRecurEnd,
													 &strEndType,
													 &strNumOccur,
													 &strNumExceptions,
													 &strNumOccurDeleted,
													 &strNumOccurModified,
													 &strError,
													 &strErrorNum
													);

					// If the recurrence ends in the past, and we're not checking past items, then skip it.
					ulPastRecurItem = CompareFileTime(&ftRecurEnd, &g_ftNow);
					if (ulPastRecurItem == -1)
					{
						if (!CalCheckTests.bPastItems)
						{
							goto Exit;
						}
					}

					if (bLogError) // if something was wrong in the recurrence data itself...
					{
						bLogProps = true;
						g_bMoveItem = true;
						g_ulErrors++;
						//goto Exit;
					}
					
					// do filetime checks on the start recurring time
					ulFt = TimeCheck(ftRecurStart, true);
					switch (ulFt)
					{
					case tmMin:
						{
							if (CalCheckTests.bRecurStartMin)
							{
								strError += _T("ERROR: Recurrence Start is set to zero. Check the recurrence data on this item.  ");
								strErrorNum += _T("0001 ");
							
								bLogProps = true;
								g_bMoveItem = true;
								g_ulErrors++;
								//goto Exit;
							}
							break;
						}
					case tmLow:
						{
							if (CalCheckTests.bRecurStart1995)
							{
								if(!bBdayText && !bIsAllDayEvt)
								{
									if (strError.IsEmpty())
									{
										strError += _T("WARNING: Recurrence Start Date/Time is earlier than Jan 1 1995. This might be intentional but may also indicate a problem.  ");
										strErrorNum += _T("0002 ");
										g_ulWarnings++;
									}
									if (CalCheckTests.bWarningIsError)
									{
										g_bMoveItem = true;
									}
									bLogProps = true;									
								}
							}
							break;
						}
					case tmHigh:
						{
							if (CalCheckTests.bRecurStart2025)
							{
								if(!bBdayText && !bIsAllDayEvt)
								{
									if (strError.IsEmpty())
									{
										strError += _T("WARNING: Recurrence Start Date/Time is later than Jan 1 2025. This might be intentional but may also indicate a problem.  ");
										strErrorNum += _T("0003 ");
										g_ulWarnings++;
									}
									if (CalCheckTests.bWarningIsError)
									{
										g_bMoveItem = true;
									}
									bLogProps = true;
								}
							}
							break;
						}
					case tmError:
						{
							if (CalCheckTests.bRecurStartMax)
							{
								strError += _T("ERROR: Recurrence Start Date/Time is past the upper limit. Check the recurrence data on this item.  ");
								strErrorNum += _T("0004 ");
							
								bLogProps = true;
								g_bMoveItem = true;
								g_ulErrors++;
								//goto Exit;
							}
							break;
						}
					default:
						break;
					}

					// do filetime checks on the end recurring time
					ulFt = TimeCheck(ftRecurEnd, true);
					switch (ulFt)
					{
					case tmMin:
						{
							if (CalCheckTests.bRecurEndMin)
							{
								strError += _T("ERROR: Recurrence End Date/Time is set to zero. Check the recurrence data on this item.  ");
								strErrorNum += _T("0005 ");
							
								bLogProps = true;
								g_bMoveItem = true;
								g_ulErrors++;
								//goto Exit;
							}
							break;
						}
					case tmLow:
						{
							if (CalCheckTests.bRecurEnd1995)
							{
								if (!bBdayText && !bIsAllDayEvt)
								{
									if (strError.IsEmpty())
									{
										strError += _T("WARNING: Recurrence End Date/Time is earlier than Jan 1 1995. This might be intentional but may also indicate a problem.  ");
										strErrorNum += _T("0006 ");
										g_ulWarnings++;
									}
									if (CalCheckTests.bWarningIsError)
									{
										g_bMoveItem = true;
									}
									bLogProps = true;
								}
							}
							break;
						}
					case tmHigh:
						{
							if (CalCheckTests.bRecurEnd2025)
							{
								if (!bBdayText && !bIsAllDayEvt)
								{
									if (strEndType != _T("No End Date")) // Only report this for items that are not "No End Date" items.
									{
										if (strError.IsEmpty())
										{
											strError += _T("WARNING: Recurrence End Date/Time is later than Jan 1 2025. This might be intentional but may also indicate a problem.  ");
											strErrorNum += _T("0007 ");
											g_ulWarnings++;
										}
										if (CalCheckTests.bWarningIsError)
										{
											g_bMoveItem = true;
										}
										bLogProps = true;
									}
								}
							}
							break;
						}
					case tmError:
						{
							if (CalCheckTests.bRecurEndMax)
							{
								if (strEndType != _T("No End Date")) // bypassing this since NoEndDate items might go in here when they are actually okay
								{
									strError += _T("ERROR: Recurrence End Date/Time is past the upper limit. Check the recurrence data on this item.  ");
									strErrorNum += _T("0008 ");
								
									bLogProps = true;
									g_bMoveItem = true;
									g_ulErrors++;
									//goto Exit;
								}
							}
							break;
						}
					default:
						break;
					}
				}
				else
				{
					if(g_bMRMAPI) // if MrMAPI is present, then it's not FileNotFound - it's something else
					{
						strError += strRetError;
						strErrorNum += _T("0009 "); // this could be one of three errors - all having to do with the recurrence prop not happening when calling MrMapi.
						                       //  From Logging.cpp in 
						                       //  ERROR: Could not open or write recurrence property to local file. Check that you are able to write to the folder CalCheck is located in.
						                       //  ERROR: + <error code> +  when parsing the Recurrence data.
						                       //  ERROR: Recurrence data file was not created. Check for the correct MAPI DLLs and that the user has correct permissions / context to write files here.
						
						bLogProps = true;
						g_bMoveItem = true;
						g_ulErrors++;
						//goto Exit;
					}
				}
			}
			else
			{
				strError += _T("ERROR: The Appointment Recurrence data is empty! Please check this item.  ");
				strErrorNum += _T("0010 ");
				
				bLogProps = true;
				g_bMoveItem = true;
				g_ulErrors++;
				//goto Exit;
			}

		}// if(g_bMRMAPI)

		// if this prop is there, but dispidRecurring is false, then it is in ERROR or is bad
		if (CalCheckTests.bRecurringProp)
		{
			if (ulRecurType != rectypeNone && bRecur == false)
			{
				strError += _T("ERROR: The dispidRecurring property and other RecurType properties do not agree. Please check this item.  ");
				strErrorNum += _T("0011 ");
			
				bLogProps = true;
				g_bMoveItem = true;
				g_ulErrors++;
				//goto Exit;
			}
		}

	}// if dispidApptRecur
	else
	{
		// no dispidApptRecur is set, so make sure other recur props aren't set to true...
		if (CalCheckTests.bRecurringProp)
		{
			if(bRecur)
			{
				strError += _T("ERROR: There is no Appointment Recurrence but dispidRecurring is set to True! Please check this item.  ");
				strErrorNum += _T("0012 ");
			
				bLogProps = true;
				g_bMoveItem = true;
				g_ulErrors++;
				//goto Exit;
			}
		}
	}

	// dispidApptStartWhole
	if (PT_SYSTIME == PROP_TYPE(lpsRow->lpProps[tagStart].ulPropTag))
	{
		// ftStart = lpsRow->lpProps[tagStart].Value.ft;  <<<<< This gets set at the top of this function if it's present

		// Check to see if the item falls into a "normal" timeframe
		ulFt = TimeCheck(ftStart, false);
		switch (ulFt)
		{
		case tmMin:
			{
				if (CalCheckTests.bStartTimeMin)
				{
					strError += _T("ERROR: Start Date/Time is to Zero. Please check this item.  ");
					strErrorNum += _T("0016 ");
				
					bLogProps = true;
					g_bMoveItem = true;
					g_ulErrors++;
					//goto Exit;
				}
				break;
				
			}
		case tmLow:
			{
				if (CalCheckTests.bStartTime1995)
				{
					if (!bBdayText && !bIsAllDayEvt)
					{
						if (strError.IsEmpty())
						{
							strError += _T("WARNING: Start Date/Time is earlier than Jan 1 1995. This might be intentional but may also indicate a problem.  ");
							strErrorNum += _T("0017 ");
							g_ulWarnings++;
						}
						if (CalCheckTests.bWarningIsError)
						{
							g_bMoveItem = true;
						}
						bLogProps = true;
					}
				}
				break;
			}
		case tmHigh:
			{
				if (CalCheckTests.bStartTime2025)
				{
					if (!bBdayText && !bIsAllDayEvt)
					{
						if (strError.IsEmpty())
						{
							strError += _T("WARNING: Start Date/Time is later than Jan 1 2025. This might be intentional but may also indicate a problem.  ");
							strErrorNum += _T("0018 ");
							g_ulWarnings++;
						}
						if (CalCheckTests.bWarningIsError)
						{
							g_bMoveItem = true;
						}
						bLogProps = true;
					}
				}
				break;
			}
		case tmError:
			{
				if (CalCheckTests.bStartTimeMax)
				{
					strError += _T("ERROR: Start Date/Time is past the upper limit. Please check this item.  ");
					strErrorNum += _T("0019 ");
				
					bLogProps = true;
					g_bMoveItem = true;
					g_ulErrors++;
					//goto Exit;
				}
				break;
			}
		default:
			break;
		}
	}
	else
	{
		if (CalCheckTests.bStartTimeX)
		{
			strError += _T("ERROR: Missing Appointment Start Time! Please check this item.  ");
			strErrorNum += _T("0020 ");
		
			bLogProps = true;
			g_bMoveItem = true;
			g_ulErrors++;
			//goto Exit;
		}
	}

	// dispidApptEndWhole
	if (PT_SYSTIME == PROP_TYPE(lpsRow->lpProps[tagEnd].ulPropTag))
	{
		// ftEnd = lpsRow->lpProps[tagEnd].Value.ft; <<<<< Set at the top of this function if present with Dup Check Props

		// Check to see if the item falls into a "normal" timeframe
		ulFt = TimeCheck(ftEnd, false);
		switch (ulFt)
		{
		case tmMin:
			{
				if (CalCheckTests.bEndTimeMin)
				{
					strError += _T("ERROR: End Date/Time is set to Zero. Please check this item.  ");
					strErrorNum += _T("0021 ");
				
					bLogProps = true;
					g_bMoveItem = true;
					g_ulErrors++;
					//goto Exit;
				}
				break;
			}
		case tmLow:
			{
				if (CalCheckTests.bEndTime1995)
				{
					if (!bBdayText && !bIsAllDayEvt)
					{
						if (strError.IsEmpty())
						{
							strError += _T("WARNING: End Date/Time is earlier than Jan 1 1995. This might be intentional but may also indicate a problem.  ");
							strErrorNum += _T("0022 ");
							g_ulWarnings++;
						}
						if (CalCheckTests.bWarningIsError)
						{
							g_bMoveItem = true;
						}
						bLogProps = true;
					}
				}
				break;
			}
		case tmHigh:
			{
				if (CalCheckTests.bEndTime2025)
				{
					if (!bBdayText && !bIsAllDayEvt)
					{
						if (strError.IsEmpty())
						{
							strError += _T("WARNING: End Date/Time is later than Jan 1 2025. This might be intentional but may also indicate a problem.  ");
							strErrorNum += _T("0023 ");
							g_ulWarnings++;
						}
						if (CalCheckTests.bWarningIsError)
						{
							g_bMoveItem = true;
						}
						bLogProps = true;
					}
				}
				break;
			}
		case tmError:
			{
				if (CalCheckTests.bEndTimeMax)
				{
					strError += _T("ERROR: End Date/Time is past the upper limit. Please check this item.  ");
					strErrorNum += _T("0024 ");
				
					bLogProps = true;
					g_bMoveItem = true;
					g_ulErrors++;
					//goto Exit;
				}
				break;
			}
		default:
			break;
		}
	}
	else
	{
		if (CalCheckTests.bEndTimeX)
		{
			strError += _T("ERROR: Missing Appointment End Time! Please check this item.  ");
			strErrorNum += _T("0025 ");
		
			bLogProps = true;
			g_bMoveItem = true;
			g_ulErrors++;
			//goto Exit;
		}
	}

	// dispidApptStateFlags
	if (PT_LONG == PROP_TYPE(lpsRow->lpProps[tagStateFlags].ulPropTag))
	{
		ulApptStateFlags = lpsRow->lpProps[tagStateFlags].Value.l;

		// check to see if the state flags has this mailbox as the organizer
		if (FAsfOrgdMtg(lpsRow->lpProps[tagStateFlags].Value.l))
		{
			bIsOrganizer = true;
		}
	}

	// Sent Representing Name prop not present test...
	if (!(PR_SENT_REPRESENTING_NAME_A == lpsRow->lpProps[tagSentRepresenting].ulPropTag))
	{
		if(CalCheckTests.bOrganizerName)
		{
			if(FAsfMeeting(ulApptStateFlags)) // if this is a meeting then flag this... single appts may not need this prop...
			{
				strError += _T("ERROR: Missing PR_SENT_REPRESENTING_NAME property! See KB 2825677 for additional information: https://support.microsoft.com/en-us/kb/2825677  ");
				strErrorNum += _T("0028 ");

				bLogProps = true;
				g_bMoveItem = true;
				g_ulErrors++;
				//goto Exit;
			}
		}
	}


	if (PR_SENT_REPRESENTING_ADDRTYPE_A == lpsRow->lpProps[tagSentByAddrType].ulPropTag)
	{
		strSentByAddrType = lpsRow->lpProps[tagSentByAddrType].Value.lpszA;
	}

	if (PR_SENT_REPRESENTING_EMAIL_ADDRESS_A == lpsRow->lpProps[tagSentByAddr].ulPropTag)
	{
		strSentByAddr = lpsRow->lpProps[tagSentByAddr].Value.lpszA;
	}

	if (PR_SENDER_NAME_A == lpsRow->lpProps[tagSenderName].ulPropTag)
	{
		strSenderName = lpsRow->lpProps[tagSenderName].Value.lpszA;
	}
	else
	{
		if(CalCheckTests.bSenderName)
		{
			if(FAsfMeeting(ulApptStateFlags)) // if this is a meeting then flag this... single appts may not need this prop...
			{
				strError += _T("ERROR: Missing PR_SENDER_NAME property! See KB 2825677 for additional information: https://support.microsoft.com/en-us/kb/2825677  ");
				strErrorNum += _T("0029 ");

				bLogProps = true;
				g_bMoveItem = true;
				g_ulErrors++;
				//goto Exit;
			}
		}
	}

	if (PR_SENDER_EMAIL_ADDRESS_A == lpsRow->lpProps[tagSenderAddr].ulPropTag)
	{
		strSenderAddr = lpsRow->lpProps[tagSenderAddr].Value.lpszA;		
	}

	// dispidApptAuxFlags
	if (PT_LONG == PROP_TYPE(lpsRow->lpProps[tagApptAuxFlags].ulPropTag))
	{
		ulApptAuxFlags = lpsRow->lpProps[tagApptAuxFlags].Value.l;
	}

	// Check for no Organizer Email Address
	if(strSentByAddr == "" || lpsRow->lpProps[tagSentByAddr].Value.lpszA == NULL)
	{
		if (CalCheckTests.bOrganizerAddress)
		{
			if (!(ulApptStateFlags == 0)) // a meeting requiring responses will be > 0
			{
				strError += _T("ERROR: No Organizer Address! Check the PR_SENT_REPRESENTING properties on this item.  ");
				strErrorNum += _T("0030 ");
			
				bLogProps = true;
				g_bMoveItem = true;
				g_ulErrors++;
				//goto Exit;
			}
		}
	}
	else
	{
		strSentByAddr.MakeLower();
		StringFix(&strSentByAddr);

		g_strSentRepresenting = strSentByAddr; // set the global to compare with Recipient table
	}

	// Check for no Sender Email Address
	if(strSenderAddr == "" || lpsRow->lpProps[tagSenderAddr].Value.lpszW == NULL)
	{
		if (CalCheckTests.bSenderAddress)
		{
			if (!(ulApptStateFlags == 0)) // a meeting requiring responses will be > 0
			{
				strError += _T("ERROR: No Sender Address! Check the PR_SENDER properties on this item.  ");
				strErrorNum += _T("0031 ");
			
				bLogProps = true;
				g_bMoveItem = true;
				g_ulErrors++;
				//goto Exit;
			}
		}
	}
	else
	{
		strSenderAddr.MakeLower();
	}

	// Check to see if this item has attachments on it
	if (PR_HASATTACH == lpsRow->lpProps[tagHasAttach].ulPropTag)
	{
		bHasAttach = lpsRow->lpProps[tagHasAttach].Value.b;
	}

	// PR_ENTRYID - used to open the item and get Recipient and Attachment tables
	if (PR_ENTRYID == lpsRow->lpProps[tagEID].ulPropTag)
	{
		BOOL bLogError = FALSE;
		LPBOOL lpbLogError = &bLogError;
		ulNumRecips = GetRecipData(szDN,
									lpFolder,
									lpsRow->lpProps[tagEID].Value.bin.cb,
									(LPENTRYID)lpsRow->lpProps[tagEID].Value.bin.lpb, 
									bRecur,
									lpbLogError,
									&strError,
									&strErrorNum);
		if(bLogError == TRUE)
		{
			bLogProps = true;
			g_bMoveItem = true;
			g_ulErrors++;
			//goto Exit;
		}

		if(-1 != strErrorNum.Find(_T("0058"))) // this is a warning coming back from the GetRecipData call...
		{
			bLogProps = true;
			g_ulWarnings++;
			if (CalCheckTests.bWarningIsError)
			{
				g_bMoveItem = true;
			}
		}

		if (CalCheckTests.bAttachCount)
		{
			if (bHasAttach == TRUE && bRecur == TRUE)
			{
				ulNumAttach = GetAttachCount(szDN,
											 lpFolder,
											 lpsRow->lpProps[tagEID].Value.bin.cb,
											 (LPENTRYID)lpsRow->lpProps[tagEID].Value.bin.lpb,
											 lpbLogError, 
											 &strError,
											 &strErrorNum);

				if (ulNumAttach > 25)
				{
					CString strNumAttach = _T("");
					strNumAttach.Format(_T("%i"), ulNumAttach);

					
					strError += _T("WARNING: More than 25 attachments (" + strNumAttach + ") which may indicate a problem with exceptions on this recurring meeting.  ");
					strErrorNum += _T("0037 ");
					g_ulWarnings++;
					
					if (CalCheckTests.bWarningIsError)
					{
						g_bMoveItem = true;
					}
					bLogProps = true;
				}

				if(bLogError == TRUE)
				{
					bLogProps = true;
					g_bMoveItem = true;
					g_ulErrors++;
					//goto Exit;
				}
			}
		}
	}

	// See if SentRepresenting is the same as RecipTable Organizer - if this is a recurring item
	if (CalCheckTests.bRTOrganizerIsOrganizer)
	{
		if (bRecur && g_strRecipOrganizer != "")
		{
			if (!(0 == g_strSentRepresenting.CompareNoCase(g_strRecipOrganizer)))
			{
				strError += _T("ERROR: SENT_REPRESENTING address does not match Organizer Address from the Recipient Table.  /RecipTable: " + g_strRecipOrganizer + "  /SentRepresenting: " + g_strSentRepresenting + ".  ");
				strErrorNum += _T("0041 ");

				bLogProps = true;
				g_bMoveItem = true;
				g_ulErrors++;
				//goto Exit;
			}
		}
	}

	// Test to see if this is a meeting Hijack (attendee becomes organizer)
	if(CalCheckTests.bAttendToOrganizer)
	{
		// if this is the organizer - then the Sent Representing DN should match this mailbox
		// if not - then EAS + device may have changed it (KB 2563324)
		if (bIsOrganizer)
		{
			if (!(strSentByAddr == "" || lpsRow->lpProps[tagSentByAddr].Value.lpszW == NULL)) // Checking this above... no need to double it with this test.
			{
				BOOL bEqual = false;
				CString strLegDN = g_MBXLegDN;
				strLegDN.MakeLower();
				strSentByAddr.MakeLower(); // Just in case that hasn't been done yet...

				// if the easy test works then no need to parse a list...
				if(0 == strSentByAddr.Compare(strLegDN))
				{
					bEqual = true;
				}
				else
				{
					// now check against the list of proxy addresses we have for this user
					if(!AddrList.IsEmpty())
					{
						POSITION pos;
						CString strReadAddr = _T("");
						CString strTemp = _T("");
						// BOOL bEqual = false;
						int iColon = 0;
						int iCopy = 0;
						int iStrLen = 0;

						for(pos = AddrList.GetHeadPosition(); pos != NULL; )
						{
							strTemp = AddrList.GetNext(pos);
							iStrLen = strTemp.GetLength();
							iColon = strTemp.Find(':');
							if(iColon > 0)
							{
								iCopy = iStrLen - (iColon + 1);
							}
							else
							{
								iCopy = iStrLen;
							}

							strReadAddr = strTemp.Right(iCopy);
							strReadAddr.MakeLower();
					
							if(0 == strSentByAddr.Compare(strReadAddr))
							{
								bEqual = true;
								break;
							}
							strTemp = _T("");
							strReadAddr = _T("");
						}
					}
				}
				if(bEqual == false)
				{
					strError += _T("ERROR: The organizer of this meeting is potentially incorrect. See KB 2563324: https://support.microsoft.com/en-us/kb/2563324  ");
					strErrorNum += _T("0042 ");
					
					bLogProps = true;
					g_bMoveItem = true;
					g_ulErrors++;
					//goto Exit;
				}
			}			
		}
	}

    // Check for limits on number of recurring meetings in the Calendar
	if (CalCheckTests.bRecurItemLimit)
	{
		if(g_ulRecurItems == 1299)
		{
			WriteLogFile(_T("ERROR: Number of Recurring Appointments reached the limit of 1300. To correct this please delete some older recurring appointments."), 1);
			strErrorNum += _T("0061 ");
			g_ulErrors++;
		}

		if(g_ulRecurItems == 1250)
		{
			WriteLogFile(_T("WARNING: Number of Recurring Appointments is greater than 1250. The maximum is 1300. To prevent hitting the limit please delete some older recurring appointments."), 1);
			strErrorNum += _T("0060 ");
			g_ulWarnings++;
		}
	}

	// PidLidIsRecurring
	if (PT_BOOLEAN == PROP_TYPE(lpsRow->lpProps[tagPidLidRecur].ulPropTag))
	{
		bPidLidRecur = lpsRow->lpProps[tagPidLidRecur].Value.b;
	}

	// PidLidIsException
	if (PT_BOOLEAN == PROP_TYPE(lpsRow->lpProps[tagIsException].ulPropTag))
	{
		bIsException = lpsRow->lpProps[tagIsException].Value.b;
	}

	//dispidRecurType
	if (PT_LONG == PROP_TYPE(lpsRow->lpProps[tagRecurType].ulPropTag))
	{
		ulRecurType = lpsRow->lpProps[tagRecurType].Value.l; // it's there for future test
	}

	//PR_LAST_MODIFIER_NAME
	if (PR_LAST_MODIFIER_NAME == lpsRow->lpProps[tagLastModName].ulPropTag)
	{
		strLastModName = lpsRow->lpProps[tagLastModName].Value.lpszW;
	}

	//PR_LAST_MODIFICATION_TIME
	if (PR_LAST_MODIFICATION_TIME == lpsRow->lpProps[tagLastModTime].ulPropTag)
	{
		ftLastMod = lpsRow->lpProps[tagLastModTime].Value.ft;
	}

	// Test for Empty PR_SUBJECT after getting the dispidApptEndWhole of the item - we only log if this is a recurring or future event
	if (CalCheckTests.bNoSubject)
	{
		if(strSubject.IsEmpty())
		{
			// if the item is recurring or ends after the current date/time - ONLY then do we log this 
			ulFTCompare = CompareFileTime(&g_ftNow, &ftEnd);

			if(ulFTCompare == -1 || ulFTCompare == 0 || bRecur)
			{
				if (strError.IsEmpty())
				{
					strError += _T("WARNING: No Subject on this item. You should add a Subject to this item.  ");
					strErrorNum += _T("0032 ");
					g_ulWarnings++;
				}
				if (CalCheckTests.bWarningIsError)
				{
					g_bMoveItem = true;
				}
				bLogProps = true;
			}
		}
	}

	if (PR_MESSAGE_CLASS == lpsRow->lpProps[tagMsgClass].ulPropTag)
	{
		strMsgClass = lpsRow->lpProps[tagMsgClass].Value.lpszW;
		BOOL bGoodMsgClass = false;

		// No PR_MESSAGE_CLASS prop is a problem...
		if (CalCheckTests.bMessageClass)
		{
			if(strMsgClass.IsEmpty())
			{
				strError += _T("ERROR: No Message Class! Please check this item.  ");
				strErrorNum += _T("0033 ");
			
				bLogProps = true;
				g_bMoveItem = true;
				g_ulErrors++;
				//goto Exit;
			}
			else
			{
				// Only check/log this if the item is recurring or ends after the current date/time
				ulFTCompare = CompareFileTime(&g_ftNow, &ftEnd);

				if(ulFTCompare == -1 || ulFTCompare == 0 || bRecur)
				{				
					// check to see if it's a "known" IPM.Appointment.xxx Message Class. Array of strings for this is above.
					for(ULONG i=0; i < cMsgClassChk; i++)
					{
						if(!strMsgClass.CompareNoCase(rgstrMsgClassChk[i]))
						{
							bGoodMsgClass = true;
						}
					}
					if(bGoodMsgClass == false)
					{
						if (strError.IsEmpty())
						{
							strError += _T("WARNING: Message Class (" + strMsgClass + ") is not standard for a Calendar item and may indicate a problem.  ");
							strErrorNum += _T("0034 ");
							g_ulWarnings++;
						}
						if (CalCheckTests.bWarningIsError)
						{
							g_bMoveItem = true;
						}
						bLogProps = true;
					}
				}
			}
		}
	}
	else
	{
			strError += _T("ERROR: Missing required prop PR_MESSAGE_CLASS. Please check this item.  ");
			strErrorNum += _T("0035 ");
		
			bLogProps = true;
			g_bMoveItem = true;
			g_ulErrors++;
			//goto Exit;
	}

	//PR_MESSAGE_SIZE
	if (PR_MESSAGE_SIZE == lpsRow->lpProps[tagSize].ulPropTag)
	{
		ulSize = lpsRow->lpProps[tagSize].Value.l;
		
		bool b50M = false;
		bool b25M = false;

		if (ulSize >= 52428800)
		{
			if (CalCheckTests.bItemSize50)
			{
				strError += _T("WARNING: Message Size Exceeds 50M which may indicate a problem with attachments/exceptions/properties on this item.  ");
				strErrorNum += _T("0038 ");
				if (CalCheckTests.bWarningIsError)
				{
					g_bMoveItem = true;
				}
				b50M = true;
				bLogProps = true;
				g_ulWarnings++;
			}
		}

		if (b50M == false)
		{
			if (ulSize >= 26214400) 
			{
				if (CalCheckTests.bItemSize25)
				{
					strError += _T("WARNING: Message Size Exceeds 25M which may indicate a problem with attachments/exceptions/properties on this item.  ");
					strErrorNum += _T("0039 ");
					if (CalCheckTests.bWarningIsError)
					{
						g_bMoveItem = true;
					}
					b25M = true;
					bLogProps = true;
					g_ulWarnings++;
				}
			}

			if (b25M == false)
			{
				if (ulSize >= 10485760)
				{
					if (CalCheckTests.bItemSize10)
					{
						strError += _T("WARNING: Message Size Exceeds 10M which may indicate a problem with attachments/exceptions/properties on this item.  ");
						strErrorNum += _T("0040 ");
						if (CalCheckTests.bWarningIsError)
						{
							g_bMoveItem = true;
						}
						bLogProps = true;
						g_ulWarnings++;
					}
				}
			}
		}
	}

	//PidLidGlobalObjectId
	if (PT_BINARY == PROP_TYPE(lpsRow->lpProps[tagGlobalObjId].ulPropTag))
	{
		ULONG cbBinVal = lpsRow->lpProps[tagGlobalObjId].Value.bin.cb;
		LPBYTE lpBinVal = lpsRow->lpProps[tagGlobalObjId].Value.bin.lpb;

		//string-ize the blob as hex text for comparisons
		strGlobalObjId = Binary2String(cbBinVal, lpBinVal); 
	}

	//PidLidCleanGlobalObjectId
	if (PT_BINARY == PROP_TYPE(lpsRow->lpProps[tagCleanGlobalObjId].ulPropTag))
	{
		ULONG cbBinVal = lpsRow->lpProps[tagCleanGlobalObjId].Value.bin.cb;
		LPBYTE lpBinVal = lpsRow->lpProps[tagCleanGlobalObjId].Value.bin.lpb;

		//string-ize the blob as hex text for comparisons
		strCleanGlobalObjId = Binary2String(cbBinVal, lpBinVal); 
	}

	//Check for Items with Duplicate GlobalObjectID values
	if(!(strGlobalObjId.IsEmpty()))
	{
		if(!(strCleanGlobalObjId.IsEmpty()))
		{
			if (CalCheckTests.bDupGlobalObjectIDs)
			{
				ULONG ulDup = 0;
				ulDup = CheckGlobalObjIds(strGlobalObjId, 
										  strCleanGlobalObjId, 
										  strSubject, 
										  ftStart, 
										  ftEnd, 
										  ftRecurEnd, 
										  bIsException, 
										  &strError,
										  &strErrorNum,
										  &strRptSubject,
										  &strRptStart,
										  &strRptEnd);

				if (ulDup > 0)
				{
					bLogProps = true;
					g_bMoveItem = true;
					g_ulErrors++;
					//goto Exit;
				}
			}
		}
		else if (!FAafCopied(ulApptAuxFlags)) // if this is a copied meeting then GOIDs will be removed - by design
		{
			if (CalCheckTests.bNoGlobalObjectIDs)
			{
				if (strError.IsEmpty())
				{
					strError += _T("WARNING: The dispidCleanGlobalObjectID property is not populated on this item. Please see KB 2714118: https://support.microsoft.com/en-us/kb/2714118  ");
					strErrorNum += _T("0043 ");
					g_ulWarnings++;
				}

				if (CalCheckTests.bWarningIsError)
				{
					g_bMoveItem = true;
				}
				bLogProps = true;
			}
		}
	}
	else if (!FAafCopied(ulApptAuxFlags))
	{
		if (CalCheckTests.bNoGlobalObjectIDs)
		{
			if (strError.IsEmpty())
			{
				if(strCleanGlobalObjId.IsEmpty())
				{
					strError += _T("WARNING: The dispidGlobalObjectID and the dispidCleanGlobalObjectID properties are not populated on this item. Please see KB 2714118: https://support.microsoft.com/en-us/kb/2714118  ");
					strErrorNum += _T("0044 ");
				}
				else
				{
					strError += _T("WARNING: The dispidGlobalObjectID property is not populated on this item. Please see KB 2714118: https://support.microsoft.com/en-us/kb/2714118  ");
					strErrorNum += _T("0045 ");
				}
				g_ulWarnings++;
			}

			if (CalCheckTests.bWarningIsError)
			{
				g_bMoveItem = true;
			}

			bLogProps = true;
		}
	}

Exit:

	// Check for duplicate Calendar Items
	if(CalCheckTests.bDuplicates)
	{
		// Check to see if this is a duplicate item in the Calendar
		BOOL bDuplicate = false;
		bDuplicate = CheckDuplicate(strSubject, strSentRepresenting, strLocation, bRecur, ftStart, ftEnd);

		if(bDuplicate)
		{
			strError += _T("ERROR: Duplicate item in the Calendar. Please check this item.  ");
			strErrorNum += _T("0049 ");

			bLogProps = true;
			g_bMoveItem = true;
			g_ulErrors++;
		}
	}

	// if the command switch was present, output item props to the CSV file - all items
	if(g_bCSV)
	{
		Props2CSV(
					szDN,
					strSubject,
					strLocation,
					strSentRepresenting,
					strSentByAddr,
					strSenderName,
					strSenderAddr,
					strMsgClass,
					strLastModName,
					bRecur,
					false,  // this will change when we start expanding recurrences
					bIsAllDayEvt,
					ftLastMod,
					ftStart, 
					ftEnd, 
					ulSize,
					ulNumRecips,
					bHasAttach,
					ulApptStateFlags,
					bIsOrganizer,
					strRecurStart,
					strRecurEnd,
					strEndType,
					strNumExceptions,
					strNumOccurDeleted,
					strNumOccurModified,
					strTZStruct,
					bLogProps
				   );
	}
    
Exit2:

	// Log to the LOG file if the flag has been set on a particular item
	if(bLogProps)
	{
		g_bFoundProb = true; // set global "problems were found" switch
		g_ulProbItems++; // allows reporting total #items found per mailbox

		logProps(	strError,
					strSubject,
					strLocation,
					strSentRepresenting,
					bRecur,
					ftStart, 
					ftEnd, 
					ftRecurEnd,
					strRptSubject,
					strRptStart,
					strRptEnd,
					strEntryID,
					strErrorNum
			   );
		bLogProps = false;
	}
		
	return hRes;
}

// Log item props for items that are marked as potentially causing a problem
void logProps(	CString strError,
				CString strSubject,
				CString strLocation,
				CString strSentRepresenting,
				BOOL bRecur,
				FILETIME ftStart, 
				FILETIME ftEnd,
				FILETIME ftRecurEnd,
				CString strRptSubject,
				CString strRptStart,
				CString strRptEnd,
				CString strEntryID,
				CString strErrorNum
			 )
{
	CString strStartTime = _T("");
	CString strEndTime = _T("");
	CString strRecurEnd = _T("");
	CString strIsRecurring = _T("");
	CString strPastItem = _T("");
	ULONG ulPastItem = CompareFileTime(&ftEnd, &g_ftNow);
	ULONG ulPastRecurItem = CompareFileTime(&ftRecurEnd, &g_ftNow);
	
	// format the props
	strStartTime = formatFiletime(ftStart);
	strEndTime = formatFiletime(ftEnd);
	strIsRecurring = formatBool(bRecur);
	strRecurEnd = formatFiletime(ftRecurEnd);

	StringFix(&strSubject);
	StringFix(&strLocation);
	StringFix(&strSentRepresenting);
		
	// for recurring items...
	if (bRecur)
	{
		// mark if this ended before the current date/time
		if (ulPastRecurItem == -1)
		{
			strPastItem = _T("True");
		}
		
		if(g_bEmitErrorNum)
			WriteErrorFile(strSubject + _T(",") +
					   strLocation + _T(",") +
					   strStartTime + _T(",") +
					   strRecurEnd + _T(",") +
					   strIsRecurring + _T(",") +
					   strSentRepresenting + _T(",") +
					   strPastItem + _T(",") +
					   strError + _T(",") +
					   strRptSubject  + _T(",") +
					   strRptStart  + _T(",") +
					   strRptEnd + _T(",") +
					   strEntryID + _T(",") +
					   strErrorNum + ("\r")
					  );
		else
			WriteErrorFile(strSubject + _T(",") +
					   strLocation + _T(",") +
					   strStartTime + _T(",") +
					   strRecurEnd + _T(",") +
					   strIsRecurring + _T(",") +
					   strSentRepresenting + _T(",") +
					   strPastItem + _T(",") +
					   strError + _T(",") +
					   strRptSubject  + _T(",") +
					   strRptStart  + _T(",") +
					   strRptEnd + ("\r")
					  );
	}
	else // for single instance items
	{
		// mark if this ended before the current date/time
		if (ulPastItem == -1)
		{
			strPastItem = _T("True");
		}
		if(g_bEmitErrorNum)
			WriteErrorFile(strSubject + _T(",") +
					   strLocation + _T(",") +
					   strStartTime + _T(",") +
					   strEndTime + _T(",") +
					   strIsRecurring + _T(",") +
					   strSentRepresenting + _T(",") +
					   strPastItem + _T(",") +
					   strError + _T(",") +
					   strRptSubject  + _T(",") +
					   strRptStart  + _T(",") +
					   strRptEnd + _T(",") +
					   strEntryID + _T(",") +
					   strErrorNum + ("\r")
					  );
		else
			WriteErrorFile(strSubject + _T(",") +
					   strLocation + _T(",") +
					   strStartTime + _T(",") +
					   strEndTime + _T(",") +
					   strIsRecurring + _T(",") +
					   strSentRepresenting + _T(",") +
					   strPastItem + _T(",") +
					   strError + _T(",") +
					   strRptSubject  + _T(",") +
					   strRptStart  + _T(",") +
					   strRptEnd + ("\r")
					  );
	}
}


// Put items in the CSV file - all of them
void Props2CSV(
				LPTSTR szDN,
				CString strSubject,
				CString strLocation,
				CString strSentRepresenting,
				CString strSentByAddr,
				CString strSenderName,
				CString strSenderAddr,
				CString strMsgClass,
				CString strLastModName,
				BOOL bRecur,
				BOOL bExpanded,
				BOOL bIsAllDayEvt,
				FILETIME ftLastMod,
				FILETIME ftStart, 
				FILETIME ftEnd, 
				ULONG ulSize,
				ULONG ulNumRecips,
				BOOL bHasAttach,
				ULONG ulApptStateFlags,
				BOOL bIsOrganizer,
				CString strRecurStartTime,
				CString strRecurEndTime,
				CString strEndType,
				CString strNumExceptions,
				CString strNumOccurDeleted,
				CString strNumOccurModified,
				CString strTZStruct,
				BOOL bLoggedByTool
			)
{
	CString strLastModTime = _T("");
	CString strLastModDate = _T("");
	CString strStartTime = _T("");
	CString strStartDate = _T("");
	CString strEndTime = _T("");
	CString strEndDate = _T("");
	CString strHasAttach = _T("");
	CString strIsRecurring = _T("");
	CString strIsAllDayEvt = _T("");
	CString strIsOrganizer = _T("");
	CString strApptStateFlags = _T("");
	CString strNumRecips = _T("");
	CString strMsgSize = _T("");
	CString strFlaggedByTool = _T("");
	CString strRecurStartDate = _T("");
	CString strRecurEndDate = _T("");
		
	// Format so it can be output
	strLastModTime = formatFiletime(ftLastMod);
	strStartTime = formatFiletime(ftStart);
	strEndTime = formatFiletime(ftEnd);
	strHasAttach = formatBool(bHasAttach);
	strIsRecurring = formatBool(bRecur);
	strIsAllDayEvt = formatBool(bIsAllDayEvt);
	strIsOrganizer = formatBool(bIsOrganizer);
	strApptStateFlags.Format(_T("%02X"), ulApptStateFlags);
	strNumRecips.Format(_T("%i"), ulNumRecips);
	strMsgSize.Format(_T("%i"), ulSize);
	strFlaggedByTool = formatBool(bLoggedByTool);
	
	// Separate Date-Times into Date and Time fields
	// First param needs to be combined Date-Time going in
	//
	SeparateDateTime(&strLastModTime, &strLastModDate);
	SeparateDateTime(&strStartTime, &strStartDate);
	SeparateDateTime(&strEndTime, &strEndDate);
	SeparateRecurDateTime(&strRecurStartTime, &strRecurStartDate);
	SeparateRecurDateTime(&strRecurEndTime, &strRecurEndDate);
	

	/*strSubject.Replace(_T("\""), _T(" "));
	strSubject.Replace(_T(","), _T("_"));
	strLocation.Replace(_T("\""), _T(" "));
	strLocation.Replace(_T(","), _T("_"));
	strLastModName.Replace(_T("\""), _T(" "));
	strLastModName.Replace(_T(","), _T("_"));
	strSentRepresenting.Replace(_T("\""), _T(" "));
	strSentRepresenting.Replace(_T(","), _T("_"));
	strSentByAddr.Replace(_T("\""), _T(" "));
	strSentByAddr.Replace(_T(","), _T("_"));
	strSenderName.Replace(_T("\""), _T(" "));
	strSenderName.Replace(_T(","), _T("_"));
	strSenderAddr .Replace(_T("\""), _T(" "));
	strSenderAddr .Replace(_T(","), _T("_"));*/

	// "Fix" potential problems with certain characters in strings
	// 
	StringFix(&strSubject);
	StringFix(&strLocation);
	StringFix(&strLastModName);
	StringFix(&strSentRepresenting);
	StringFix(&strSentByAddr);
	StringFix(&strSenderName);
	StringFix(&strSenderAddr);

	// Now write it to the CSV
	WriteCSVFile(strSubject + _T(", ") +
				  strLocation + _T(", ") +
		          strMsgSize + _T(", ") +
				  strStartDate + _T(", ") +
				  strStartTime + _T(", ") +
				  strEndDate + _T(", ") +
				  strEndTime + _T(", ") +
				  strMsgClass + _T(", ") +
				  strLastModName + _T(", ") +
				  strLastModDate + _T(", ") +
				  strLastModTime + _T(", ") +
				  strNumRecips + _T(", ") +
				  strIsRecurring + _T(", ") +
				  strHasAttach + _T(", ") +
				  strIsAllDayEvt + _T(", ") +
				  strSentRepresenting + _T(", ") +
				  strSentByAddr + _T(", ") +
				  strSenderName + _T(", ") +
				  strSenderAddr + _T(", ") +
				  strApptStateFlags + _T(", ") +
				  strRecurStartDate + _T(", ") +
				  strRecurEndDate + _T(", ") +
				  strEndType + _T(", ") +
				  strNumExceptions + _T(", ") +
				  strNumOccurDeleted + _T(", ") +
				  strNumOccurModified + _T(", ") +
				  strTZStruct + _T(", ") +
				  strFlaggedByTool + ("\r") // make sure new line is at the end here
		        );
}

CString Binary2String(ULONG cbBinVal, LPBYTE lpBinVal)
{

	char hexval[16] = {'0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'a', 'b', 'c', 'd', 'e', 'f'};
	char* chBinVal = new char[((cbBinVal+1)*2)];
		
	if(cbBinVal > 0)
	{		
		for(int i=0; i<=cbBinVal; i++)
		{
			chBinVal[i*2] = hexval[((lpBinVal[i] >> 4) & 0xF)];
			chBinVal[(i*2)+1] = hexval[(lpBinVal[i]) & 0x0F];
		}

		CString strOut(chBinVal, ((cbBinVal)*2));
		delete[] chBinVal;
		chBinVal = 0;
		return strOut;
	}

	return _T("");
}

// copy some data out of the arcvt.txt file - which contains the converted dispidApptRecur data
// ADD - pull the HEX time value for recur start and recur end so some comparisons can be done.
BOOL GetMainApptRecurData(CString* strRecurStart,
						  FILETIME* ftRecurStart,
						  CString* strRecurEnd,
						  FILETIME* ftRecurEnd,
						  CString* strEndType,
						  CString* strNumOccur,
						  CString* strNumExceptions,
						  CString* strNumOccurDeleted,
						  CString* strNumOccurModified,
						  CString* strError,
						  CString* strErrorNum)
{
	BOOL bLogError = false;
	BOOL bHaveOrigStartTime = false;
	int iStrLength = 0;  // total string length
	int iEquals = 0;  // place where the "=" is in the string
	int iColon = 0; // place where the ":" is in the string
	int iCopy = 0;  // number of characters to copy starting from the right side of the string
	int iChar = 0; // to see where the text is located in the string
	DWORD dwHexTime = 0;
	DWORD dwStartDate = 0;
	DWORD dwStartTime = 0;
	DWORD dwEndDate = 0;
	DWORD dwEndTime = 0;
	DWORD dwExceptStart = 0;
	DWORD dwExceptEnd = 0;
	DWORD dwStartNew = 0;
	DWORD dwEndNew = 0;
	FILETIME ftCompStart = {0};
	FILETIME ftExceptStart = {0};
	FILETIME ftCompEnd = {0};
	FILETIME ftExceptEnd = {0};
	const char* szHexConv = NULL;
	CStringA strConvert = ("");
	FILETIME ftConv = {0};
	LARGE_INTEGER liNumSec = {0};
	ULONG ulCompare = 0;

	CStdioFile ARFile;
	CString strARFile = _T("");
	CFileStatus status;
	CFileException e;
	CString strReadLine;
	CString strReturnLine = _T("");

	int iExSubj = 0;
	int iExStrSubjLen1 = 0;
	int iExStrSubjLen2 = 0;
	int iExLoc = 0;
	int iExStrLocLen1 = 0;
	int iExStrLocLen2 = 0;

	strARFile = g_szOutputPath;
	strARFile += _T("\\arcvt.txt");

	// if the file is there...
	if(ARFile.GetStatus(strARFile, status))
	{
		//open it
		if(ARFile.Open((LPCTSTR)strARFile, CFile::modeRead | CFile::typeText))
		{
			// find our data
			ARFile.SeekToBegin();
			while(ARFile.ReadString(strReadLine))
			{
				// get returned data for caller
				for(int i=0; i<cApptRecurMain; i++)
				{
					if(strReadLine.Find(rgstrApptRecurMain[i]) != -1)
					{
						// set variables to get to the correct place in the string to copy what is needed
						iStrLength = strReadLine.GetLength();
						iEquals = strReadLine.ReverseFind('=');
						iColon = strReadLine.Find(':');
						
						switch (i)
						{
						case 0: // StartDate:
							{
								// make sure this is the right Line in the file...
								iChar = strReadLine.Find('S', 0);
								if(iChar == 0)
								{
									// get starting time string
									iCopy = ((iStrLength - iEquals)-2);     // the extra "-2" removes the extra "= " from the string
									strReturnLine = strReadLine.Right(iCopy);
									*strRecurStart = strReturnLine;
									strReturnLine = _T("");
								
									//get filetime value
									iCopy = 8;
									for(int i=0; i<iCopy; i++)
									{
										strConvert.AppendChar(strReadLine.GetAt((iColon+4)+i));
									}
									dwStartDate = strtoul((LPCSTR)strConvert, 0, 16);
								
									//ftConv = RTime2FileTime(dwHexTime);
									//*ftRecurStart = ftConv;
								
									//dwHexTime = 0;
									//szHexConv = NULL;
									strConvert = ("");
									strReadLine = _T("");
								}
								iChar = 0;
								break;
							}
						case 1:  // EndDate:
							{
								// make sure this is the right Line in the file...
								iChar = strReadLine.Find('E', 0);
								if(iChar == 0)
								{
									//get ending time string
									iCopy = ((iStrLength - iEquals)-2);
									strReturnLine = strReadLine.Right(iCopy);
									*strRecurEnd = strReturnLine;
									strReturnLine = _T("");

									//get filetime value
									iCopy = 8;
									for(int i=0; i<iCopy; i++)
									{
										strConvert.AppendChar(strReadLine.GetAt((iColon+4)+i));
									}
									dwEndDate = strtoul((LPCSTR)strConvert, 0, 16);
								
									//ftConv = RTime2FileTime(dwHexTime);
									//*ftRecurEnd = ftConv;
								
									//dwHexTime = 0;
									//szHexConv = NULL;
									strConvert = ("");
									strReadLine = _T("");
								}
								iChar = 0;
								break;
							}
						case 2:  // EndType
							{
								iCopy = ((iStrLength - iEquals)-2);
								strReturnLine = strReadLine.Right(iCopy);

								if(!strReturnLine.Compare(_T("IDC_RCEV_PAT_ERB_END")))
								{
									strReturnLine = _T("End By X Date");
								}

								if(!strReturnLine.Compare(_T("IDC_RCEV_PAT_ERB_AFTERNOCCUR")))
								{
									strReturnLine = _T("End After X Occurrences");
								}

								if(!strReturnLine.Compare(_T("IDC_RCEV_PAT_ERB_NOEND")))
								{
									strReturnLine = _T("No End Date");
								}

								*strEndType = strReturnLine;
								strReturnLine = _T("");
								strReadLine = _T("");
								break;
							}
						case 3:  // ExceptionCount:
							{
								iCopy = ((iStrLength - iColon)-2);        // no "=" on this - so using the colon instead
								strReturnLine = strReadLine.Right(iCopy);
								*strNumExceptions = strReturnLine;
								strReturnLine = _T("");
								strReadLine = _T("");
								break;
							}
						case 4:  // DeletedInstanceCount:
							{
								iCopy = ((iStrLength - iEquals)-2);
								strReturnLine = strReadLine.Right(iCopy); // unless there are some deleted instances - this won't hit
								*strNumOccurDeleted = strReturnLine;
								strReturnLine = _T("");
								strReadLine = _T("");
								break;
							}
						case 5:  // ModifiedInstanceCount
							{
								iCopy = ((iStrLength - iEquals)-2);
								strReturnLine = strReadLine.Right(iCopy); // unless there are some modified instances - this won't hit
								*strNumOccurModified = strReturnLine;
								strReturnLine = _T("");
								strReadLine = _T("");
								break;
							}
						case 6:  // OccurrenceCount
							{
								iCopy = ((iStrLength - iEquals)-2);
								strReturnLine = strReadLine.Right(iCopy);
								*strNumOccur = strReturnLine;
								strReturnLine = _T("");
								strReadLine = _T("");
								break;
							}
						case 7:  // StartTimeOffset
							{
								iChar = strReadLine.Find('S', 0);
								if(iChar == 0)
								{
									iCopy = 8;
									for(int i=0; i<iCopy; i++)
									{
										strConvert.AppendChar(strReadLine.GetAt((iColon+4)+i));
									}
									dwStartTime = strtoul((LPCSTR)strConvert, 0, 16);

									strConvert = ("");
									strReadLine = _T("");
								}
								iChar = 0;
								break;
							}
						case 8:  // EndTimeOffset
							{
								iChar = strReadLine.Find('E', 0);
								if(iChar == 0)
								{
									iCopy = 8;
									for(int i=0; i<iCopy; i++)
									{
										strConvert.AppendChar(strReadLine.GetAt((iColon+4)+i));
									}
									dwEndTime = strtoul((LPCSTR)strConvert, 0, 16);

									strConvert = ("");
									strReadLine = _T("");
								}
								iChar = 0;
								break;
							}
						default:
							{break;}
						} // switch
					} // if
				} // for
			} // while

			// now set adjusted filetimes for recurring start and recurring end
			if (dwStartDate > 0)
			{				
				dwStartNew = (ULONG)dwStartDate + (ULONG)dwStartTime;
				ftConv = RTime2FileTime(dwStartNew);
				*ftRecurStart = ftConv;

				ftCompStart = RTime2FileTime(dwStartDate);
				
			}
			if (dwEndDate > 0)
			{
				ULONG ulPastRecurItem = 0;
				dwEndNew = (ULONG)dwEndDate + (ULONG)dwEndTime;
				ftConv = RTime2FileTime(dwEndNew);
				*ftRecurEnd = ftConv;

				ftCompEnd = RTime2FileTime(dwEndNew);

				// If the recurrence ends in the past, and we're not checking past items, then skip it.
				ulPastRecurItem = CompareFileTime(&ftCompEnd, &g_ftNow);
				if (ulPastRecurItem == -1)
				{
					if (!CalCheckTests.bPastItems)
					{
						goto Exit;
					}
				}
			}

			// start over to do some more checks on exception data
			ARFile.SeekToBegin();

			BOOL bExStartPrior = false;
			BOOL bExEndPrior = false;
			BOOL bExStartPost = false;
			BOOL bExEndPost = false;

			while(ARFile.ReadString(strReadLine))
			{
				// check exception info data to ensure it is right
				if(strReadLine.Find(_T("ExceptionInfo")) != -1)
				{
					int iStrEq = strReadLine.ReverseFind('=');
					int iStrCol = strReadLine.Find(':');
					int iStrLen = strReadLine.GetLength();
					CString strTemp = _T("");

					if (CalCheckTests.bExceptionData)
					{
						// Make sure the subject length matches the string length for the subject
						if(strReadLine.Find(_T("SubjectLength:")) != -1)
						{						
							iCopy = ((iStrLen - iStrEq) - 2);
							strTemp = strReadLine.Right(iCopy);
							iExStrSubjLen1 = _wtoi(strTemp);
							strTemp = _T("");
						}

						if(strReadLine.Find(_T("SubjectLength2:")) != -1)
						{
							iCopy = ((iStrLen - iStrEq) - 2);
							strTemp = strReadLine.Right(iCopy);
							iExStrSubjLen2 = _wtoi(strTemp);
							strTemp = _T("");
						}

						if(strReadLine.Find(_T("Subject:")) != -1)
						{
							iExSubj = ((iStrLen - iStrCol) - 4);						
							if (iExSubj != iExStrSubjLen2)
							{
								*strError += _T("ERROR: Recurrence Exception data mismatch. Could cause Free Busy details not to display.  ");
								*strErrorNum += _T("0013 ");
								bLogError = true;
								//goto Exit;
							}						
						}
						// Make sure the location length matches the string length for the location
						if(strReadLine.Find(_T("LocationLength:")) != -1)
						{
							iCopy = ((iStrLen - iStrEq) - 2);
							strTemp = strReadLine.Right(iCopy);
							iExStrLocLen1 = _wtoi(strTemp);
							strTemp = _T("");
						}

						if(strReadLine.Find(_T("LocationLength2:")) != -1)
						{
							iCopy = ((iStrLen - iStrEq) - 2);
							strTemp = strReadLine.Right(iCopy);
							iExStrLocLen2 = _wtoi(strTemp);
							strTemp = _T("");
						}

						if(strReadLine.Find(_T("Location:")) != -1)
						{
							iExLoc = ((iStrLen - iStrCol) - 4);												
							if (iExLoc != iExStrLocLen2)
							{
								*strError += _T("ERROR: Recurrence Exception data mismatch. Could cause Free Busy details not to display.  ");
								*strErrorNum += _T("0013 ");
								bLogError = true;
								//goto Exit;
							}						
						}
					}

					// CompareFileTime returns:  
					// -1 (1st time earlier than second), 0 (equal), 1 (1st time later than second)
					// make sure this exception happens in the time range of the meeting
					if (CalCheckTests.bExceptionBounds)
					{
						if (strReadLine.Find(_T("StartDateTime:")) != -1)
						{
							iCopy = 8;
							for(int i=0; i<iCopy; i++)
							{
								strConvert.AppendChar(strReadLine.GetAt((iStrCol+4)+i));
							}
							dwExceptStart = strtoul((LPCSTR)strConvert, 0, 16);
							ftExceptStart = RTime2FileTime(dwExceptStart);
							ulCompare = CompareFileTime(&ftExceptStart, &ftCompStart);

							if(ulCompare == -1)
							{
								// Turns out moving the first item before the series start is valid, so turning this report off since we don't seem to hit it very often
								/*if (bExStartPrior == false)
								{
									*strError = _T("ERROR: Recurrence Exception Start Time occurs before the start of the series. This could cause a hang or crash in Outlook.");
									bExStartPrior = true;
									bLogError = true;
									goto Exit;
								}*/
							}

							ulCompare = 0;
							strConvert = _T("");
							strReadLine = _T("");
						}

						// make sure this exception happens in the time range of the meeting
						if (strReadLine.Find(_T("EndDateTime:")) != -1)
						{
							iCopy = 8;
							for(int i=0; i<iCopy; i++)
							{
								strConvert.AppendChar(strReadLine.GetAt((iStrCol+4)+i));
							}
							dwExceptEnd = strtoul((LPCSTR)strConvert, 0, 16);
							ftExceptEnd = RTime2FileTime(dwExceptEnd);
							ulCompare = CompareFileTime(&ftExceptEnd, &ftCompStart);

							if(ulCompare == -1)
							{
								// Same as above with Exception start time...
								/*if (bExEndPrior == false)
								{
									*strError = _T("ERROR: Recurrence Exception End Time occurs before the start of the series. This could cause a hang or crash in Outlook.");
									bExEndPrior = true;
									bLogError = true;
									goto Exit;
								}*/
							}

							ulCompare = 0;
							strConvert = _T("");
							strReadLine = _T("");
						}

						// CompareFileTime returns:  
						// -1 (1st time earlier than second), 0 (equal), 1 (1st time later than second)
						// make sure this exception happens in the time range of the meeting
						if (strReadLine.Find(_T("OriginalStartDate:")) != -1)
						{
							iCopy = 8;
							for(int i=0; i<iCopy; i++)
							{
								strConvert.AppendChar(strReadLine.GetAt((iStrCol+4)+i));
							}
							dwExceptStart = strtoul((LPCSTR)strConvert, 0, 16);
							ftExceptStart = RTime2FileTime(dwExceptStart);
							ulCompare = CompareFileTime(&ftExceptStart, &ftCompStart);

							if(ulCompare == -1)
							{
								if (bExStartPrior == false)
								{
									*strError += _T("ERROR: Recurrence Original Start Time is set before the start of the series. This could cause a hang or crash in Outlook.  ");
									*strErrorNum += _T("0014 ");
									bExStartPrior = true;
									bLogError = true;
									//goto Exit;
								}
							}
						
							// If Original Start Date is after the end of the series - then that's WRONG
							ulCompare = CompareFileTime(&ftExceptStart, &ftCompEnd);
							if (ulCompare != -1)
							{
								if (bExStartPost == false)
								{
									*strError += _T("ERROR: Recurrence Original Start Time occurs after the end of the series. This could cause a hang or crash in Outlook.  ");
									*strErrorNum += _T("0015 ");
									bExStartPost = true;
									bLogError = true;
									//goto Exit;
								}
							}

							ulCompare = 0;
							strConvert = _T("");
							strReadLine = _T("");
						}
					}
				}
			}

		} // if

Exit:

		// close the file now that we're done with it
		ARFile.Close();
	} // if - file exists
	
	return bLogError;
}

BOOL GetPropDefStreamData(CString* strError, 
						  CString* strErrorNum)
{
	CString strNmidName = _T("");
	CString strNmidNameLength = _T("");
	BOOL bNmidProb = false; // Missing NmidName line can happen multiple times when the stream is corrupt
	BOOL bLogError = false;
	CString strReadLine;
	int iStrLength = 0;  // total string length
	int iEquals = 0;  // place where the "=" is in the string
	int iColon = 0; // place where the ":" is in the string
	int iCopy = 0;  // number of characters to copy starting from the right side of the string
	int iChar = 0; // to see where the text is located in the string

	CStdioFile RETFile;
	CString strRETFile = _T("");
	CFileStatus status;
	CFileException e;
	strRETFile = g_szOutputPath;
	strRETFile += _T("\\PDSret.txt");

	// if the file is there...
	if(RETFile.GetStatus(strRETFile, status))
	{
		//open it
		if(RETFile.Open((LPCTSTR)strRETFile, CFile::modeRead | CFile::typeText))
		{
			// find our data
			RETFile.SeekToBegin();
			while(RETFile.ReadString(strReadLine))
			{
				iStrLength = strReadLine.GetLength();
				iEquals = strReadLine.ReverseFind('=');
				iColon = strReadLine.Find(':');

				// for now we only know this problem with Mac Outlook interaction
				if(strReadLine.Find(_T("NmidNameLength =")) != -1)
				{
					iCopy = ((iStrLength - iEquals)-2);     // the extra "-2" removes the extra "= " from the string
					strNmidNameLength = strReadLine.Right(iCopy);
				}
				if(strReadLine.Find(_T("NmidName =")) != -1) // After finding NmidNameLength this one should be the next line
				{
					iCopy = ((iStrLength - iEquals)-2);     // the extra "-2" removes the extra "= " from the string
					strNmidName = strReadLine.Right(iCopy);

					if (strNmidName.IsEmpty() && (strNmidNameLength != _T("0x0000"))) // if no NmidName, and NmidNameLength isn't 0 then error
					{
						if (!bNmidProb)
						{
							*strError += _T("ERROR: The dispidPropDefStream property is corrupt. This is a known problem that can cause Outlook to crash. See: https://support.microsoft.com/en-us/kb/3017283 ");
							*strErrorNum += _T("0056 ");
							bLogError = true;
							bNmidProb = true; // only need to log this once per item...
						}
					}
					// reset these as they can appear several times in this prop stream and we want to check all of them
					strNmidNameLength = _T("");
					strNmidName = _T("");
				}
			}
		}
	}
	RETFile.Close();
	return bLogError;
}

ULONG ExpandDL(LPTSTR /*szAddress*/)
{
	return 1;
}


ULONG GetRecipData(LPTSTR szDN, LPMAPIFOLDER lpFolder, ULONG cbEID, LPENTRYID lpEID, BOOL bRecur, LPBOOL lpbLogError, CString* strError, CString* strErrorNum)
{
	HRESULT			hRes = S_OK;
	ULONG			ulObjType = 0;
	ULONG			ulRecipType = 0;
	ULONG			ulObjectType = 0;
	ULONG			ulRecipFlags = 0;
	LPMESSAGE		lpMessage = NULL;
	ULONG			ulNumRecips = 0;
	CString			strDN = szDN;
	CString			strEX = "EX";
	CString			strSMTP = "SMTP";
	CString			strOrganizer = "";
	CString			strDispName = "";
	CString			strAddress = "";
	CString			strAdrType = "";
	CString			strRecipType = "";
	CString			strRecipFlags = "";
	BOOL			bOrganizer = FALSE;
	BOOL			bBadAddr = FALSE;
	BOOL			bBadType = FALSE;
	BOOL			bDupEntry = FALSE;

	CStringList		RecipDupList;
	POSITION		pos;
	BOOL			bMatch = FALSE;
	CString			strWriteLine = "";
	CString			strReadLine = "";

	enum {ePR_ADDRTYPE,
		ePR_EMAIL_ADDRESS,
		ePR_OBJECT_TYPE,
		ePR_DISPLAY_NAME,
		ePR_RECIPIENT_TYPE,
		ePR_RECIPIENT_FLAGS,
		NUM_COLS};
	//These tags represent the message information we would like to pick up
	static SizedSPropTagArray(NUM_COLS,sptCols) = { NUM_COLS,
		PR_ADDRTYPE,
		PR_EMAIL_ADDRESS,
		PR_OBJECT_TYPE,
		PR_DISPLAY_NAME,
		PR_RECIPIENT_TYPE,
		PR_RECIPIENT_FLAGS
	};

	hRes = lpFolder->OpenEntry(
		cbEID,
		lpEID,
		NULL,
		MAPI_DEFERRED_ERRORS|MAPI_BEST_ACCESS,
		&ulObjType,
		(LPUNKNOWN*)&lpMessage);
	if (SUCCEEDED(hRes) && lpMessage)
	{
		LPMAPITABLE lpRecipTable = NULL;
		hRes = lpMessage->GetRecipientTable(NULL,&lpRecipTable);
		if (SUCCEEDED(hRes) && lpRecipTable)
		{
			hRes = lpRecipTable->SetColumns(
				(LPSPropTagArray) &sptCols,
				TBL_BATCH);

			if (SUCCEEDED(hRes))
			{

				LPSRowSet pRows = NULL;
				hRes = HrQueryAllRows(lpRecipTable, NULL, NULL, NULL, 0, &pRows);

				if (SUCCEEDED(hRes) && NULL != pRows)
				{
					ULONG i = 0;
					for (i = 0; i < pRows -> cRows; i++)
					{
						ulRecipType = 0;
						ulObjectType = 0;
						ulRecipFlags = 0;
						ULONG ulRecipFlagsCopy = 0;
						strDispName = "";
						strAddress = "";
						strAdrType = "";
						strRecipType = "";
						strRecipFlags = "";
						

						if (PR_ADDRTYPE == pRows->aRow[i].lpProps[ePR_ADDRTYPE].ulPropTag)
						{
							if (!(PT_ERROR == PROP_TYPE(pRows->aRow[i].lpProps[ePR_ADDRTYPE].ulPropTag)))
							{
								strAdrType = pRows->aRow[i].lpProps[ePR_ADDRTYPE].Value.LPSZ;
							}
							else
							{
								if (CalCheckTests.bRTAddressType)
								{
									if (bBadType == FALSE)
									{
										*strError += _T("ERROR: Bad or missing address type exists on entries in the Recipient Table for this item.  ");
										*strErrorNum += _T("0050 ");
										*lpbLogError = TRUE;
										bBadType = TRUE;
										//goto Exit;
									}
								}
							}
						}
						// Check address types
						if (!(strAdrType.IsEmpty()))
						{
							if (0 != strAdrType.Compare(strEX))
							{
								if (0 != strAdrType.Compare(strSMTP))
								{
									if (bBadType == FALSE)
									{
										*strError += _T("WARNING: Email address type not Exchange or SMTP on entries in Recipient Table. This could cause performance problems on address lookups.  ");
										*strErrorNum += _T("0058");
										bBadType = TRUE;
									}
								}
							}
							ulNumRecips += 1;
							continue;
						}
						else
						{
							// if address type is empty, that's not good either...
							strAdrType = "NONE";
							/*if (bBadType == FALSE)
							{
								// Seems this is REALLY chatty, and probably not of much use in the real world. So commenting out for now...
								WriteLogFile(_T("ERROR: Bad or missing address type exists on entries in the Recipient Table for this item."), 1);						
								*lpbLogError = TRUE;
								bBadType = TRUE;
							}*/
						}	

						if (PR_EMAIL_ADDRESS == pRows->aRow[i].lpProps[ePR_EMAIL_ADDRESS].ulPropTag)
						{
							if (!(PT_ERROR == PROP_TYPE(pRows->aRow[i].lpProps[ePR_EMAIL_ADDRESS].ulPropTag)))
							{
								strAddress = pRows->aRow[i].lpProps[ePR_EMAIL_ADDRESS].Value.LPSZ;
							}
							else
							{
								if (CalCheckTests.bRTAddress)
								{
									if (bBadAddr == FALSE)
									{
										*strError += _T("ERROR: Bad or missing email address exists on entries in the Recipient Table for this item.  ");
										*strErrorNum += _T("0051 ");
										*lpbLogError = TRUE;
										bBadAddr = TRUE;
										//goto Exit;
									}
								}
							}
						}
						// If EmailAddress is empty - that's an error
						if (strAddress.IsEmpty())
						{
							strAddress = "NONE";
							/*if (bBadAddr == FALSE)
							{
								// Seems this is REALLY chatty, and probably not of much use in the real world. So commenting out for now...
								WriteLogFile(_T("ERROR: Bad or missing email address exists on entries in the Recipient Table for this item."), 1);								
								*lpbLogError = TRUE;
								bBadAddr = TRUE;
							}*/
						}

						if (PR_DISPLAY_NAME == pRows->aRow[i].lpProps[ePR_DISPLAY_NAME].ulPropTag)
						{
							if (!(PT_ERROR == PROP_TYPE(pRows->aRow[i].lpProps[ePR_DISPLAY_NAME].ulPropTag)))
							{
								strDispName = pRows->aRow[i].lpProps[ePR_DISPLAY_NAME].Value.LPSZ;
							}
							else
							{
								if (CalCheckTests.bRTDisplayName)
								{
									*strError += _T("ERROR: Bad or missing display name entry exists in the Recipient Table on this item.  ");
									*strErrorNum += _T("0052 ");
									*lpbLogError = TRUE;
									//goto Exit;
								}
							}
						}
						// if DisplayName is empty - could cause confusion or other problems...
						if (strDispName.IsEmpty())
						{
							strDispName = "NONE";
							if (CalCheckTests.bRTDisplayName)
							{
								*strError += _T("ERROR: Bad or missing display name entry exists in the Recipient Table on this item.  ");
								*strErrorNum += _T("0052 ");
								*lpbLogError = TRUE;
								//goto Exit;
							}
						}

						// Check who's on To and CC lines...
						if (PR_RECIPIENT_TYPE == pRows->aRow[i].lpProps[ePR_RECIPIENT_TYPE].ulPropTag)
						{
							ulRecipType = pRows->aRow[i].lpProps[ePR_RECIPIENT_TYPE].Value.l;
						}
						if (MAPI_TO == ulRecipType)
						{
							strRecipType = "MAPI_TO";
						}
						else if (MAPI_CC == ulRecipType)
						{
							strRecipType = "MAPI_CC";
						}
						else if (MAPI_BCC == ulRecipType)
						{
							strRecipType = "MAPI_BCC";
						}
						else
						{
							strRecipType = "OTHER";
						}

						// Does this item have an Organizer set in the recip table?
						if (PR_RECIPIENT_FLAGS == pRows->aRow[i].lpProps[ePR_RECIPIENT_FLAGS].ulPropTag)
						{
							ulRecipFlags = pRows->aRow[i].lpProps[ePR_RECIPIENT_FLAGS].Value.l;
							ulRecipFlagsCopy = ulRecipFlags;
						}
						if ((ulRecipFlagsCopy & recipSendable) == recipSendable)
						{
							strRecipFlags += "Sendable | ";
						}
						if ((ulRecipFlagsCopy & recipOrganizer) == recipOrganizer)
						{
							bOrganizer = TRUE;
							strRecipFlags += "Organizer | ";
							if(!(strAddress.IsEmpty()))
							{
								StringFix(&strAddress);
								g_strRecipOrganizer = strAddress;
							}
							else
							{
								if (CalCheckTests.bRTOrganizerAddress)
								{
									*strError += _T("ERROR: Recipient Table Organizer has a bad or missing email address property.  ");
									*strErrorNum += _T("0053 ");
									*lpbLogError = TRUE;
									//goto Exit;
								}
							}
						}
						if ((ulRecipFlagsCopy & recipExceptionalResponse) == recipExceptionalResponse)
						{
							strRecipFlags += "ExceptResp | ";
						}
						if ((ulRecipFlagsCopy & recipExceptionalDeleted) == recipExceptionalDeleted)
						{
							strRecipFlags += "ExceptDelete | ";
						}
						if ((ulRecipFlagsCopy & recipOriginal) == recipOriginal)
						{
							strRecipFlags += "Original | ";
						}
						if (ulRecipFlagsCopy == 0)
						{
							strRecipFlags += "NONE";
						}

						// Object type stuff for counting recipients
						if (PR_OBJECT_TYPE == pRows->aRow[i].lpProps[ePR_OBJECT_TYPE].ulPropTag)
						{
							ulObjectType = pRows->aRow[i].lpProps[ePR_OBJECT_TYPE].Value.l;
						}
						if (MAPI_MAILUSER == ulObjectType)
						{
							// If szDN and szAddress are the same, it must be the organizer, so don't count it
							if (!(strDN.IsEmpty()))
							{
								if (!(strAddress.IsEmpty()))
								{
									if (0 != strDN.Compare(strAddress))
									//if (0 != _tcsicmp(strDN, strAddress))
									{
										ulNumRecips += 1;
									}
								}
							}
						}
						if (MAPI_DISTLIST == ulObjectType)
						{
							ulNumRecips += ExpandDL((LPTSTR)(LPCTSTR)strAddress);
						}

						if (CalCheckTests.bRTDuplicates) // only do all the duplicate testing work if the test is turned on
						{
							if (!(strAdrType.IsEmpty()) && !(strAddress.IsEmpty()) && !(strDispName.IsEmpty()) && !(strRecipType.IsEmpty()) && !(strRecipFlags.IsEmpty()))
							{
								strWriteLine = strDispName + (", ") +
											   strAddress + (", ") +
											   strAdrType + (", ") +
											   strRecipType +(", ") +
											   strRecipFlags;

								// do stuff here to check against the list
								if(!RecipDupList.IsEmpty())
								{
									for(pos = RecipDupList.GetHeadPosition(); pos != NULL; )
									{
										strReadLine = RecipDupList.GetNext(pos);
										if(0 == strWriteLine.Compare(strReadLine))
										{
											bMatch = TRUE;
										}
									}
								}
								// if we didn't find a match - then add this to the list
								if(!bMatch)
								{
									RecipDupList.AddHead(strWriteLine);
		
								}
								else
								{
									if (bDupEntry == FALSE)
									{
										bDupEntry = TRUE;
										*strError += _T("ERROR: Recipient Table contains duplicated entries. Free Busy lookups could be affected.  ");
										*strErrorNum += _T("0054 ");
										*lpbLogError = TRUE;
										//goto Exit;
									}
								}
							}
						}
					} //for

			Exit:

					FreeProws(pRows);
					pRows = NULL;

					if(!RecipDupList.IsEmpty())
					{
						RecipDupList.RemoveAll(); // Done with this list - so reset it.
					}
				}

				if (bOrganizer == FALSE)
				{
					if (bRecur) // seems Outlook only sets Organizer in Recip Table for RECURRING items...
					{
						// Skipping this report for right now - seems like it is fairly prevalent in my Calendar - but doesn't seem to cause issues...
						// WriteLogFile(_T("WARNING: The Recipient Table doesn't contain a mailbox set as the Organizer on this item."), 1);
						g_strRecipOrganizer = "";
						//*lpbLogError = TRUE;
					}
				}
			}
		}
		if (lpRecipTable) lpRecipTable->Release();
	}
	if (lpMessage) lpMessage->Release();

	return ulNumRecips;
}

ULONG GetAttachCount(LPTSTR szDN, LPMAPIFOLDER lpFolder, ULONG cbEID, LPENTRYID lpEID, LPBOOL lpbLogError, CString* strError, CString* strErrorNum)
{
	HRESULT			hRes = S_OK;
	ULONG			ulObjType = 0;
	LPMESSAGE		lpMessage = NULL;
	LPSRowSet		lpAttachRows = NULL;
	ULONG			ulNumAttach = 0;
	CString			strHRes = _T("");

	hRes = lpFolder->OpenEntry(cbEID,
							    lpEID,
								NULL,
								MAPI_DEFERRED_ERRORS|MAPI_BEST_ACCESS,
								&ulObjType,
								(LPUNKNOWN*)&lpMessage);

	if (SUCCEEDED(hRes) && lpMessage)
	{
		LPMAPITABLE lpAttachTable = NULL;
		hRes = lpMessage->GetAttachmentTable(NULL, &lpAttachTable);

		if (SUCCEEDED(hRes) && lpAttachTable)
		{
			hRes = HrQueryAllRows(lpAttachTable, NULL, NULL, NULL, 0, &lpAttachRows);

			if (SUCCEEDED(hRes))
			{
				ulNumAttach = lpAttachRows->cRows;
			}
			FreeProws(lpAttachRows);
			lpAttachRows = NULL;
		}

		if (lpAttachTable) lpAttachTable->Release();
		
	}

    if (!(SUCCEEDED(hRes)))
	{
		strHRes.Format(_T("%08X"), hRes);
		*strError += _T("ERROR: Unable to access the attachment table for this item. Error: " + strHRes + ".  ");
		*strErrorNum += _T("0036 ");
		*lpbLogError = TRUE;
	}

	if (lpMessage) lpMessage->Release();

	return ulNumAttach;
}


ULONG TimeCheck(FILETIME ftCheck1, BOOL bRecurCheck)
{
	// CompareFileTime returns:  
	// -1 (1st time earlier than second), 0 (equal), 1 (1st time later than second)

	ULONG ulReturn = tmGood;
	ULONG ulCompare = 0;

	ulCompare = CompareFileTime(&ftCheck1, &ftMin);
	if(ulCompare == 0)
	{
		ulReturn = tmMin;
		goto Exit;
	}

	ulCompare = CompareFileTime(&ftCheck1, &ftLow);
	if(ulCompare == 0 || ulCompare == 0xffffffff)
	{
		ulReturn = tmLow;
		goto Exit;
	}

	if (bRecurCheck == false) // check different dates for non-recurring than for recurring, since the upper limit is set on NO END DATE items
	{
		ulCompare = CompareFileTime(&ftCheck1, &ftClipEndPlus);
		if(ulCompare == 0 || ulCompare == 1)
		{
			ulReturn = tmError;
			goto Exit;
		}

		ulCompare = CompareFileTime(&ftCheck1, &ftHigh);
		if(ulCompare == 0 || ulCompare == 1)
		{
			ulReturn = tmHigh;
			goto Exit;
		}
	}
	else // do a check for recurrence times
	{
		ulCompare = CompareFileTime(&ftCheck1, &ftClipEndPlus);
		if(ulCompare == 1) // if time on item > upper limit then it's an error
		{
			ulReturn = tmError;
			goto Exit;
		}
		
		ulCompare = CompareFileTime(&ftCheck1, &ftHigh);
		if(ulCompare == 0 || ulCompare == 1)
		{
			ulReturn = tmHigh;
			goto Exit;
		}
	}

Exit:

	return ulReturn;  // 0 = Good,  1 = ftMin, 2 <= ftLow, 3 >= ftHigh, 4 >= Outlook boundary
}

// Get the folder permissions for the Calendar Folder and log them
HRESULT GetACLs(LPMAPIFOLDER lpMAPIFolder)
{
	HRESULT					hRes = S_OK;
	LPEXCHANGEMODIFYTABLE	lpExchTbl = NULL;
	LPMAPITABLE				lpMAPITable = NULL;
	LPSPropTagArray			lpPropTags = NULL;  // for all props
	LPSPropValue			pv = NULL;
	CString					strHRes = _T("");

	//Put the props that we want to retreive into the array
	lpPropTags = (LPSPropTagArray)&rgACLPropTags;


	hRes = lpMAPIFolder->OpenProperty(PR_ACL_TABLE,
									  (LPGUID)&IID_IExchangeModifyTable,
									  0,
									  MAPI_DEFERRED_ERRORS,
									  (LPUNKNOWN FAR *)&lpExchTbl);

	if (SUCCEEDED(hRes) && lpExchTbl)
	{
		hRes = lpExchTbl->GetTable(0, &lpMAPITable); // Sending 0 here since this tool doesn't care about showing the FB permission state

		if (SUCCEEDED(hRes) && lpMAPITable)
		{
			// get stuff out of the table.
			hRes = lpMAPITable->SetColumns(lpPropTags, 0);

			if (SUCCEEDED(hRes))
			{
				hRes = lpMAPITable->SeekRow(BOOKMARK_BEGINNING, 0, NULL);

				if (SUCCEEDED(hRes))
				{
					WriteLogFile(_T("Permissions on this Calendar:"), 0);
					WriteLogFile(_T("============================="), 0);
					LPSRowSet pRows = NULL;
					for (;;)
					{						
						hRes = lpMAPITable->QueryRows(50, 0, &pRows);
					
						if (FAILED(hRes))
						{
							strHRes.Format(_T("%08X"), hRes);
							WriteLogFile(_T("Could not query rows on the Calendar's ACL Table. Check connectivity and access to this user's Calendar folder."), 1);
							WriteLogFile(_T("Error: ") + strHRes, 1);
							break;
						}

						if (!pRows || !pRows->cRows) break;

						ULONG i = 0;
						for (i = 0 ; i < pRows->cRows ; i++)
						{
							// Process the rows
							ProcessACLRows(lpMAPIFolder, &pRows->aRow[i]);

						}
						FreeProws(pRows);
						pRows = NULL;
					}					
				}
				else
				{
					strHRes.Format(_T("%08X"), hRes);
					WriteLogFile(_T("Unable to access the Calendar's ACL Table. Check connectivity and access to this user's Calendar folder."), 1);
					WriteLogFile(_T("Error: ") + strHRes, 1);
				}
			}
		}
		else
		{
			strHRes.Format(_T("%08X"), hRes);
			WriteLogFile(_T("Could not open the Calendar's ACL Table. Check connectivity and access to this user's Calendar folder."), 1);
			WriteLogFile(_T("Error: ") + strHRes, 1);
			goto Exit;
		}
	}
	else
	{
		strHRes.Format(_T("%08X"), hRes);
		WriteLogFile(_T("Could not open the Calendar's ACL Table. Check connectivity and access to this user's Calendar folder."), 1);
		WriteLogFile(_T("Error: ") + strHRes, 1);
		goto Exit;
	}

	WriteLogFile(_T("=============================\r\n"), 0);


Exit:
	MAPIFreeBuffer(pv);

	if (lpExchTbl)
	{
		lpExchTbl->Release();
	}
	if (lpMAPITable)
	{
		lpMAPITable->Release();
	}

	return hRes;
}

HRESULT ProcessACLRows(LPMAPIFOLDER lpFolder, LPSRow lpsRow)
{
	HRESULT		hRes = S_OK;
	CString		strMemberName = _T("");
	CString		strRights = _T("");
	ULONG		ulRights = 0; 

	if (PR_MEMBER_NAME == lpsRow->lpProps[ePR_MEMBER_NAME].ulPropTag)
	{
		strMemberName = lpsRow->lpProps[ePR_MEMBER_NAME].Value.lpszW;
	}

	if (PR_MEMBER_RIGHTS == lpsRow->lpProps[ePR_MEMBER_RIGHTS].ulPropTag)
	{
		ulRights = lpsRow->lpProps[ePR_MEMBER_RIGHTS].Value.l;
		

		switch(ulRights)
		{
		case aclOwner:
			{
				strRights = _T("Owner");
				break;
			}
		case aclPubEditor:
			{
				strRights = _T("Publishing Editor");
				break;
			}
		case aclEditor:
			{
				strRights = _T("Editor");
				break;
			}
		case aclPubAuthor:
			{
				strRights = _T("Publishing Author");
				break;
			}
		case aclAuthor:
			{
				strRights = _T("Author");
				break;
			}
		case aclNonEditAuthor:
			{
				strRights = _T("Nonediting Author");
				break;
			}
		case aclReviewer:
			{
				strRights = _T("Reviewer");
				break;
			}
		case aclContributor:
			{
				strRights = _T("Contributor");
				break;
			}
		case aclNone:
			{
				strRights = _T("None");
				break;
			}
		case aclZero:
			{
				strRights = _T("None");
				break;
			}
		default:
			{
				strRights.Format(_T("Custom"));
				break;
			}
		}
	}

	WriteLogFile(strMemberName + _T(": ") + strRights, 0);
	

	return hRes;
}

FILETIME RTime2FileTime(DWORD rTime)
{	
	FILETIME fTime = {0};
	LARGE_INTEGER liNumSec = {0};

	liNumSec.LowPart = rTime;
	liNumSec.QuadPart = liNumSec.QuadPart*10000000*60;
	fTime.dwLowDateTime = liNumSec.LowPart;
	fTime.dwHighDateTime= liNumSec.HighPart;

	return fTime;
}

// checks Subject string
BOOL CheckSubject(CString strSubject)
{	
	CString strTemp = strSubject;
	BOOL bFound = false;
	strTemp.MakeLower();
	
	// Check for various birthday texts (eng,ger,frn,spa so far)
	for (int i = 0; i < cBDayText; i++)
	{
		if (strTemp.Find(rgstrBDayText[i], 0) != -1)
		{
			bFound = true;
		}
	}

	return bFound;
}

// Check an item against the in-memory list - if not there then add it to the list
BOOL CheckDuplicate(CString strSubject,
					CString strOrganizer,
					CString strLocation,
					BOOL bRecur,
					FILETIME ftStart,
					FILETIME ftEnd)
{
	CString	strWriteLine;
	CString strReadLine;
	CString strIsRecurring;
	CString strApptStart;
	CString strApptEnd;
	POSITION pos;
	BOOL	bMatch = false;

	strIsRecurring = formatBool(bRecur);
	strApptStart = formatFiletime(ftStart);
	strApptEnd = formatFiletime(ftEnd);
	
	StringFix(&strSubject);
	StringFix(&strOrganizer);
	StringFix(&strLocation);

	strWriteLine = strSubject + _T(", ") +
				   strOrganizer + _T(", ") +
				   strLocation + _T(", ") +
				   strIsRecurring + _T(", ") +
				   strApptStart + _T(", ") +
				   strApptEnd;
	
	// do stuff here to check against the list
	if(!DupList.IsEmpty())
	{
		for(pos = DupList.GetHeadPosition(); pos != NULL; )
		{
			strReadLine = DupList.GetNext(pos);
			if(0 == strWriteLine.Compare(strReadLine))
			{
				bMatch = true;
			}
		}
	}
		
	// if we didn't find a match - then add this to the list
	if(!bMatch)
	{
		DupList.AddHead(strWriteLine);
		
	}

	return bMatch;
}

//
// Check for items with duplicate dispdGlobalObjectID and/or dspidCleanGlobalObjectID values
//
ULONG CheckGlobalObjIds(CString strGlobalObjId, 
						CString strCleanGlobalObjId, 
						CString strSubject, 
						FILETIME ftStart, 
						FILETIME ftEnd, 
						FILETIME ftRecurEnd, 
						BOOL bIsException,
						CString* strError,
						CString* strErrorNum,
						CString* strRptSubject,
						CString* strRptStart,
						CString* strRptEnd)
{
	ULONG ulDup = 0;
	BOOL bMatch = false;
	CString	strWriteLine;
	CString strReadLine;
	CString strApptStart;
	CString strApptEnd;
	CString strRecurEnd;
	POSITION pos;
	ULONG ulPastItem = CompareFileTime(&ftEnd, &g_ftNow);
	ULONG ulPastRecurItem = CompareFileTime(&ftRecurEnd, &g_ftNow);

	strApptStart = formatFiletime(ftStart);
	strApptEnd = formatFiletime(ftEnd);
	strRecurEnd = formatFiletime(ftRecurEnd);

	StringFix(&strSubject);

	strWriteLine = strSubject + _T(",") +
				   strApptStart + _T(",") +
				   strApptEnd + _T(",") +
				   strGlobalObjId + _T(",") +
				   strCleanGlobalObjId;

	// do stuff here to check against the list
	if(!GOIDDupList.IsEmpty())
	{
		for(pos = GOIDDupList.GetHeadPosition(); pos != NULL; )
		{
			strReadLine = GOIDDupList.GetNext(pos);
			if(strReadLine.Find(strGlobalObjId) != -1)
			{
				ulDup = 1;
			}
			if(bIsException == false) // if it's an exception, then CleanGlobObjID will be the same as the Master and other exceptions
			{
				if(strReadLine.Find(strCleanGlobalObjId) != -1)
				{
					ulDup += 2;
				}
			}

			if(ulDup > 0)
			{
				bMatch = true;
				g_bFoundProb = true; // set global "problems were found" switch

				int idx = strReadLine.Find(',');
				*strRptSubject = strReadLine.Left(idx);
				strReadLine.Delete(0,idx+1);
				
				idx = strReadLine.Find(',');
				*strRptStart = strReadLine.Left(idx);
				strReadLine.Delete(0,idx+1);

				idx = strReadLine.Find(',');
				*strRptEnd = strReadLine.Left(idx);
				
				switch (ulDup)
				{
				case 1:
					{
						*strError += _T("ERROR: PidLidGlobalObjectId values match on two items.  ");
						*strErrorNum += _T("0046 ");
						break;
					}
				case 2:
					{
						*strError += _T("ERROR: PidLidCleanGlobalObjectId values match on two items.  ");
						*strErrorNum += _T("0047 ");
						break;
					}
				case 3:
					{
						*strError += _T("ERROR: PidLidGlobalObjectId and PidLidCleanGlobalObjectId values match on two items.  ");
						*strErrorNum += _T("0048 ");
						break;
					}
				default:
					{
						break;
					}					
				}
				goto Exit;
			}
		}
	}
		
	// if we didn't find a match - then add this to the list
	if(bMatch == false)
	{
		GOIDDupList.AddHead(strWriteLine);		
	}

	Exit:

	return ulDup;
}

BOOL ReadTestSetting(CString strTest)
{
	BOOL bEnabled = 1;

	return bEnabled;
}

BOOL StringFix(CString* strTemp)
{
	BOOL bFixed = false;
	CString strFix = *strTemp;

	strFix.Replace(_T("\""), _T(" "));
	strFix.Replace(_T(","), _T("_"));

	ULONG ulLen1 = 2*(strFix.GetLength());
	LPCSTR lpszTemp2 = (LPCSTR)strFix.GetBuffer(0);
	LPBYTE bEdit = (LPBYTE)lpszTemp2;
	BYTE b1 = (BYTE)*lpszTemp2;

	for (int i = 0; i < ulLen1; i++) // Check for characters that cause truncation when writing to the CSV
	{
		b1 = (BYTE)*lpszTemp2;

		if (b1 == 0x13)  // 2013(hex) is a "smart" hyphen. 
		{
			bEdit = (LPBYTE)lpszTemp2;
			*bEdit = 0x2d; // insert a regular hyphen here instead
			bEdit += 1;
			*bEdit = 0x00;
			lpszTemp2 += 1;
			bFixed = true;
		}
		if (b1 == 0x18)  // 2018(hex) is a "smart" left apostrophe. 
		{
			bEdit = (LPBYTE)lpszTemp2;
			*bEdit = 0x27; // insert a regular ansi apotrophe here instead
			bEdit += 1;
			*bEdit = 0x00;
			lpszTemp2 += 1;
			bFixed = true;
		}
		if (b1 == 0x19)  // 2019(hex) is a "smart" right apostrophe. 
		{
			bEdit = (LPBYTE)lpszTemp2;
			*bEdit = 0x27; // insert a regular ansi apostrophe here instead
			bEdit += 1;
			*bEdit = 0x00;
			lpszTemp2 += 1;
			bFixed = true;
		}
		else if (b1 == 0x1d) // 201d(hex) is a "smart" right double quote.
		{
			bEdit = (LPBYTE)lpszTemp2;
			*bEdit = 0x20; // insert a space here instead - a regular quote might cause truncation as well
			bEdit += 1;
			*bEdit = 0x00;
			lpszTemp2 += 1;
			bFixed = true;
		}
		else if (b1 == 0x1c) // 201c(hex) is a "smart" left double quote.
		{
			bEdit = (LPBYTE)lpszTemp2;
			*bEdit = 0x20; // insert a space here instead - a regular quote might cause truncation as well
			bEdit += 1;
			*bEdit = 0x00;
			lpszTemp2 += 1;
			bFixed = true;
		}
		else if (b1 == 0x1a) // 201a(hex) is a "smart" comma
		{
			bEdit = (LPBYTE)lpszTemp2;
			*bEdit = 0x5f; // insert an underline character here instead
			bEdit += 1;
			*bEdit = 0x00;
			lpszTemp2 += 1;
			bFixed = true;
		}
		else if (b1 == 0x0D) // 0D(hex) is carriage return
		{
			bEdit = (LPBYTE)lpszTemp2;
			*bEdit = 0x00; // insert a null here instead
			lpszTemp2 += 1;
			bFixed = true;
		}
		else if (b1 == 0x0A) // 0A(hex) is line feed
		{
			bEdit = (LPBYTE)lpszTemp2;
			*bEdit = 0x00; // insert a null here instead
			lpszTemp2 += 1;
			bFixed = true;
		}
		else
		{
			lpszTemp2 += 1;
		}
	}

	*strTemp = strFix.GetString();

	return bFixed;
}

BOOL SeparateDateTime(CString* strDateTime, CString* strDate)
{
	BOOL bSeparated = false;
	int iStrLength = 0;  // total string length
	int iColon = 0; // place where the ":" is in the string
	int iCopy = 0;  // number of characters to copy
	int iChar = 0; // to set a location in the string
	int iInsert = 0;
	CString strTemp1 = *strDateTime;
	CString strTempTime = "";
	CString strTempDate = "";

	iStrLength = strTemp1.GetLength();
	iColon = strTemp1.Find(':', 0);

	iChar = iColon-3;
	strTempDate = strTemp1.Left(iChar);
	*strDate = strTempDate;

	iChar = iColon-2;
	strTempTime = strTemp1.Right(iStrLength - iChar);
	iInsert = strTempTime.GetLength();
	strTempTime.Insert((iInsert-2),_T(" "));
	*strDateTime = strTempTime;	

	return bSeparated;
}

BOOL SeparateRecurDateTime(CString* strDateTime, CString* strDate)
{
	BOOL bSeparated = false;
	int iStrLength = 0;  // total string length
	int iM = 0; // place where the "M" is in the string
	int iCopy = 0;  // number of characters to copy
	int iChar = 0; // to set a location in the string
	CString strTemp1 = *strDateTime;
	CString strTempDate = "";

	iStrLength = strTemp1.GetLength();
	iM = strTemp1.ReverseFind('M');

	iChar = iM+2;
	strTempDate = strTemp1.Right(iStrLength - iChar);
	*strDate = strTempDate;

	return bSeparated;
}

// Allocates return value with new.
// clean up with delete[].
TZDEFINITION* BinToTZDEFINITION(ULONG cbDef, LPBYTE lpbDef)
{
    if (!lpbDef) return NULL;

    // Update this if parsing code is changed!
    // this checks the size up to the flags member
    if (cbDef < 2*sizeof(BYTE) + 2*sizeof(WORD)) return NULL;

    TZDEFINITION tzDef = {0};
    TZRULE* lpRules = NULL;
    LPBYTE lpPtr = lpbDef;
    WORD cchKeyName = NULL;
    WCHAR* szKeyName = NULL;
    WORD i = 0;

    BYTE bMajorVersion = *((BYTE*)lpPtr);
    lpPtr += sizeof(BYTE);
    BYTE bMinorVersion = *((BYTE*)lpPtr);
    lpPtr += sizeof(BYTE);

    // We only understand TZ_BIN_VERSION_MAJOR
    if (TZ_BIN_VERSION_MAJOR != bMajorVersion) return NULL;

    // We only understand if >= TZ_BIN_VERSION_MINOR
    if (TZ_BIN_VERSION_MINOR > bMinorVersion) return NULL;

    WORD cbHeader = *((WORD*)lpPtr);
    lpPtr += sizeof(WORD);

    tzDef.wFlags = *((WORD*)lpPtr);
    lpPtr += sizeof(WORD);

    if (TZDEFINITION_FLAG_VALID_GUID & tzDef.wFlags)
    {
        if (lpbDef + cbDef - lpPtr < sizeof(GUID)) return NULL;
        tzDef.guidTZID = *((GUID*)lpPtr);
        lpPtr += sizeof(GUID);
    }

    if (TZDEFINITION_FLAG_VALID_KEYNAME & tzDef.wFlags)
    {
        if (lpbDef + cbDef - lpPtr < sizeof(WORD)) return NULL;
        cchKeyName = *((WORD*)lpPtr);
        lpPtr += sizeof(WORD);
        if (cchKeyName)
        {
            if (lpbDef + cbDef - lpPtr < (BYTE)sizeof(WORD)*cchKeyName) return NULL;
            szKeyName = (WCHAR*)lpPtr;
            lpPtr += cchKeyName*sizeof(WORD);
        }
    }

    if (lpbDef+ cbDef - lpPtr < sizeof(WORD)) return NULL;
    tzDef.cRules = *((WORD*)lpPtr);
    lpPtr += sizeof(WORD);
    if (tzDef.cRules)
    {
        lpRules = new TZRULE[tzDef.cRules];
        if (!lpRules) return NULL;

        LPBYTE lpNextRule = lpPtr;
        BOOL bRuleOK = false;
		
        for (i = 0;i < tzDef.cRules;i++)
        {
            bRuleOK = false;
            lpPtr = lpNextRule;
			
            if (lpbDef + cbDef - lpPtr < 
                2*sizeof(BYTE) + 2*sizeof(WORD) + 3*sizeof(long) + 2*sizeof(SYSTEMTIME)) return NULL;
            bRuleOK = true;
            BYTE bRuleMajorVersion = *((BYTE*)lpPtr);
            lpPtr += sizeof(BYTE);
            BYTE bRuleMinorVersion = *((BYTE*)lpPtr);
            lpPtr += sizeof(BYTE);
			
            // We only understand TZ_BIN_VERSION_MAJOR
            if (TZ_BIN_VERSION_MAJOR != bRuleMajorVersion) return NULL;
			
            // We only understand if >= TZ_BIN_VERSION_MINOR
            if (TZ_BIN_VERSION_MINOR > bRuleMinorVersion) return NULL;
			
            WORD cbRule = *((WORD*)lpPtr);
            lpPtr += sizeof(WORD);
			
            lpNextRule = lpPtr + cbRule;
			
            lpRules[i].wFlags = *((WORD*)lpPtr);
            lpPtr += sizeof(WORD);
			
            lpRules[i].stStart = *((SYSTEMTIME*)lpPtr);
            lpPtr += sizeof(SYSTEMTIME);
			
            lpRules[i].TZReg.lBias = *((long*)lpPtr);
            lpPtr += sizeof(long);
            lpRules[i].TZReg.lStandardBias = *((long*)lpPtr);
            lpPtr += sizeof(long);
            lpRules[i].TZReg.lDaylightBias = *((long*)lpPtr);
            lpPtr += sizeof(long);
			
            lpRules[i].TZReg.stStandardDate = *((SYSTEMTIME*)lpPtr);
            lpPtr += sizeof(SYSTEMTIME);
            lpRules[i].TZReg.stDaylightDate = *((SYSTEMTIME*)lpPtr);
            lpPtr += sizeof(SYSTEMTIME);
        }
        if (!bRuleOK)
        {
            delete[] lpRules;
            return NULL;			
        }
    }
    // Now we've read everything - allocate a structure and copy it in
    size_t cbTZDef = sizeof(TZDEFINITION) +
        sizeof(WCHAR)*(cchKeyName+1) +
        sizeof(TZRULE)*tzDef.cRules;

    TZDEFINITION* ptzDef = (TZDEFINITION*) new BYTE[cbTZDef];
    
    if (ptzDef)
    {
        // Copy main struct over
        *ptzDef = tzDef;
        lpPtr = (LPBYTE) ptzDef;
        lpPtr += sizeof(TZDEFINITION);

        if (szKeyName)
        {
            ptzDef->pwszKeyName = (WCHAR*)lpPtr;
            memcpy(lpPtr,szKeyName,cchKeyName*sizeof(WCHAR));
            ptzDef->pwszKeyName[cchKeyName] = 0;
    
            lpPtr += (cchKeyName+1)*sizeof(WCHAR);
        }

        if (ptzDef -> cRules)
        {
            ptzDef -> rgRules = (TZRULE*)lpPtr;
            for (i = 0;i < ptzDef -> cRules;i++)
            {
                ptzDef -> rgRules[i] = lpRules[i];
            }
        }
    }
    delete[] lpRules;

    return ptzDef;
}