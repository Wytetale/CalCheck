
#include "stdafx.h"
#include <mapix.h>
#include <mapiutil.h>
#include <mapitags.h>
#include "defaults.h"
#include "Logging.h"


STDMETHODIMP OpenDefaultMessageStore(LPMAPISESSION lpMAPISession, ULONG ulExplicitFlags, LPMDB * lpMDB)
{
	HRESULT			hRes = S_OK;
    LPMAPITABLE		lpStoresTbl = NULL;
	
	enum {EID, NAME, NUM_COLS};
    static SizedSPropTagArray(NUM_COLS,sptCols) = {NUM_COLS, PR_ENTRYID, PR_DISPLAY_NAME};

	if (!lpMAPISession || !lpMDB)
	{
		return MAPI_E_INVALID_PARAMETER;
	}
		
    *lpMDB = NULL;
	
	//Get the table of all the message stores available
	hRes = lpMAPISession->GetMsgStoresTable(0, &lpStoresTbl);
	if (SUCCEEDED(hRes) && lpStoresTbl)
	{	
		SRestriction	sres;
		SPropValue		spv;
		LPSRowSet		pRow       = NULL;
		
		//Set up restriction for the default store
		sres.rt = RES_PROPERTY;								//Comparing a property
		sres.res.resProperty.relop = RELOP_EQ;				//Testing equality
		sres.res.resProperty.ulPropTag = PR_DEFAULT_STORE;	//Tag to compare
		sres.res.resProperty.lpProp = &spv;					//Prop tag and value to compare against
		
		spv.ulPropTag = PR_DEFAULT_STORE;	//Tag type
		spv.Value.b   = TRUE;				//Tag value
		
		//Convert the table to an array which can be stepped through
		//Only one message store should have PR_DEFAULT_STORE set to true, so only one will be returned
		hRes = HrQueryAllRows(
			lpStoresTbl,					//Table to query
			(LPSPropTagArray) &sptCols,	//Which columns to get
			&sres,						//Restriction to use
			NULL,						//No sort order
			0,							//Max number of rows (0 means no limit)
			&pRow);						//Array to return
		if (SUCCEEDED(hRes) && pRow && pRow->cRows == 1)
		{		
			LPMDB	lpTempMDB = NULL;
			//Open the first returned (default) message store
			hRes = lpMAPISession->OpenMsgStore(
				NULL,//Window handle for dialogs
				pRow->aRow[0].lpProps[EID].Value.bin.cb,//size and...
				(LPENTRYID)pRow->aRow[0].lpProps[EID].Value.bin.lpb,//value of entry to open
				NULL,//Use default interface (IMsgStore) to open store
				MAPI_BEST_ACCESS | ulExplicitFlags,//Flags
				&lpTempMDB);//Pointer to place the store in
			if (SUCCEEDED(hRes) && lpTempMDB)
			{
				//Assign the out parameter
				*lpMDB = lpTempMDB;
			}
		}

		FreeProws(pRow);
	}

	if (lpStoresTbl)
		lpStoresTbl->Release();
	
	return hRes;
}

STDMETHODIMP OpenDefaultByProp(ULONG ulPropTag, LPMDB lpMDB, ULONG ulExplicitFlags, LPMAPIFOLDER *lpFolder)
{
	HRESULT			hRes = S_OK;
	LPMAPIFOLDER	lpRoot = NULL;
	LPMAPIFOLDER	lpInbox = NULL;
	ULONG			ulObjType = 0;
	LPSPropValue	lpEIDProp = NULL;

	*lpFolder = NULL;

	hRes = lpMDB->OpenEntry(0, NULL, NULL, MAPI_BEST_ACCESS, 
		&ulObjType, (LPUNKNOWN*)&lpRoot);

	if (SUCCEEDED(hRes) && lpRoot)
	{
		// Get the entry id from the root folder
		hRes = HrGetOneProp(lpRoot, ulPropTag, &lpEIDProp);
		if (MAPI_E_NOT_FOUND == hRes)
		{
			// Ok..just to confuse things, Outlook 11 moved the props
			// to the Inbox folder. Technically that's legal, so we 
			// have to handle it. If this isn't found, try the inbox.
			hRes = OpenInbox(lpMDB, NULL, &lpInbox);
			if (SUCCEEDED(hRes) && lpInbox)
			{
				hRes = HrGetOneProp(lpInbox, ulPropTag, &lpEIDProp);
			}
		}
	}

	// Open whatever folder we got..
	if (SUCCEEDED(hRes) && lpEIDProp)
	{
		LPMAPIFOLDER	lpTemp = NULL;

		hRes = lpMDB->OpenEntry(
			lpEIDProp->Value.bin.cb,
			(LPENTRYID) lpEIDProp->Value.bin.lpb, 
			NULL, 
			MAPI_BEST_ACCESS | ulExplicitFlags, 
			&ulObjType, 
			(LPUNKNOWN*)&lpTemp);
		if (SUCCEEDED(hRes) && lpTemp)
		{
			*lpFolder = lpTemp;
		}
	}

	MAPIFreeBuffer(lpEIDProp);
	if (lpInbox) lpInbox->Release();
	if (lpRoot) lpRoot->Release();
	return hRes;
}

STDMETHODIMP OpenPropFromMDB(ULONG ulPropTag, LPMDB lpMDB, ULONG ulExplicitFlags, LPMAPIFOLDER *lpFolder)
{
	HRESULT			hRes = S_OK;
	ULONG			ulObjType = 0;
	LPSPropValue	lpEIDProp = NULL;

	*lpFolder = NULL;

	hRes = HrGetOneProp(lpMDB, ulPropTag, &lpEIDProp);

	// Open whatever folder we got..
	if (SUCCEEDED(hRes) && lpEIDProp)
	{
		LPMAPIFOLDER	lpTemp = NULL;

		hRes = lpMDB->OpenEntry(
			lpEIDProp->Value.bin.cb,
			(LPENTRYID) lpEIDProp->Value.bin.lpb, 
			NULL, 
			MAPI_BEST_ACCESS | ulExplicitFlags, 
			&ulObjType, 
			(LPUNKNOWN*)&lpTemp);
		if (SUCCEEDED(hRes) && lpTemp)
		{
			*lpFolder = lpTemp;
		}
	}

	MAPIFreeBuffer(lpEIDProp);
	return hRes;
}

STDMETHODIMP OpenDefaultFolder(ULONG ulFolder, LPMDB lpMDB, ULONG ulExplicitFlags, LPMAPIFOLDER *lpFolder)
{
	HRESULT			hRes = S_OK;

	if (!lpMDB || !lpFolder)
	{
		return MAPI_E_INVALID_PARAMETER;
	}

	*lpFolder = NULL;

	switch(ulFolder)
	{
	case DEFAULT_CALENDAR:
		hRes = OpenDefaultByProp(PR_CALENDAR_ENTRYID,lpMDB,ulExplicitFlags,lpFolder);
		break;
	case DEFAULT_CONTACTS:
		hRes = OpenDefaultByProp(PR_CONTACTS_ENTRYID,lpMDB,ulExplicitFlags,lpFolder);
		break;
	case DEFAULT_JOURNAL:
		hRes = OpenDefaultByProp(PR_JOURNAL_ENTRYID,lpMDB,ulExplicitFlags,lpFolder);
		break;
	case DEFAULT_NOTES:
		hRes = OpenDefaultByProp(PR_NOTES_ENTRYID,lpMDB,ulExplicitFlags,lpFolder);
		break;
	case DEFAULT_TASKS:
		hRes = OpenDefaultByProp(PR_TASKS_ENTRYID,lpMDB,ulExplicitFlags,lpFolder);
		break;
	case DEFAULT_REMINDERS:
		hRes = OpenDefaultByProp(PR_REMINDERS_ENTRYID,lpMDB,ulExplicitFlags,lpFolder);
		break;
	case DEFAULT_DRAFTS:
		hRes = OpenDefaultByProp(PR_DRAFTS_ENTRYID,lpMDB,ulExplicitFlags,lpFolder);
		break;
	case DEFAULT_SENTITEMS:
		hRes = OpenPropFromMDB(PR_IPM_SENTMAIL_ENTRYID,lpMDB,ulExplicitFlags,lpFolder);
		break;
	case DEFAULT_OUTBOX:
		hRes = OpenPropFromMDB(PR_IPM_OUTBOX_ENTRYID,lpMDB,ulExplicitFlags,lpFolder);
		break;
	case DEFAULT_DELETEDITEMS:
		hRes = OpenPropFromMDB(PR_IPM_WASTEBASKET_ENTRYID,lpMDB,ulExplicitFlags,lpFolder);
		break;
	case DEFAULT_FINDER:
		hRes = OpenPropFromMDB(PR_FINDER_ENTRYID,lpMDB,ulExplicitFlags,lpFolder);
		break;
	case DEFAULT_IPM_SUBTREE:
		hRes = OpenPropFromMDB(PR_IPM_SUBTREE_ENTRYID,lpMDB,ulExplicitFlags,lpFolder);
		break;
	case DEFAULT_INBOX:
		hRes = OpenInbox(lpMDB, ulExplicitFlags, lpFolder);
		break;
	default:
		hRes = MAPI_E_INVALID_PARAMETER;
	}

	return hRes;
}

STDMETHODIMP OpenInbox(LPMDB lpMDB, ULONG ulExplicitFlags, LPMAPIFOLDER *lpInboxFolder)
{
	HRESULT			hRes = S_OK;
	ULONG			cbInbox = 0;
	LPENTRYID		lpbInbox = NULL;

	if (!lpMDB || !lpInboxFolder)
		return MAPI_E_INVALID_PARAMETER;

	*lpInboxFolder = NULL;

	hRes = lpMDB->GetReceiveFolder(
			_T("IPM.Note"),	//Get default receive folder
			fMapiUnicode,	//Flags
			&cbInbox,		//Size and ...
			&lpbInbox,		//Value of the EntryID to be returned
			NULL);			//We don't care to see the class returned
	if (SUCCEEDED(hRes) && cbInbox && lpbInbox)
	{
		ULONG			ulObjType = 0;
		LPMAPIFOLDER	lpTempFolder = NULL;

		hRes = lpMDB->OpenEntry(
			cbInbox,						//Size and...
			lpbInbox,						//Value of the Inbox's EntryID
			NULL,							//We want the default interface (IMAPIFolder)
			MAPI_BEST_ACCESS | ulExplicitFlags,	//Flags
			&ulObjType,						//Object returned type
			(LPUNKNOWN *) &lpTempFolder);	//Returned folder
		if (SUCCEEDED(hRes) && lpTempFolder)
		{
			//Assign the out parameter
			*lpInboxFolder = lpTempFolder;
		}
	}

	//Always clean up your memory here!
	MAPIFreeBuffer(lpbInbox);
	return hRes;
}

STDMETHODIMP GetFirstMessage(LPMAPIFOLDER lpFolder, ULONG ulExplicitFlags, LPMESSAGE * lppMessage)
{
	HRESULT hRes = S_OK;
	LPMAPITABLE lpContents = NULL;

	enum {EID, NUM_COLS};
    static SizedSPropTagArray(NUM_COLS,sptCols) = {NUM_COLS, PR_ENTRYID};

	if (!lpFolder || !lppMessage)
		return MAPI_E_INVALID_PARAMETER;

	*lppMessage = NULL;

	hRes = lpFolder->GetContentsTable(0, &lpContents);
	if (SUCCEEDED(hRes) && lpContents)
	{
		hRes = lpContents->SetColumns((LPSPropTagArray)&sptCols, 0);
		if (SUCCEEDED(hRes))
		{
			hRes = lpContents->SeekRow(BOOKMARK_BEGINNING, 0, NULL);
			if (SUCCEEDED(hRes))
			{
				LPSRowSet pRow = NULL;
	
				hRes = lpContents->QueryRows(1, 0, &pRow);
				if (SUCCEEDED(hRes) && pRow)
				{
					if (pRow->cRows < 1)
					{
						hRes = MAPI_E_TABLE_EMPTY;
					}
					else
					{
						if (pRow->aRow[0].lpProps[EID].ulPropTag == PR_ENTRYID)
						{
							LPMESSAGE lpTempMessage = NULL;
							ULONG ulObjType = 0;

							hRes = lpFolder->OpenEntry(pRow->aRow[0].lpProps[EID].Value.bin.cb,
													   (LPENTRYID)pRow->aRow[0].lpProps[EID].Value.bin.lpb,
													   NULL,
													   ulExplicitFlags,
													   &ulObjType,
													   (LPUNKNOWN*)&lpTempMessage);
							if (SUCCEEDED(hRes) && lpTempMessage)
							{
								*lppMessage = lpTempMessage;
							}
						}
					}
				}

				FreeProws(pRow);
			}
		}
	}

	if (lpContents)lpContents->Release();

	return hRes;
}

// Gets the delegate names from the local FreeBusy message in the store
STDMETHODIMP GetDelegates(LPMDB lpMDB)
{
	HRESULT			hRes = S_OK;
	LPSPropValue	lpPropVal = NULL;
	ULONG			ulFlag = 0;
	LPMAPIFOLDER	lpRoot = NULL;
	ULONG			ulObjType = 0;
	ULONG			ulEIDs = 0;
	LPBYTE			lpBinTemp;
	ULONG			ulNumDelegates = 0;
	CString			strDelegateName = _T("");
	CString			strDelegateFlag = _T("");
	CString			strFBMonths = _T("");
	BOOL			bFreeBusyMsg = false;
	BOOL			bBossCopy = false;
	BOOL			bBossInfo = false;
	BOOL			bGotOptions = true;
	CString			strHR = _T("");
		

	// open the root folder
	hRes = lpMDB->OpenEntry(0, NULL, NULL, MAPI_BEST_ACCESS | MAPI_NO_CACHE, // MAPI_NO_CACHE gets us the Online version
		                    &ulObjType, (LPUNKNOWN*)&lpRoot);

	if (SUCCEEDED(hRes) && lpRoot)
	{
		// get the PR_FREEBUSY_ENTRYIDS prop - it has the EID for the folder and the message
		hRes = HrGetOneProp(lpRoot, PR_FREEBUSY_ENTRYIDS, &lpPropVal);

		if (SUCCEEDED(hRes) && lpPropVal)
		{
			// This is a multivalued SBinary - so need to parse through it
			ulEIDs = lpPropVal->Value.MVbin.cValues; 
			lpBinTemp = (LPBYTE)lpPropVal->Value.MVbin.lpbin;

			for(int i=0; i<ulEIDs; i++)
			{
				ULONG			cbBinVal = 0;
				LPBYTE*			lpBinVal = new LPBYTE;
				cbBinVal = *(lpBinTemp + (i*(sizeof(_SBinary))));
				lpBinVal = (LPBYTE*)(lpBinTemp + (i*(sizeof(_SBinary))+(sizeof(LPBYTE))));
				ulObjType = 0;
				LPSPropValue	lpPVDelegates = NULL;
				LPSPropValue	lpDelegateFlags = NULL;
				LPSPropValue	lpFBMonths = NULL;
				LPSPropValue	lpBossCopy = NULL;
				LPSPropValue	lpBossInfo = NULL;
				LPSPropValue	lpAutoAccept = NULL;
				LPSPropValue	lpAANoRecur = NULL;
				LPSPropValue	lpAANoOverlap = NULL;
				LPMAPIFOLDER	lpFreeBusy = NULL;

				// open the item
				hRes = lpMDB->OpenEntry(cbBinVal,
										(LPENTRYID)*lpBinVal,
										NULL,
										MAPI_BEST_ACCESS | MAPI_NO_CACHE,
										&ulObjType,
										(LPUNKNOWN*)&lpFreeBusy);

				if (SUCCEEDED(hRes) && (ulObjType == MAPI_MESSAGE))
				{
					// we found the FB Msg - set that it is there
					bFreeBusyMsg = true;
					
					WriteLogFile(_T("Successfully located and opened the local free busy message for this mailbox."), 0);

					// Free Busy Publishing option
					hRes = HrGetOneProp(lpFreeBusy, PR_FREEBUSY_COUNT_MONTHS, &lpFBMonths);
					if (SUCCEEDED(hRes) && lpFBMonths)
					{
						strFBMonths.Format(_T("%d"), lpFBMonths->Value.l);
						WriteLogFile(_T("Publishing ") + strFBMonths + _T(" month(s) of free/busy data on the server."), 0);
					}

					//Direct Booking options (since we have the FB message open - might as well...)
					hRes = HrGetOneProp(lpFreeBusy, PR_SCHDINFO_AUTO_ACCEPT_APPTS, &lpAutoAccept);
					if(SUCCEEDED(hRes) && lpAutoAccept)
					{
						if(lpAutoAccept->Value.b == 1)
						{
							WriteLogFile(_T("Resource Scheduling / Automatically accept meeting requests is enabled."), 0);

							hRes = HrGetOneProp(lpFreeBusy, PR_SCHDINFO_DISALLOW_RECURRING_APPTS, &lpAANoRecur);
							if(SUCCEEDED(hRes) && lpAANoRecur)
							{
								if(lpAANoRecur->Value.b == 1)
								{
									WriteLogFile(_T("Decline recurring meeting requests is enabled."), 0);
								}
								else
								{
									WriteLogFile(_T("Decline recurring meeting requests is not enabled."), 0);
								}
							}

							hRes = HrGetOneProp(lpFreeBusy, PR_SCHDINFO_DISALLOW_OVERLAPPING_APPTS, &lpAANoOverlap);
							if(SUCCEEDED(hRes) && lpAANoOverlap)
							{
								if(lpAANoOverlap->Value.b == 1)
								{
									WriteLogFile(_T("Decline conflicting meeting requests is enabled."), 0);
								}
								else
								{
									WriteLogFile(_T("Decline conflicting meeting requests is not enabled."), 0);
								}
							}
						}
						else
						{
							WriteLogFile(_T("Resource Scheduling / Automatically accept meeting requests is disabled."), 0);
						}
					}
					WriteLogFile(_T("====================================\r\n"), 0);

					// Now on to the actual delgates
					WriteLogFile(_T("Delegates for this mailbox:"), 0);
					WriteLogFile(_T("==========================="), 0);

					// try to get the delegate names and delegate flags props
					hRes = HrGetOneProp(lpFreeBusy, PR_SCHDINFO_DELEGATE_NAMES_W, &lpPVDelegates);
					if (SUCCEEDED(hRes) && lpPVDelegates)
					{
						hRes = HrGetOneProp(lpFreeBusy, PR_DELEGATE_FLAGS, &lpDelegateFlags);
						if(SUCCEEDED(hRes) && lpDelegateFlags)
						{
							// parse the delegate names and flags props - they are multivalued and are coupled
							ulNumDelegates = lpPVDelegates->Value.MVszA.cValues; 
													
							for(int j=0; j<ulNumDelegates; j++)
							{
								// get each name and flag and write them to the log file
								strDelegateName = lpPVDelegates->Value.MVszA.lppszA[j];
								ulFlag = lpDelegateFlags->Value.MVl.lpl[j];
								if (ulFlag == 1)
								{
									strDelegateFlag = _T("true");
								}
								else
								{
									strDelegateFlag = _T("false");
								}
								WriteLogFile(strDelegateName + _T(":  View Private Items set to ") + strDelegateFlag, 0);
								strDelegateName = _T("");
								strDelegateFlag = _T("");
							}

							WriteLogFile(_T("===========================\r\n"), 0);
						}
						else
						{
							//delegates were found, but not the delegate flags - that's not good...
							WriteLogFile(_T("ERROR: Delegates were found, but the Delegate Flags property was not found!"), 0);

							// since we have the delegate names prop - go ahead and print those out
							ulNumDelegates = lpPVDelegates->Value.MVszA.cValues; 
																			
							for(int j=0; j<ulNumDelegates; j++)
							{
								// get each name and write it to the log file
								strDelegateName = lpPVDelegates->Value.MVszA.lppszA[j];
								WriteLogFile(strDelegateName, 0);
								strDelegateName = _T("");
							}

							WriteLogFile(_T("===========================\r\n"), 0);
						}

						// For Meeting Request and Response Option	
						WriteLogFile(_T("Delegate Meeting Request/Response Option:"), 0);
						WriteLogFile(_T("========================================="), 0);

						hRes = HrGetOneProp(lpFreeBusy, PR_SCHDINFO_BOSS_WANTS_COPY, &lpBossCopy);
						if(SUCCEEDED(hRes) && lpBossCopy)
						{
							bBossCopy = lpBossCopy->Value.b;
						}
						else
						{
							WriteLogFile(_T("ERROR: Missing the PR_SCHDINFO_BOSS_WANTS_COPY property!"), 0);
							bGotOptions = false;
						}

						hRes = HrGetOneProp(lpFreeBusy, PR_SCHDINFO_BOSS_WANTS_INFO, &lpBossInfo);
						if(SUCCEEDED(hRes) && lpBossInfo)
						{
							bBossInfo = lpBossInfo->Value.b;
						}
						else
						{
							WriteLogFile(_T("ERROR: Missing the PR_SCHDINFO_BOSS_WANTS_INFO property!"), 0);
							bGotOptions = false;
						}

						if(bGotOptions)
						{
							// Now write out delegate option...
							if(bBossCopy && bBossInfo) 
								WriteLogFile(_T("Deliver meeting requests & responses to delegates only, but send a copy of meeting requests and responses to me."), 0);

							if(!bBossCopy && bBossInfo) 
								WriteLogFile(_T("Deliver meeting requests & responses to delegates only."), 0);

							if(bBossCopy && !bBossInfo) 
								WriteLogFile(_T("Deliver meeting requests & responses to delegates and me."), 0);

						}
						
						WriteLogFile(_T("=========================================\r\n"), 0);

					}
					else
					{
						// the message was opened, but there were no delegates found on it
						WriteLogFile(_T("No delegates are set."), 0);
						WriteLogFile(_T("===========================\r\n"), 0);
					}
				}


				MAPIFreeBuffer(lpPVDelegates);
				MAPIFreeBuffer(lpDelegateFlags);
				MAPIFreeBuffer(lpFBMonths);
				MAPIFreeBuffer(lpBossCopy);
				MAPIFreeBuffer(lpBossInfo);
				MAPIFreeBuffer(lpAutoAccept);
				MAPIFreeBuffer(lpAANoRecur);
				MAPIFreeBuffer(lpAANoOverlap);

				if (lpFreeBusy)
				{
					lpFreeBusy->Release();
				}

				/*if (lpBinVal && SUCCEEDED(hRes))
				{
					delete (LPBYTE)lpBinVal;
					lpBinVal = NULL;
				}*/
			}
			if (!bFreeBusyMsg)
			{
				WriteLogFile(_T("ERROR: The local free busy message was not found!"), 0);
			}
		}
		else
		{
			strHR.Format(_T("%08X"), hRes);
			WriteLogFile(_T("ERROR: Did not find PR_FREEBUSY_ENTRYIDS on the root folder. Error: " + strHR), 0);
		}
	}
	else
	{
		strHR.Format(_T("%08X"), hRes);
		WriteLogFile(_T("ERROR: Couldn't access the root folder. Error: " + strHR), 0);
	}


	// free stuff we don't need anymore
	MAPIFreeBuffer(lpPropVal);

	if (lpRoot)
	{
		lpRoot->Release();
	}

	/*if (lpBinTemp && SUCCEEDED(hRes))
	{
		delete (LPBYTE)lpBinTemp;
		lpBinTemp = NULL;
	}*/
	
	return hRes;
}

STDMETHODIMP CreateCalCheckFolder(LPMDB lpMDB, LPMAPIFOLDER *lpCalCheckFld)
{
	HRESULT hRes = S_OK;
	LPMAPIFOLDER lpRootFolder;

	hRes = OpenDefaultFolder(DEFAULT_IPM_SUBTREE, lpMDB, NULL, &lpRootFolder);

	if (SUCCEEDED(hRes) && lpRootFolder)
	{
		hRes = lpRootFolder->CreateFolder(FOLDER_GENERIC, 
			                              _T("CalCheck"), 
			                              _T("Folder Created By CalCheck Utility"), 
								          NULL, 
								          MAPI_UNICODE | OPEN_IF_EXISTS, 
								          lpCalCheckFld);
	}

	lpRootFolder->Release();

	return hRes;
}


// below enums are for prop arrays for the message and attachment objects
// used in CreateReportMsg

enum {
		ptagSubject,
		ptagMsgClass,
		ptagBody,
		ptagDeliverTime,
		ptagMsgFlags,
		ptagConvTopic,
		NUM_MSG_PROPS
	 };

enum {
		ptagAttachMethod,
		ptagAttachLongName,
		ptagAttachName,
		ptagRenderPos,
		NUM_ATTACH_PROPS
	 };

STDMETHODIMP CreateReportMsg(LPMDB lpMDB)
{
	HRESULT			hRes = S_OK;
	ULONG			cbInbox = 0;
	LPENTRYID		lpbInbox = NULL;
	LPMESSAGE		lpTempMsg = 0;
	IAttach*		pAttach = NULL;
	ULONG			ulAttachNum = 0;
	CFile			file;
	CString			strFileName;
	SYSTEMTIME		st;
	FILETIME		ft;
	CString			strSubject;
	CString			strBodyText;
	SYSTEMTIME		stTime;
	CString			strLocalTime;

	// Need current GMT so we can place the message at the time the test was done
	GetSystemTime(&st);
	SystemTimeToFileTime(&st, &ft);
	
	// for attaching the log to the message
	strFileName = g_szAppPath;
	strFileName += _T("\\CalCheck.log");


	// Create the subject with the current time
	GetLocalTime(&stTime);
	strLocalTime.Format(_T("%02d/%02d/%d %02d:%02d:%02d%s"), 
							stTime.wMonth, 
							stTime.wDay, 
							stTime.wYear, 
							(stTime.wHour <= 12)?stTime.wHour:stTime.wHour-12,  // adjust to non-military time
							stTime.wMinute, 
							stTime.wSecond,
							(stTime.wHour <= 12)?_T("AM"):_T("PM"));


	strSubject = _T("CalCheck Report:  ") + strLocalTime;


	// Now set the body text
	strBodyText = _T("The Calendar Checking Tool has been run against this mailbox. \r\n");
	strBodyText += _T("Open the attached CalCheck.log file to see the results of these tests. \r\n\n");

	if(g_bMultiMode) // if multi MBX mode - then give additional instruction
	{
		strBodyText += _T("For any questions or assistance with resolving any reported issues, contact your Exchange System Administrator or Help Desk. \r\n\n");
	}

	strBodyText +=_T("Privacy Note:\n");
	strBodyText +=_T("The data files produced by CalCheck can contain PII such as e-mail addresses.\n"); 
	strBodyText +=_T("Please delete any files produced by this tool from your system after analysis and/or supplying them to Microsoft for analysis.\n");
	strBodyText +=_T("For more information on Microsoft's privacy standards and practices, please go to http://privacy.microsoft.com/en-us/default.mspx. \n");
	strBodyText +=_T("\n");
	
	// Now go open the Inbox
	hRes = lpMDB->GetReceiveFolder(	_T("IPM.Note"),	//Get default receive folder
									fMapiUnicode,	//Flags
									&cbInbox,		//Size and ...
									&lpbInbox,		//Value of the EntryID to be returned
									NULL
								  );

	if (SUCCEEDED(hRes) && cbInbox && lpbInbox)
	{
		ULONG			ulObjType = 0;
		LPMAPIFOLDER	lpInboxFolder = NULL;

		hRes = lpMDB->OpenEntry(cbInbox,						//Size and...
								lpbInbox,						//Value of the Inbox's EntryID
								NULL,							//We want the default interface (IMAPIFolder)
								MAPI_BEST_ACCESS,			    //Flags
								&ulObjType,						//Object returned type
								(LPUNKNOWN *) &lpInboxFolder	//Returned folder
							   );

		if (SUCCEEDED(hRes) && lpInboxFolder)
		{
			// Make the message object
			hRes = lpInboxFolder->CreateMessage(0,				// IID
												0,				// Flags
												&lpTempMsg		// returned message
											   );

			if (SUCCEEDED(hRes) && lpTempMsg)
			{
				// set props on the message
				SPropValue spvProps[NUM_MSG_PROPS] = {0};
				spvProps[ptagSubject].ulPropTag			= PR_SUBJECT_W;
				spvProps[ptagMsgClass].ulPropTag		= PR_MESSAGE_CLASS_W;
				spvProps[ptagBody].ulPropTag			= PR_BODY_W;
				spvProps[ptagDeliverTime].ulPropTag		= PR_MESSAGE_DELIVERY_TIME;
				spvProps[ptagMsgFlags].ulPropTag		= PR_MESSAGE_FLAGS;
				spvProps[ptagConvTopic].ulPropTag		= PR_CONVERSATION_TOPIC;
				
				spvProps[ptagSubject].Value.lpszW		= (LPWSTR)(LPCTSTR)strSubject;
				spvProps[ptagMsgClass].Value.lpszW		= _T("IPM.Note");
				spvProps[ptagBody].Value.lpszW			= (LPWSTR)(LPCTSTR)strBodyText;
				spvProps[ptagDeliverTime].Value.ft		= ft; 
				spvProps[ptagMsgFlags].Value.l			= MSGFLAG_UNMODIFIED;
				spvProps[ptagConvTopic].Value.lpszW		= (LPWSTR)(LPCTSTR)strSubject;

				hRes = lpTempMsg->SetProps(NUM_MSG_PROPS, spvProps, NULL);

				if (SUCCEEDED(hRes))
				{
					// now add the attachment
					if (!file.Open(strFileName, CFile::modeRead))
					{
						MessageBoxEx(NULL,_T("Could not attach log file."),_T("CalCheck"), MB_ICONERROR | MB_OK, 0);
					}
					
					hRes = lpTempMsg->CreateAttach(NULL, 0, &ulAttachNum, &pAttach);

					if (SUCCEEDED(hRes) && pAttach)
					{
						// set props on the attachment
						SPropValue spvAttach[NUM_ATTACH_PROPS] = {0};
						spvAttach[ptagAttachMethod].ulPropTag		= PR_ATTACH_METHOD;
						spvAttach[ptagAttachLongName].ulPropTag		= PR_ATTACH_LONG_FILENAME;
						spvAttach[ptagAttachName].ulPropTag			= PR_ATTACH_FILENAME;
						spvAttach[ptagRenderPos].ulPropTag			= PR_RENDERING_POSITION;

						spvAttach[ptagAttachMethod].Value.l			= ATTACH_BY_VALUE;
						spvAttach[ptagAttachLongName].Value.lpszW	= (LPWSTR)(LPCTSTR)strFileName;
						spvAttach[ptagAttachName].Value.lpszW		= (LPWSTR)(LPCTSTR)strFileName;
						spvAttach[ptagRenderPos].Value.l			= -1;

						hRes = pAttach->SetProps(NUM_ATTACH_PROPS, spvAttach, NULL);

						if (SUCCEEDED(hRes))
						{
							LPSTREAM pStream = NULL;

							hRes = pAttach->OpenProperty(PR_ATTACH_DATA_BIN, &IID_IStream, 0, MAPI_MODIFY | MAPI_CREATE, (LPUNKNOWN*)&pStream);

							if (SUCCEEDED(hRes))
							{
								const int BUF_SIZE=4096;
								BYTE pData[BUF_SIZE];
								ULONG ulSize=0;
								ULONG ulRead;
								ULONG ulWritten;

								ulRead=file.Read(pData,BUF_SIZE);
								while(ulRead) 
								{
									pStream->Write(pData,ulRead,&ulWritten);
									ulSize+=ulRead;
									ulRead=file.Read(pData,BUF_SIZE);
								}
								pStream->Commit(STGC_DEFAULT);
								MAPIFreeBuffer(pStream);
								file.Close();

								pAttach->SaveChanges(KEEP_OPEN_READONLY);
								MAPIFreeBuffer(pAttach);

							}
						}
					}

					hRes = lpTempMsg->SaveChanges(FORCE_SAVE);
				}
				MAPIFreeBuffer(lpTempMsg);
			}
		}
		MAPIFreeBuffer(lpbInbox);
	}
	return hRes;
}




