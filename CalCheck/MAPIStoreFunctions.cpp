// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
// Copyright 2006 Microsoft Corporation.  All Rights Reserved.
//
// MAPIStorefunctions.cpp : Collection of useful MAPI Store functions
//
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "MAPIStoreFunctions.h"

//Build a server DN. Allocates memory. Free with MAPIFreeBuffer.
HRESULT BuildServerDN(
					  LPCTSTR szServerName,
					  LPCTSTR szPost,
					  LPTSTR* lpszServerDN)
{
	HRESULT hRes = S_OK;
	if (!lpszServerDN) return MAPI_E_INVALID_PARAMETER;

	static LPTSTR szPre = _T("/cn=Configuration/cn=Servers/cn=");
	size_t cbPreLen = 0;
	size_t cbServerLen = 0;
	size_t cbPostLen = 0;
	size_t cbServerDN = 0;
	
	WC_H(StringCbLength(szPre,STRSAFE_MAX_CCH * sizeof(TCHAR),&cbPreLen));
	WC_H(StringCbLength(szServerName,STRSAFE_MAX_CCH * sizeof(TCHAR),&cbServerLen));
	WC_H(StringCbLength(szPost,STRSAFE_MAX_CCH * sizeof(TCHAR),&cbPostLen));
	
	cbServerDN = cbPreLen + cbServerLen + cbPostLen + sizeof(TCHAR);
	
	WC_H(MAPIAllocateBuffer(
		(ULONG) cbServerDN,
		(LPVOID*)lpszServerDN));
	
	WC_H(StringCbPrintf(
		*lpszServerDN,
		cbServerDN,
		_T("%s%s%s"),
		szPre,
		szServerName,
		szPost));
	return hRes;
}

HRESULT GetMailboxTable1(
						LPMDB lpMDB,
						LPCTSTR szServerDN,
						ULONG ulFlags,
						LPMAPITABLE* lpMailboxTable)
{
	if (!lpMDB || !lpMailboxTable || !szServerDN) return MAPI_E_INVALID_PARAMETER;
	*lpMailboxTable = NULL;
	
	HRESULT	hRes = S_OK;
	LPEXCHANGEMANAGESTORE lpManageStore1 = NULL;
	
	WC_H(lpMDB->QueryInterface(
		IID_IExchangeManageStore,
		(void **) &lpManageStore1));
	
	if (lpManageStore1)
	{
		WC_H(lpManageStore1->GetMailboxTable(
			(LPSTR) szServerDN,
			lpMailboxTable,
			ulFlags));
		
		lpManageStore1->Release();
	}
	return hRes;
} // GetMailboxTable1

HRESULT GetMailboxTable3(
						LPMDB lpMDB,
						LPCTSTR szServerDN,
						ULONG ulOffset,
						ULONG ulFlags,
						LPMAPITABLE* lpMailboxTable)
{
	if (!lpMDB || !lpMailboxTable || !szServerDN) return MAPI_E_INVALID_PARAMETER;
	*lpMailboxTable = NULL;

	DebugPrint(DBGVerbose,_T("GetMailboxTable3: asking for offset %d\n"),ulOffset);
	
	HRESULT	hRes = S_OK;
	LPEXCHANGEMANAGESTORE3 lpManageStore3 = NULL;
	
	WC_H(lpMDB->QueryInterface(
		IID_IExchangeManageStore3,
		(void **) &lpManageStore3));
	
	if (lpManageStore3)
	{
		WC_H(lpManageStore3->GetMailboxTableOffset(
			(LPSTR) szServerDN,
			lpMailboxTable,
			ulFlags,
			ulOffset));
		
		lpManageStore3->Release();
	}
	return hRes;
} // GetMailboxTable3

//lpMDB needs to be an Exchange MDB - OpenPrivateMessageStore can get one if there's one to be had
// Will try IID_IExchangeManageStore3 first and fail back to IID_IExchangeManageStore
HRESULT GetMailboxTable(
						LPMDB lpMDB,
						LPCTSTR szServerName,
						ULONG ulOffset,
						LPMAPITABLE* lpMailboxTable)
{
	if (!lpMDB || !lpMailboxTable) return MAPI_E_INVALID_PARAMETER;
	*lpMailboxTable = NULL;

	HRESULT	hRes = S_OK;
	LPTSTR	szServerDN = NULL;
	LPMAPITABLE lpLocalTable = NULL;

	WC_H(BuildServerDN(
		szServerName,
		_T(""),
		&szServerDN));
	if (szServerDN)
	{
		WC_H(GetMailboxTable3(
			lpMDB,
			szServerDN,
			ulOffset,
			fMapiUnicode,
			&lpLocalTable));

		// If we asked for an offset, we're not gonna get it - fail
		if (!lpLocalTable && !ulOffset)
		{
			WC_H(GetMailboxTable1(
				lpMDB,
				szServerDN,
				fMapiUnicode,
				&lpLocalTable));
		}
	}

	*lpMailboxTable = lpLocalTable;
	MAPIFreeBuffer(szServerDN);
	return hRes;
}//GetMailboxTable

// Delete with delete[]
int AnsiToUnicode(LPCSTR pszA, LPWSTR* ppszW)
{
	if (!ppszW) return MAPI_E_INVALID_PARAMETER;
	*ppszW = NULL;
	if (NULL == pszA) return S_OK;

	// Get our buffer size
	int iRet = 0;
	iRet = MultiByteToWideChar(
		CP_ACP, 
		0, 
		pszA, 
		(int) -1, 
		NULL,
		NULL);

	if (0 != iRet)
	{
		// MultiByteToWideChar returns num of chars
		LPWSTR pszW = new WCHAR[iRet];

		iRet = MultiByteToWideChar(
			CP_ACP, 
			0, 
			pszA, 
			(int) -1, 
			pszW,
			iRet);
		if (0 != iRet)
		{
			*ppszW = pszW;
		}
	}
	return iRet;
}

// let WideCharToMultiByte compute the length
// Delete with delete[]
int UnicodeToAnsi(LPCWSTR pszW, LPSTR* ppszA)
{
	if (!ppszA) return 0;
	*ppszA = NULL;
	if (NULL == pszW) return 0;

	// Get our buffer size
	int iRet = 0;
	iRet = WideCharToMultiByte(
		CP_ACP, 
		0, 
		pszW, 
		(int) -1, 
		NULL,
		NULL, 
		NULL, 
		NULL);

	if (0 != iRet)
	{
		// WideCharToMultiByte returns num of bytes
		LPSTR pszA = (LPSTR) new BYTE[iRet];

		iRet = WideCharToMultiByte(
			CP_ACP, 
			0, 
			pszW, 
			(int) -1, 
			pszA,
			iRet, 
			NULL, 
			NULL);
		if (0 != iRet)
		{
			*ppszA = pszA;
		}
	}
	return iRet;
;
}

//Stolen from MBLogon.c in the EDK to avoid compiling and linking in the entire beast
//Cleaned up to fit in with other functions
//$--HrMailboxLogon------------------------------------------------------
// Logon to a mailbox.  Before calling this function do the following:
//  1) Create a profile that has Exchange administrator privileges.
//  2) Logon to Exchange using this profile.
//  3) Open the mailbox using the Message Store DN and Mailbox DN.
//
// This version of the function needs the server and mailbox names to be
// in the form of distinguished names.  They would look something like this:
//		/CN=Configuration/CN=Servers/CN=%s/CN=Microsoft Private MDB
//		/CN=Configuration/CN=Servers/CN=%s/CN=Microsoft Public MDB
//		/O=<Organization>/OU=<Site>/CN=<Container>/CN=<MailboxName>
// where items in <brackets> would need to be set to appropriate values
// 
// Note1: The message store DN is nearly identical to the server DN, except
// for the addition of a trailing '/CN=' part.  This part is required although
// its actual value is ignored.
//
// Note2: A NULL lpszMailboxDN indicates the public store should be opened.
// -----------------------------------------------------------------------------

HRESULT HrMailboxLogon(
	LPMAPISESSION	lpMAPISession,	// MAPI session handle
	LPMDB			lpMDB,			// open message store
	LPCTSTR			lpszMsgStoreDN,	// desired message store DN
	LPCTSTR			lpszMailboxDN,	// desired mailbox DN or NULL
	BOOL			bUseAdminPriv,
	LPMDB			*lppMailboxMDB)	// ptr to mailbox message store ptr
{
	HRESULT					hRes			= S_OK;
	LPEXCHANGEMANAGESTORE	lpXManageStore  = NULL;
	LPMDB					lpMailboxMDB	= NULL;
	ULONG					ulFlags			= 0L;
	SBinary					sbEID			= {0};

	*lppMailboxMDB = NULL;

	if (!lpMAPISession || !lpMDB || !lpszMsgStoreDN || !lppMailboxMDB)
	{
		if (!lpMAPISession) DebugPrint(DBGVerbose,_T("HrMailboxLogon: Session was NULL\n"));
		if (!lpMDB) DebugPrint(DBGVerbose,_T("HrMailboxLogon: MDB was NULL\n"));
		if (!lpszMsgStoreDN) DebugPrint(DBGVerbose,_T("HrMailboxLogon: lpszMsgStoreDN was NULL\n"));
		if (!lppMailboxMDB) DebugPrint(DBGVerbose,_T("HrMailboxLogon: lppMailboxMDB was NULL\n"));
		return MAPI_E_INVALID_PARAMETER;
	}

	// Use a NULL MailboxDN to open the public store
	if (lpszMailboxDN == NULL || !*lpszMailboxDN)
	{
		ulFlags |= OPENSTORE_PUBLIC;
	}

	if (bUseAdminPriv)
	{
		ulFlags |= OPENSTORE_USE_ADMIN_PRIVILEGE;
	}
	
	WC_H(lpMDB->QueryInterface(
		IID_IExchangeManageStore,
		(LPVOID*) &lpXManageStore));
	
	if (lpXManageStore)
	{
		DebugPrint(DBGVerbose,_T("HrMailboxLogon: Creating EntryID. StoreDN = \"%s\", MailboxDN = \"%s\"\n"),lpszMsgStoreDN,lpszMailboxDN);
		
#ifdef _UNICODE
		{
			char *szAnsiMsgStoreDN = NULL;
			char *szAnsiMailboxDN = NULL;

			if (lpszMsgStoreDN) UnicodeToAnsi(lpszMsgStoreDN,&szAnsiMsgStoreDN);
			if (lpszMailboxDN) UnicodeToAnsi(lpszMailboxDN,&szAnsiMailboxDN);
			
			WC_H(lpXManageStore->CreateStoreEntryID(
				szAnsiMsgStoreDN, 
				szAnsiMailboxDN,
				ulFlags | OPENSTORE_TAKE_OWNERSHIP,
				&sbEID.cb, 
				(LPENTRYID*) &sbEID.lpb));
			delete[] szAnsiMsgStoreDN;
			delete[] szAnsiMailboxDN;
		}
#else
		WC_H(lpXManageStore->CreateStoreEntryID(
			(LPSTR) lpszMsgStoreDN, 
			(LPSTR) lpszMailboxDN,
			ulFlags | OPENSTORE_TAKE_OWNERSHIP,
			&sbEID.cb, 
			(LPENTRYID*) &sbEID.lpb));
#endif
		
		WC_H(lpMAPISession->OpenMsgStore(
			NULL,
			sbEID.cb,
			(LPENTRYID) sbEID.lpb,
			NULL, // Interface
			MDB_NO_DIALOG |
			MDB_NO_MAIL |		// spooler not notified of our presence
			MDB_TEMPORARY |	 // message store not added to MAPI profile
			MAPI_BEST_ACCESS,	// normally WRITE, but allow access to RO store
			&lpMailboxMDB));
		
		*lppMailboxMDB = lpMailboxMDB;
	}
	
	MAPIFreeBuffer(sbEID.lpb);
	if (lpXManageStore) lpXManageStore->Release();
	return hRes;
}

//Build DN's for call to HrMailboxLogon
HRESULT OpenOtherUsersMailbox(LPMAPISESSION	lpMAPISession,
							  LPMDB lpMDB,
							  LPCTSTR szServerName,
							  LPCTSTR szMailboxDN,
							  BOOL bUseAdminPriv,
							  LPMDB* lppOtherUserMDB)
{
	HRESULT		 hRes			= S_OK;
	*lppOtherUserMDB = NULL;

	DebugPrint(DBGVerbose,_T("OpenOtherUsersMailbox called with lpMAPISession = 0x%08X, lpMDB = 0x%08X, Server = \"%s\", MailboxDN = \"%s\"\n"),lpMAPISession, lpMDB,szServerName,szMailboxDN);
	if (!szServerName || !lpMAPISession || !lpMDB || !szMailboxDN) return MAPI_E_INVALID_PARAMETER;

	LPTSTR	szServerDN = NULL;

	WC_H(BuildServerDN(
		szServerName,
		_T("/cn=Microsoft Private MDB"),
		&szServerDN));

	if (szServerDN)
	{
		DebugPrint(DBGVerbose,_T("Calling HrMailboxLogon with Server DN = \"%s\"\n"),szServerDN);
		//Any leak in this call is in emsmdb32...
		//See Q237174 
		WC_H(HrMailboxLogon(
			lpMAPISession, 
			lpMDB, 
			szServerDN, 
			szMailboxDN,
			bUseAdminPriv,
			lppOtherUserMDB));
		MAPIFreeBuffer(szServerDN);
	}

	return hRes;
}

//Use these guids:
//pbExchangeProviderPrimaryUserGuid
//pbExchangeProviderDelegateGuid
//pbExchangeProviderPublicGuid
//pbExchangeProviderXportGuid

HRESULT OpenMessageStoreGUID(
							 LPMAPISESSION	lpMAPISession,
							 LPCSTR lpGUID,
							 LPMDB* lppMDB)
{
	LPMAPITABLE	pStoresTbl = NULL;
	LPSRowSet	pRow		= NULL;
	ULONG		ulRowNum;
	HRESULT		hRes = S_OK;

	enum {EID,STORETYPE,NUM_COLS};
	static SizedSPropTagArray(NUM_COLS,sptCols) = {NUM_COLS,
		PR_ENTRYID,
		PR_MDB_PROVIDER
	};

	*lppMDB = NULL;
	if (!lpMAPISession) return MAPI_E_INVALID_PARAMETER;

	WC_H(lpMAPISession->GetMsgStoresTable(0, &pStoresTbl));

	if (pStoresTbl)
	{
		WC_H(HrQueryAllRows(
			pStoresTbl,					//table to query
			(LPSPropTagArray) &sptCols,	//columns to get
			NULL,						//restriction to use
			NULL,						//sort order
			0,							//max number of rows
			&pRow));
		if (pRow)
		{	
			if (!FAILED(hRes)) for (ulRowNum=0; ulRowNum<pRow->cRows; ulRowNum++)
			{
				hRes = S_OK;
				//check to see if we have a folder with a matching GUID
				if (IsEqualMAPIUID(
					pRow->aRow[ulRowNum].lpProps[STORETYPE].Value.bin.lpb,
					lpGUID))
				{
					hRes = lpMAPISession->OpenMsgStore(
						NULL,
						pRow->aRow[ulRowNum].lpProps[EID].Value.bin.cb,
						(LPENTRYID)pRow->aRow[ulRowNum].lpProps[EID].Value.bin.lpb,
						NULL,//Interface
						MAPI_BEST_ACCESS,
						lppMDB);
					break;
				}
			}
		}
	}
	if (!*lppMDB) hRes = MAPI_E_NOT_FOUND;

	FreeProws(pRow);
	if (pStoresTbl) pStoresTbl->Release();
	return hRes;	
}//OpenMessageStoreGUID

///////////////////////////////////////////////////////////////////////////////
//	CopySBinary()
//
//	Parameters
//		psbDest - Address of the destination binary
//		psbSrc  - Address of the source binary
//		lpParent - Pointer to parent object (not, however, pointer to pointer!)
//	  
//	Purpose
//	  Allocates a new SBinary and copies psbSrc into it
//	  
HRESULT CopySBinary(LPSBinary psbDest,const LPSBinary psbSrc, LPVOID lpParent)
{
	HRESULT	 hRes = S_OK;

	if (!psbDest || !psbSrc) return MAPI_E_INVALID_PARAMETER;
	
	psbDest->cb = psbSrc->cb;
	
	if (psbSrc->cb)
	{
		if (lpParent)
		{
			WC_H(MAPIAllocateMore(
					psbSrc->cb,
					lpParent, 
					(LPVOID *) &psbDest->lpb))
		}
		else
		{
			WC_H(MAPIAllocateBuffer(
					psbSrc->cb,
					(LPVOID *) &psbDest->lpb));
		}
		if (S_OK == hRes)
			CopyMemory(psbDest->lpb,psbSrc->lpb,psbSrc->cb);
	}
	
	return hRes;
}//CopySBinary
