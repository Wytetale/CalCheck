// MAPIStoreFunctions.h : Stand alone MAPI Store functions

#pragma once

#include "StdAfx.h"
#include "Logging.h"



HRESULT BuildServerDN(
					  LPCTSTR szServerName,
					  LPCTSTR szPost,
					  LPTSTR* lpszServerDN);
HRESULT GetMailboxTable(
						LPMDB lpMDB,
						LPCTSTR szServerName,
						ULONG ulOffset,
						LPMAPITABLE* lpMailboxTable);
HRESULT GetMailboxTable1(
						LPMDB lpMDB,
						LPCTSTR szServerDN,
						ULONG ulFlags,
						LPMAPITABLE* lpMailboxTable);
HRESULT GetMailboxTable3(
						LPMDB lpMDB,
						LPCTSTR szServerDN,
						ULONG ulOffset,
						ULONG ulFlags,
						LPMAPITABLE* lpMailboxTable);
HRESULT HrMailboxLogon(
					   LPMAPISESSION	lplhSession,	// ptr to MAPI session handle
					   LPMDB			lpMDB,			// ptr to message store
					   LPCTSTR			lpszMsgStoreDN,	// ptr to message store DN
					   LPCTSTR			lpszMailboxDN,  // ptr to mailbox DN
					   BOOL				bUseAdminPriv,
					   LPMDB*			lppMailboxMDB);	// ptr to mailbox message store ptr
HRESULT OpenOtherUsersMailbox(
							  LPMAPISESSION	lpMAPISession,
							  LPMDB lpMDB,
							  LPCTSTR szServerName,
							  LPCTSTR szMailboxDN,
							  BOOL bUseAdminPriv,
							  LPMDB* lppOtherUserMDB);
HRESULT OpenMessageStoreGUID(
							  LPMAPISESSION	lpMAPISession,
							  LPCSTR lpGUID,
							  LPMDB* lppMDB);


HRESULT	CopySBinary(LPSBinary psbDest,const LPSBinary psbSrc, LPVOID lpParent);
int AnsiToUnicode(LPCSTR pszA, LPWSTR* ppszW);
int UnicodeToAnsi(LPCWSTR pszW, LPSTR* ppszA);