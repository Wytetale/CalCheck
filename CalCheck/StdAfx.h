// stdafx.h : include file for standard system include files,
//  or project specific include files that are used frequently, but
//      are changed infrequently
//

#pragma once

#define VC_EXTRALEAN		// Exclude rarely-used stuff from Windows headers

#include <afx.h>
#include <afxwin.h>         // MFC core and standard components
#include <afxext.h>         // MFC extensions
#include <afxdtctl.h>		// MFC support for Internet Explorer 4 Common Controls
#ifndef _AFX_NO_AFXCMN_SUPPORT
#include <afxcmn.h>			// MFC support for Windows Common Controls
#endif // _AFX_NO_AFXCMN_SUPPORT

#include <tchar.h>
#include <strsafe.h>
#include <WinNls.h>

#include <MAPIX.h>
#include <MAPIUtil.h>
#include <MAPIHook.h>
#include <MSPST.h>
#include <EdkMdb.h>
#include <EdkGuid.h>
#include <EMSABTAG.H>
#include <MAPITags.h>


#define USES_IID_IMAPIFolder
#include <initguid.h>
#include <mapiguid.h>
#ifdef EDKGUID_INCLUDED
#undef EDKGUID_INCLUDED
#endif
#include <EdkGuid.h>

#include <Msi.h>


//Globals for app
extern BOOL g_bEmitErrorNum;
extern BOOL g_bVerbose;
extern BOOL g_bTime;
extern BOOL g_bCSV; // global for writing to the CSV file
extern BOOL g_bFolder; // global for -F > make a folder and move error items there
extern BOOL g_bMoveItem; // if the item needs to be moved...
extern BOOL g_bOutput; // choose output path
extern BOOL g_bMRMAPI; // If MRMAPI is not present, we don't want to do conversions
extern BOOL g_bInput;
extern CString g_szAppPath;
extern CString g_szOutputPath;
extern CString g_szInputPath;
extern CString g_MBXDisplayName;
extern CString g_MBXLegDN;
extern CString g_UserDisplayName;
extern CString g_UserLegDN;
extern FILETIME g_ftNow; // used to check against appt with warnings. Won't log if they are in the past.
extern BOOL g_bFoundProb;
extern BOOL g_bClearLogs; // For clearing the final Calcheck files when run in server mode
extern BOOL g_bMultiMode; // if we're running in Multi-Mailbox mode
extern BOOL g_bOtherMBXMode; // if we're running against a single other mailbox
extern BOOL g_bReport;  // put a message in the Inbox with the CalCheck.log attached
extern ULONG g_ulItems; // number of items found in the Calendar folder
extern ULONG g_ulRecurItems; // number of recurring appointments in the Calendar
extern int g_ulErrors;
extern int g_ulErrorsSvr;
extern int g_ulWarnings;
extern int g_ulWarningsSvr;
extern int g_ulProbItems;
extern int g_ulProbItemsSvr;
extern CStringList DupList; // duplicate list to check against for duplicate Calendar items
extern CStringList AddrList; // list of proxy addresses for this user
extern CStringList MBXInfoList; // list of mailboxes to run against - read in as an argument



// Flags for cached/offline mode - See http://msdn2.microsoft.com/en-us/library/bb820947.aspx
// Used in OpenMsgStore
#define MDB_ONLINE ((ULONG) 0x00000100)

// Used in OpenEntry
#define MAPI_NO_CACHE ((ULONG) 0x00000200)

/* Flag to keep calls from redirecting in cached mode */
#define MAPI_CACHE_ONLY         ((ULONG) 0x00004000)

#define MAPI_BG_SESSION         0x00200000 /* Used for async profile access */


struct TESTLIST  
{
	BOOL bOrganizerAddress;		// PR_SENT_REPRESENTING_EMAIL_ADDRESS missing
	BOOL bOrganizerName;		// PR_SENT_REPRESENTING_NAME missing
	BOOL bSenderAddress;		// PR_SENDER_EMAIL_ADDRESS missing
	BOOL bSenderName;			// PR_SENDER_NAME missing
	BOOL bNoSubject;			// No Subject if in the future or if recurring
	BOOL bMessageClass;			// Non-standard Message Class check
	BOOL bConflictItems;		// Conflict items can prevent mailbox move in Exchange
	BOOL bRecurItemLimit;		// Report at 1250, 1300 is the limit for number of recurring items in a Calendar
	BOOL bItemSize10;			// Report items over 10M in size
	BOOL bItemSize25;			// Report items over 25M in size
	BOOL bItemSize50;			// Report items over 50M in size
	BOOL bDeliverTime;			// PR_MESSAGE_DELIVERY_TIME missing
	BOOL bAttachCount;			// Report items with more than 25 attachments
	BOOL bRecurringProp;		// dispidRecurring prop missing
	BOOL bStartTimeMin;			// Start Time is set to 0
	BOOL bStartTimeX;			// Missing Start Time (dispidApptStartWhole)
	BOOL bStartTime1995;		// Start Time is before 1995
	BOOL bStartTime2025;		// Start Time is after 2025
	BOOL bStartTimeMax;			// Start Time is after 00:00:00 01/02/4501
	BOOL bEndTimeMin;			// End Time is set to 0
	BOOL bEndTimeX;				// Missing End Time (dispidApptEndWhole)
	BOOL bEndTime1995;			// End Time is before 1995
	BOOL bEndTime2025;			// End Time is after 2025
	BOOL bEndTimeMax;			// End Time is after 00:00:00 01/02/4501
	BOOL bRecurStartMin;		// Recurring Start Time is set to 0
	BOOL bRecurStart1995;		// Recurring Start Time is before 1995
	BOOL bRecurStart2025;		// Recurring Start Time is after 2025
	BOOL bRecurStartMax;		// Recurring Start Time is after 00:00:00 01/02/4501
	BOOL bRecurEndMin;			// Recurring End Time is set to 0
	BOOL bRecurEnd1995;			// Recurring End Time is before 1995
	BOOL bRecurEnd2025;			// Recurring End Time is after 2025
	BOOL bRecurEndMax;			// Recurring End Time is after 00:00:00 01/02/4501
	BOOL bExceptionBounds;		// Check Exceptions in the dispidApptRecur property for items that occur outside the series boundaries
	BOOL bExceptionData;		// Check the Recurrence property for length mismatches on Exception data
	BOOL bDuplicates;			// Find/report duplicate items
	BOOL bAttendToOrganizer;	// Check for items where you are an attendee, and you became the Organizer
	BOOL bDupGlobalObjectIDs;	// Find/report items with duplicate GlobalObjectIds
	BOOL bNoGlobalObjectIDs;	// Find/report items with no GlobalObjectIds
	BOOL bRTAddressType;		// Check valid Email Address Types for recipients in the recipient table
	BOOL bRTAddress;			// Check valid Email Addresses for recipients in the recipient table
	BOOL bRTDisplayName;		// Check valid/existing Display Names for recipients in the recipient table
	BOOL bRTDuplicates;			// Check for duplicated recipients in the recipient table
	BOOL bRTOrganizerAddress;	// Check valid Email Address for recipient marked as Organizer in the recipient table
	BOOL bRTOrganizerIsOrganizer; // Check that the recipient marked as Organizer in the recipient table matches the Organizer in SENT_REPRESENTING
	BOOL bTZDefRecur;			// Check for bad dispidApptTZDefRecur prop
	BOOL bTZDefStart;			// Check for missing dispidApptTZDefStartDisplay prop
	BOOL bPropDefStream;		// Check for corrupt dispidPropDefStream prop
	BOOL bHolidayItems;			// Run checks on Holiday items
	BOOL bBirthdayItems;		// Run checks on Birthday items
	BOOL bPastItems;			// Run checks on items that ended in the past
	BOOL bWarningIsError;		// Treat Warnings like Errors - moves Warning errors to CalCheck Folder when -F switch is used
};

extern TESTLIST CalCheckTests;

struct MYOPTIONS
{
	LPSTR lpszFileName;
	LPSTR lpszProfileName;
	LPSTR lpszMailboxDN;
	LPSTR lpszDisplayName;
	LPSTR lpszMAPIVer;
};

extern MYOPTIONS ProgOpts;