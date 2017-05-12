

#pragma once

#include "StdAfx.h"
#include "Logging.h"



#define PidLidGlobalObjectId		0x0003
#define PidLidIsRecurring			0x0005
#define PidLidIsException			0x000A
#define PidLidCleanGlobalObjectId   0x0023
#define	dispidApptAuxFlags			0x8207
#define dispidLocation				0x8208
#define dispidApptStartWhole		0x820D
#define dispidApptEndWhole			0x820E
#define dispidApptDuration			0x8213
#define dispidApptSubType			0x8215
#define dispidRecurring				0x8223
#define dispidRecurType				0x8231
#define dispidAllAttendeesString	0x8238
#define dispidToAttendeesString		0x823B
#define dispidCCAttendeesString		0x823C
#define dispidApptStateFlags		0x8217
#define dispidApptRecur				0x8216
#define dispidResponseStatus		0x8218
#define dispidTimeZoneDesc			0x8234
#define dispidTimeZoneStruct		0X8233
#define dispidApptTZDefStartDisplay 0X825E
#define dispidApptTZDefEndDisplay   0X825F
#define dispidApptTZDefRecur		0X8260
#define dispidPropDefStream			0X8540


// dispidRecurType states
#define rectypeNone		(int) 0
#define rectypeDaily	(int) 1
#define rectypeWeekly	(int) 2
#define rectypeMonthly	(int) 3
#define rectypeYearly	(int) 4

// For appointment flags (dispidApptStateFlags)
#define	asfNone			0x0000
#define	asfMeeting		0x0001
#define asfReceived		0x0002
#define asfCancelled	0x0004
#define asfForward		0x0008

// Macros to get safe boolean values for ApptStateFlags
#define FAsfMeeting(x)		((x) & asfMeeting	? true : false)
#define FAsfReceived(x)		((x) & asfReceived	? true : false)
#define FAsfCancelled(x)	((x) & asfCancelled	? true : false)
#define FAsfForward(x)		((x) & asfForward	? true : false)

#define	FAsfOrgdMtg(x)		(FAsfMeeting(x) && !FAsfReceived(x) && !FAsfForward(x))
#define	FAsfRcvdMtg(x)		(FAsfMeeting(x) && FAsfReceived(x))

// For dispidApptAuxFlags
#define auxApptFlagCopied				0x0001
#define auxApptFlagForceMtgResponse		0x0002
#define auxApptFlagForwarded			0x0004

#define FAafCopied(x)		((x) & auxApptFlagCopied ? true : false)
#define FAafForwarded(x)	((x) & auxApptFlagForwarded ? true : false)

// For response states (dispidResponseStatus)
#define respMIN				0
#define	respNone			0
#define	respOrganized		1
#define respTentative		2
#define respAccepted		3
#define respDeclined		4
#define respNotResponded	5
#define respMAX				5


// For returns from TimeCheck
#define tmGood		0
#define tmMin		1
#define tmLow		2
#define tmHigh		3
#define tmError		4


// For Permissions on Calendar Folder
#define aclOwner			2043
#define aclPubEditor		1275
#define aclEditor			1147
#define aclPubAuthor		1179
#define aclAuthor			1051
#define aclNonEditAuthor	1043
#define aclReviewer			1025
#define aclContributor		1026
#define aclNone				1024
#define aclZero				0


// Flags that can be set on PR_RECIPIENT_FLAGS on items in teh Recipients Table
#define recipSendable					0x00000001
#define recipOrganizer					0x00000002
#define recipExceptionalResponse		0x00000010
#define recipExceptionalDeleted			0x00000020
#define recipOriginal					0x00000100

//Filetime stuff for checking.
const FILETIME ftMin =  // 01/01/1601 00:00
	{
	0x00000000,		// dwLowDateTime
	0x00000000		// dwHighDateTime
	};
const FILETIME ftLow =  // 01/01/1995 00:00 - arbitrary check since most calendars wouldn't have data before this date
	{
	0x9F21C000,		// dwLowDateTime
	0x01B9B90A		// dwHighDateTime
	};
const FILETIME ftHigh =  // 01/01/2025 00:00 - arbitrary check since most people won't schedule something out this far intentionally - yet
	{
	0x19BA4000,		// dwLowDateTime
	0x01DB5BE0		// dwHighDateTime
	};
const FILETIME ftClipEnd =  // 12/31/4500 11:59 - an Outlook arbitrary boundary
	{
	0x8019fa00,		// dwLowDateTime
	0x0cb34557		// dwHighDateTime
	};
const FILETIME ftClipEndPlus =  // 01/02/4501 00:00 - just above a different boundary - added due to errors coming back for 01/01/4501 11:59 on NoEndDate items
	{
	0xce470000,		// dwLowDateTime
	0x0cb34620		// dwHighDateTime
	};
const FILETIME ftNone =  // 01/01/4501 00:00 - an Outlook arbitrary boundary
	{
	0xa3dd4000,		// dwLowDateTime
	0x0cb34557		// dwHighDateTime
	};
const FILETIME ftMax = // "BAD"
	{
	0xffffffff,		// dwLowDateTime
	0xffffffff		// dwHighDateTime
	};


//GUID Stuff
DEFINE_GUID(CLSID_MailMessage,
			0x00020D0B,
			0x0000, 0x0000, 0xC0, 0x00, 0x0, 0x00, 0x0, 0x00, 0x00, 0x46);
DEFINE_OLEGUID(PS_INTERNET_HEADERS,	0x00020386, 0, 0);
DEFINE_OLEGUID(PS_PUBLIC_STRINGS,	0x00020329, 0, 0);

// http://msdn2.microsoft.com/en-us/library/bb905283.aspx
DEFINE_OLEGUID(PSETID_Appointment,  MAKELONG(0x2000+(0x2),0x0006),0,0);
DEFINE_OLEGUID(PSETID_Task,         MAKELONG(0x2000+(0x3),0x0006),0,0);
DEFINE_OLEGUID(PSETID_Address,      MAKELONG(0x2000+(0x4),0x0006),0,0);
DEFINE_OLEGUID(PSETID_Common,       MAKELONG(0x2000+(0x8),0x0006),0,0);
DEFINE_OLEGUID(PSETID_Log,          MAKELONG(0x2000+(0xA),0x0006),0,0);
DEFINE_GUID(PSETID_Meeting,
			0x6ED8DA90,
			0x450B, 
			0x101B,	
			0x98, 0xDA, 0x0, 0xAA, 0x0, 0x3F, 0x13, 0x05);
// [MS-OXPROPS].pdf
DEFINE_OLEGUID(PSETID_Note,          MAKELONG(0x2000+(0xE),0x0006),0,0);
DEFINE_OLEGUID(PSETID_Sharing,       MAKELONG(0x2000+(64),0x0006),0,0);
DEFINE_OLEGUID(PSETID_PostRss,       MAKELONG(0x2000+(65),0x0006),0,0);
DEFINE_GUID(PSETID_UnifiedMessaging,
			0x4442858E,
			0xA9E3,
			0x4E80,
			0xB9,0x00,0x31,0x7A,0x21,0x0C,0xC1,0x5B);
DEFINE_GUID(PSETID_AirSync, 0x71035549, 0x0739, 0x4DCB, 0x91, 0x63, 0x00, 0xF0, 0x58, 0x0D, 0xBB, 0xDF);
DEFINE_GUID(PSETID_CalendarAssistant,
			0x11000E07,
			0xB51B,
			0x40D6,
			0xAF, 0x21, 0xCA, 0xA8, 0x5E, 0xDA, 0xB1, 0xD0);

// http://support.microsoft.com/kb/312900
// 53BC2EC0-D953-11CD-9752-00AA004AE40E
DEFINE_GUID(GUID_Dilkie,
			0x53BC2EC0,
			0xD953, 0x11CD,	0x97, 0x52, 0x00, 0xAA, 0x00, 0x4A, 0xE4, 0x0E);

// http://msdn2.microsoft.com/en-us/library/bb820947.aspx
DEFINE_OLEGUID(IID_IMessageRaw, 0x0002038A, 0, 0);

// http://msdn2.microsoft.com/en-us/library/bb820937.aspx
#ifndef DEFINE_PRXGUID
#define DEFINE_PRXGUID(_name, _l)								\
	DEFINE_GUID(_name, (0x29F3AB10 + _l), 0x554D, 0x11D0, 0xA9,	\
				0x7C, 0x00, 0xA0, 0xC9, 0x11, 0xF5, 0x0A)
#endif // !DEFINE_PRXGUID
DEFINE_PRXGUID(IID_IProxyStoreObject, 0x00000000L);

#ifdef _IID_IMAPIClientShutdown_MISSING_IN_HEADER
// http://blogs.msdn.com/stephen_griffin/archive/2009/03/03/fastest-shutdown-in-the-west.aspx
#if !defined(INITGUID) || defined(USES_IID_IMAPIClientShutdown)
DEFINE_OLEGUID(IID_IMAPIClientShutdown, 0x00020397, 0, 0);
#endif
#endif // _IID_IMAPIClientShutdown_MISSING_IN_HEADER

// Sometimes IExchangeManageStore5 is in edkmdb.h, sometimes it isn't
#ifdef USES_IID_IExchangeManageStore5
DEFINE_GUID(IID_IExchangeManageStore5,0x7907dd18, 0xf141, 0x4676, 0xb1, 0x02, 0x37, 0xc9, 0xd9, 0x36, 0x34, 0x30);
#endif

// http://msdn2.microsoft.com/en-us/library/bb820933.aspx
DEFINE_GUID(IID_IAttachmentSecurity,
			0xB2533636,
			0xC3F3, 0x416f, 0xBF, 0x04, 0xAE, 0xFE, 0x41, 0xAB, 0xAA, 0xE2);

// Class Identifiers
// {4e3a7680-b77a-11d0-9da5-00c04fd65685}
DEFINE_GUID(CLSID_IConverterSession, 0x4e3a7680, 0xb77a, 0x11d0, 0x9d, 0xa5, 0x0, 0xc0, 0x4f, 0xd6, 0x56, 0x85);

// Interface Identifiers
// {4b401570-b77b-11d0-9da5-00c04fd65685}
DEFINE_GUID(IID_IConverterSession, 0x4b401570, 0xb77b, 0x11d0, 0x9d, 0xa5, 0x0, 0xc0, 0x4f, 0xd6, 0x56, 0x85);

DEFINE_OLEGUID(PSUNKNOWN, MAKELONG(0x2000+(999),0x0006),0,0);

// [MS-OXCDATA].pdf
// Exchange Private Store Provider
// {20FA551B-66AA-CD11-9BC8-00AA002FC45A}
DEFINE_GUID(g_muidStorePrivate, 0x20FA551B, 0x66AA, 0xCD11, 0x9B, 0xC8, 0x00, 0xAA, 0x00, 0x2F, 0xC4, 0x5A);

// Exchange Public Store Provider
// {1002831C-66AA-CD11-9BC8-00AA002FC45A}
DEFINE_GUID(g_muidStorePublic, 0x1002831C, 0x66AA, 0xCD11, 0x9B, 0xC8, 0x00, 0xAA, 0x00, 0x2F, 0xC4, 0x5A);

// Contact Provider
// {0AAA42FE-C718-101A-E885-0B651C240000}
DEFINE_GUID(muidContabDLL, 0x0AAA42FE, 0xC718, 0x101A, 0xe8, 0x85, 0x0B, 0x65, 0x1C, 0x24, 0x00, 0x00);

// Exchange Public Folder Store Provider
// {9073441A-66AA-CD11-9BC8-00AA002FC45A}
DEFINE_GUID(pbLongTermNonPrivateGuid, 0x9073441A, 0x66AA, 0xCD11, 0x9b, 0xc8, 0x00, 0xaa, 0x00, 0x2f, 0xc4, 0x5a);

// One Off Entry Provider
// {A41F2B81-A3BE-1910-9D6E-00DD010F5402}
DEFINE_GUID(muidOOP, 0xA41F2B81, 0xA3BE, 0x1910, 0x9d, 0x6e, 0x00, 0xdd, 0x01, 0x0f, 0x54, 0x02);

// MAPI Wrapped Message Store Provider
// {10BBA138-E505-1A10-A1BB-08002B2A56C2}
DEFINE_GUID(muidStoreWrap, 0x10BBA138, 0xE505, 0x1A10, 0xa1, 0xbb, 0x08, 0x00, 0x2b, 0x2a, 0x56, 0xc2);

// Exchange Address Book Provider
// {C840A7DC-42C0-1A10-B4B9-08002B2FE182}
DEFINE_GUID(muidEMSAB, 0xC840A7DC, 0x42C0, 0x1A10, 0xB4, 0xB9, 0x08, 0x00, 0x2B, 0x2F, 0xE1, 0x82);

// Contact Address Book Wrapped Entry ID
// {D3AD91C0-9D51-11CF-A4A9-00AA0047FAA4}
DEFINE_GUID(WAB_GUID, 0xD3AD91C0, 0x9D51, 0x11CF, 0xA4, 0xA9, 0x00, 0xAA, 0x00, 0x47, 0xFA, 0xA4);

// TZREG
// =====================
//   This is an individual description that defines when a daylight
//   saving shift, and the return to standard time occurs, and how
//   far the shift is.  This is basically the same as
//   TIME_ZONE_INFORMATION documented in MSDN, except that the strings
//   describing the names "daylight" and "standard" time are omitted.
//
typedef struct TZREG
{
    long        lBias;           // offset from GMT
    long        lStandardBias;   // offset from bias during standard time
    long        lDaylightBias;   // offset from bias during daylight time
    SYSTEMTIME  stStandardDate;  // time to switch to standard time
    SYSTEMTIME  stDaylightDate;  // time to switch to daylight time
} TZREG;


// TZRULE
// =====================
//   This structure represents both a description when a daylight. 
//   saving shift occurs, and in addition, the year in which that
//   timezone rule came into effect. 
//
typedef struct TZRULE
{
    WORD        wFlags;   // indicates which rule matches legacy recur
    SYSTEMTIME  stStart;  // indicates when the rule starts
    TZREG       TZReg;    // the timezone info
} TZRULE;

const WORD TZRULE_FLAG_RECUR_CURRENT_TZREG  = 0x0001; // see dispidApptTZDefRecur
const WORD TZRULE_FLAG_EFFECTIVE_TZREG      = 0x0002;


// TZDEFINITION
// =====================
//   This represents an entire timezone including all historical, current
//   and future timezone shift rules for daylight saving time, etc.  It's
//   identified by a unique GUID.
//
typedef struct TZDEFINITION
{
    WORD     wFlags;       // indicates which fields are valid
    GUID     guidTZID;     // guid uniquely identifying this timezone
    LPWSTR   pwszKeyName;  // the name of the key for this timezone
    WORD     cRules;       // the number of timezone rules for this definition
    TZRULE*  rgRules;      // an array of rules describing when shifts occur
} TZDEFINITION;

const WORD TZDEFINITION_FLAG_VALID_GUID     = 0x0001; // the guid is valid
const WORD TZDEFINITION_FLAG_VALID_KEYNAME  = 0x0002; // the keyname is valid

const ULONG  TZ_MAX_RULES          = 1024; 
const BYTE   TZ_BIN_VERSION_MAJOR  = 0x02; 
const BYTE   TZ_BIN_VERSION_MINOR  = 0x01; 

//
// Function declares

HRESULT ProcessCalendar(LPTSTR szDN, LPMDB lpMDB, LPMAPIFOLDER lpFolder);

HRESULT ProcessRow(LPTSTR szDN, LPMAPIFOLDER lpFolder, LPSRow lpsRow);

void logProps(  CString strError,
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
			 );

void Props2CSV( LPTSTR szDN,
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
				CString strRecurStart,
				CString strRecurEnd,
				CString strEndType,
				CString strNumExceptions,
				CString strNumOccurDeleted,
				CString strNumOccurModified,
				CString strTZStruct,
				BOOL bFlaggedByTool);

ULONG GetRecipData(LPTSTR szDN, LPMAPIFOLDER lpFolder, ULONG cbEID, LPENTRYID lpEID, BOOL bRecur, LPBOOL lpbLogError, CString* strError, CString* strErrorNum);

ULONG GetAttachCount(LPTSTR szDN, LPMAPIFOLDER lpFolder, ULONG cbEID, LPENTRYID lpEID, LPBOOL lpbLogError, CString* strError, CString* strErrorNum);

ULONG TimeCheck(FILETIME ftCheck1, BOOL bRecurCheck);

CString Binary2String(ULONG cbBinVal, LPBYTE lpBinVal);

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
						  CString* strErrorNum);

BOOL GetPropDefStreamData(CString* strError, 
						  CString* strErrorNum);

HRESULT GetACLs(LPMAPIFOLDER lpMAPIFolder);
HRESULT ProcessACLRows(LPMAPIFOLDER lpFolder, LPSRow lpsRow);

FILETIME RTime2FileTime(DWORD rTime);

BOOL CheckSubject(CString strSubject);

BOOL CheckDuplicate(CString strSubject,
					CString strOrganizer,
					CString strLocation,
					BOOL bRecur,
					FILETIME ftStart,
					FILETIME ftEnd);

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
						CString* strRptEnd);


BOOL ReadTestSetting(CString strTest);
BOOL StringFix(CString* strTemp);
BOOL SeparateDateTime(CString* strDateTime, CString* strDate);
BOOL SeparateRecurDateTime(CString* strDateTime, CString* strDate);
TZDEFINITION* BinToTZDEFINITION(ULONG cbDef, LPBYTE lpbDef);

