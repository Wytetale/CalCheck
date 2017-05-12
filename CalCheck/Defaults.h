#pragma once

#include "StdAfx.h"

STDMETHODIMP OpenDefaultMessageStore(LPMAPISESSION lpMAPISession, ULONG ulExplicitFlags, LPMDB * lpMDB);
STDMETHODIMP OpenDefaultFolder(ULONG ulFolder, LPMDB lpMDB, ULONG ulExplicitFlags, LPMAPIFOLDER *lpFolder);
STDMETHODIMP OpenInbox(LPMDB lpMDB, ULONG ulExplicitFlags, LPMAPIFOLDER *lpInboxFolder);
STDMETHODIMP GetFirstMessage(LPMAPIFOLDER lpFolder, ULONG ulExplicitFlags, LPMESSAGE * lppMessage);
STDMETHODIMP GetDelegates(LPMDB lpMDB);
STDMETHODIMP CreateCalCheckFolder(LPMDB lpMDB, LPMAPIFOLDER *lpCalCheckFld);
STDMETHODIMP CreateReportMsg(LPMDB lpMDB);



const ULONG PR_CALENDAR_ENTRYID =					PROP_TAG(PT_BINARY, 0x36D0);
const ULONG PR_CONTACTS_ENTRYID =					PROP_TAG(PT_BINARY, 0x36D1);
const ULONG PR_JOURNAL_ENTRYID =					PROP_TAG(PT_BINARY, 0x36D2);
const ULONG PR_NOTES_ENTRYID =						PROP_TAG(PT_BINARY, 0x36D3);
const ULONG PR_TASKS_ENTRYID =						PROP_TAG(PT_BINARY, 0x36D4);
const ULONG	PR_REMINDERS_ENTRYID =					PROP_TAG(PT_BINARY, 0x36D5);
const ULONG PR_DRAFTS_ENTRYID =						PROP_TAG(PT_BINARY, 0x36D7);
const ULONG PR_FREEBUSY_ENTRYIDS =					PROP_TAG(PT_MV_BINARY, 0x36E4);
const ULONG PR_SCHDINFO_DELEGATE_NAMES_W =			PROP_TAG(PT_MV_STRING8, 0x684A);	// List of Delegate Names, same order as below prop (delegate flags) 
const ULONG PR_DELEGATE_FLAGS =						PROP_TAG(PT_MV_LONG, 0x686B);		// 1 if delegate can see private items, 0 otherwise
const ULONG PR_FREEBUSY_COUNT_MONTHS =				PROP_TAG(PT_LONG, 0x6869);			// #Months of FB Data 
const ULONG PR_SCHDINFO_BOSS_WANTS_COPY =			PROP_TAG(PT_BOOLEAN, 0x6842);		// Manager receives copy of meeting
const ULONG PR_SCHDINFO_DONT_MAIL_DELEGATES =		PROP_TAG(PT_BOOLEAN, 0x6843);		// Do not send copy to delegate
const ULONG PR_SCHDINFO_BOSS_WANTS_INFO =			PROP_TAG(PT_BOOLEAN, 0x684B);		// Make the manager’s copy an Informational meeting copy
const ULONG PR_SCHDINFO_AUTO_ACCEPT_APPTS =			PROP_TAG(PT_BOOLEAN, 0x686D);		// Automatically accept meeting requests and process cancellations 
const ULONG PR_SCHDINFO_DISALLOW_RECURRING_APPTS =  PROP_TAG(PT_BOOLEAN, 0x686E);		// Automatically decline recurring meeting requests
const ULONG PR_SCHDINFO_DISALLOW_OVERLAPPING_APPTS= PROP_TAG(PT_BOOLEAN, 0x686F);		// Automatically decline conflicting meeting requests



enum {DEFAULT_CALENDAR,
	  DEFAULT_CONTACTS,
	  DEFAULT_JOURNAL,
	  DEFAULT_NOTES,
	  DEFAULT_TASKS,
	  DEFAULT_REMINDERS,
	  DEFAULT_DRAFTS,
	  DEFAULT_SENTITEMS,
	  DEFAULT_OUTBOX,
	  DEFAULT_DELETEDITEMS,
	  DEFAULT_FINDER,
	  DEFAULT_IPM_SUBTREE,
	  DEFAULT_INBOX,
	  NUM_DEFAULT_PROPS};