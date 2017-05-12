// Prototypes for Logging functionality
#pragma once

#include "StdAfx.h"


BOOL WarnHResFn(HRESULT hRes,LPCSTR szFunction,LPCTSTR szErrorMsg,LPCSTR szFile,int iLine);

#define	DBGGeneric	((ULONG) 0x00000000)
#define	DBGVerbose	((ULONG) 0x00000001)
#define	DBGHelp 	((ULONG) 0x00000002)
#define	DBGTime 	((ULONG) 0x00000003)
#define DBGError	((ULONG) 0x00000004)

void __cdecl DebugPrint(ULONG ulDbgLvl, LPCTSTR szMsg,...);

void DisplayError(
				  HRESULT hRes, 
				  LPCTSTR szErrorMsg, 
				  LPCSTR szFile,
				  int iLine,
				  LPCSTR szFunction);

#define WC_H(fnx)	\
	hRes = (fnx);	\
	WarnHResFn(hRes,#fnx,_T(""),__FILE__,__LINE__);	

//macro to get allocated string lengths
#define CCH(string) (sizeof((string))/sizeof(TCHAR))

BOOL CheckStringProp(LPSPropValue lpProp, ULONG ulPropType);
//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.


HRESULT CreateLogFiles();
void RemoveLogFiles();
HRESULT CreateServerLogFile();
void WriteLogFile(LPCTSTR strWriteLine, BOOL bSrvLog);
void WriteErrorFile(LPCTSTR strWriteLine);
void WriteCSVFile(LPCTSTR strWriteLine);
void WriteSvrLogFile(LPCTSTR strWriteLine);
CString GetDNUser(LPTSTR szDN);
void AppendFileNames(LPCTSTR strAppend);
BOOL ConvertPropDefStream(LPCTSTR strStream, CString* strRetError);
BOOL ConvertApptRecurBlob(LPCTSTR strBlob, CString* strRetError);
HRESULT RunMRMAPI(LPCTSTR strSwitches);
void RemoveTempFiles();