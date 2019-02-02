/*	New Excel Routines
*
*	Adapted from Code Project article: "MS Office OLE Automation Using C++" by Naren Neelamegam, 5 Apr 2009
*	Link: https://www.codeproject.com/Articles/34998/MS-Office-OLE-Automation-Using-C
*	Also helpful - Microsoft's article "How to automate Excel from C++ without using MFC or #import"
*	Link: https://support.microsoft.com/en-us/help/216686/how-to-automate-excel-from-c-without-using-mfc-or-import
*	
*	3-28-18 JBS:	
*/
#include "stdafx.h"
#include "DC950Test.h"
#include "DC950TestDialog.h"
#include "TestApp.h"
#include "Definitions.h"
#include <windows.h>
#include <iostream>
#include <sstream>
#include <string.h>
#include "avaspec.h"
#include <time.h>
#include "MSExcel.h"
#include "OLEMethod.h"
#include <ole2.h> // OLE2 Definitions

extern CMSExcel MyExcelApp;


extern char dateBuffer[80];
extern const char *CurrentDataFilename; 
extern const char *SpreadsheetTemplateFilename;
extern char ExcelBackupFilename[];

extern BOOL copySpreadsheet;
struct testData arrTestData[TOTAL_COLUMNS];

CString DataFormat[TOTAL_COLUMNS] =
{
	"mm-dd-yyyy",
	"0",
	"0",
	"General",
	"General",
	"0.00",
	"0.00",
	"0.00",
	"0.00",
	"0.00",
	"0.00",
	"0.00",
	"0.00",
	"0.00",
	"0.00",
	"0.00",
	"0.00",
	"0.00",
	"0.00",
	"0.00",
	"0.00",
	"General",
	"0.00",
	"0.00",
	"0.00",
	"0.00",
	"0.00",
	"0.00",	
	"General"	
};


void TestApp::makeRangeString(int row, int column, char *strCell)
{		
	char strCol[16];
	if (column <= 26)
	{
		strCol[0] = 'A' + column - 1;
		strCol[1] = '\0';
	}
	else
	{
		strCol[0] = 'A';
		strCol[1] = 'A' + column - 27;
		strCol[2] = '\0';
	}
	sprintf_s(strCell, 16, "%s%d", strCol, row);	
}



void TestApp::ShutDownExcel()
{	
	;
}

BOOL TestApp::storeTestResult(int col, int testResult, char *ptrTestData)
{		
	if (col > LASTCOLUMN) 
		return false;

	if (testResult == PASS) 
	{
		arrTestData[col].color = BLACK;
		if (ptrTestData == NULL) arrTestData[col].result.SetString("PASS");
		else arrTestData[col].result.SetString(ptrTestData);
	}
	else 
	{
		arrTestData[col].color = RED;
		if (ptrTestData == NULL) arrTestData[col].result.SetString("FAIL");
		else arrTestData[col].result.SetString(ptrTestData);
	}
	return true;
}




BOOL TestApp::storeSerialNumberAndPartNumber(char *ptrSerialNumber, char *ptrPartNumber) {	
char dateBuffer[32];
struct tm  tstruct;		

	time_t now = time(0);
	tstruct = *localtime(&now);
	strftime(dateBuffer, sizeof(dateBuffer), "%m-%d-%Y", &tstruct);	

	arrTestData[0].result.SetString(dateBuffer);
	arrTestData[0].color = BLACK;
	arrTestData[1].result.SetString(ptrSerialNumber);
	arrTestData[1].color = BLACK;
	arrTestData[2].result.SetString(ptrPartNumber);
	arrTestData[2].color = BLACK;

	return TRUE;
}



BOOL TestApp::validateSpreedsheetFilename(char *ptrFilename)
{	
int i = 0;
char ch;

	if (strlen(ptrFilename) == 0) return FALSE;  // $$$$

	for (i = 0; i < MAXFILENAME-6; i++)
	{
		ch = ptrFilename[i];
		if (ch == '[' || ch == ']' || ch == '<' || ch == '>' || ch == '?' || ch == '/' || ch == '*' || ch == '"' || ch == '|')
			return FALSE;

		if (ch == '.')
		{
			if (ptrFilename[i+1] == 'x' && ptrFilename[i+2] == 'l' && ptrFilename[i+3] == 's')
			{
				ptrFilename[i+4] = '\0';
				return TRUE;
			}
		}
	}
	return FALSE;
}

BOOL TestApp::createNewSpreadsheet()
{
    DWORD fileAttr;

    fileAttr = GetFileAttributes((LPCSTR)SpreadsheetTemplateFilename);
    if (0xFFFFFFFF == fileAttr && GetLastError()==ERROR_FILE_NOT_FOUND)
        return FALSE;

	CopyFile((LPCSTR)SpreadsheetTemplateFilename, CurrentDataFilename, FALSE);
	return TRUE;
}


BOOL TestApp::SaveSpreadsheet() {	
	#define STRINGSIZE 32
	char strRange[STRINGSIZE];
	static int currentRow = 4;  
	int col;
	HRESULT hr;

	if (arrTestData[0].result.GetLength() == 0) 
		return TRUE;
		
	hr = MyExcelApp.OpenExcelBook((LPCTSTR) CurrentDataFilename, FALSE);
	if (FAILED(hr)) 
		return FALSE;	
	
	// Search down spreadsheet to find first available unused row.
	// If spreadsheet is full or corrupted, quit now.
	if (!getCurrentRow(&currentRow)) 
		return FALSE;	
		
	for (col = 0; col < TOTAL_COLUMNS; col++)
	{		
		makeRangeString(currentRow, col+1, strRange);	// Data array is numbered beginning at 0, 
														// but first Excel column is 1, so offset is added here
		BOOL boldFlag = TRUE;
		if (col < 3) boldFlag = FALSE;
		hr = MyExcelApp.SetCellValueAndColor(strRange, LPCTSTR (arrTestData[col].result), LPCSTR (DataFormat[col]), arrTestData[col].color, boldFlag);
		if (FAILED(hr)) 
			return FALSE;
	}
	hr = MyExcelApp.TurnOffDisplayAlerts();
	hr = MyExcelApp.Save((LPCTSTR) CurrentDataFilename);
	if (FAILED(hr)) 
		return FALSE;		
	MyExcelApp.Quit();
	resetTestData();
	return TRUE;
}

BOOL TestApp::getCurrentRow(int *ptrCurrentRow)
{
char strRange[32];
static int row = 4;
HRESULT hr;
CString sValue;
	
	for (row = *ptrCurrentRow; row < MAXROWS; row++) 	
	{
		makeRangeString(row, 1, strRange);
		hr = MyExcelApp.GetExcelValue((LPCTSTR) strRange, sValue);
		if (FAILED(hr)) return FALSE;
		if (sValue.GetLength() == 0) break;
	}

	if (row == MAXROWS) return FALSE;
	*ptrCurrentRow = row;
	return TRUE;
}

