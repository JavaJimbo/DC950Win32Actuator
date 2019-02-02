/*	New Excel Routines
*	
*	Created 3-21-18 
*	MyExcelApp.OpenExcelBook((LPCTSTR) CurrentDataFilename, TRUE);
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

extern char dateBuffer[80];
extern const char *CurrentDataFilename; 
extern const char *SpreadsheetTemplateFilename;
extern char ExcelBackupFilename[];

extern BOOL copySpreadsheet;
struct testData arrTestData[TOTAL_COLUMNS];

CString DataFormat[TOTAL_COLUMNS] =
{
	"mm-dd-yy",
	"General",
	"General",
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

void TestApp::Errorf( LPCTSTR pszFormat, ... )
{
	
	CString		str;
	va_list	argList;
	va_start( argList, pszFormat );
	str.FormatV( pszFormat, argList );
	MessageBox((LPCTSTR) str, _T("WTLExcel Error"), MB_ICONHAND | MB_OK );
	return;
}


void TestApp::ShutDownExcel()
{	
	;
}

BOOL TestApp::storeTestResult(int col, int testResult, char *ptrTestData)
{		
	int diddyWahDiddy = 0;

	if (col > LASTCOLUMN) 
		return false;
	if (col == FINAL_RESULT)
		diddyWahDiddy++;

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

	insertBogusData(); // $$$$


	//if (arrTestData[0].result.GetLength() == 0) return FALSE;

	//if(!SUCCEEDED(OleInitialize(NULL)))
    //{
	//	TRACE("Failed to initialize OLE");
	//	return FALSE;
    //}

	Excel::_ApplicationPtr pApplication = NULL;
	Excel::_WorkbookPtr pBook;
	Excel::_WorksheetPtr pSheet;
	
	if ( FAILED( pApplication.CreateInstance( _T("Excel.Application") ) ) )
	{
		Errorf( _T("Failed to initialize Excel::_Application!") );
		return FALSE;
	}
	
	_variant_t	varOption( (long) DISP_E_PARAMNOTFOUND, VT_ERROR );

	pBook = pApplication->Workbooks->Open(CurrentDataFilename, varOption, varOption, varOption, varOption, varOption, varOption, varOption, varOption, varOption, varOption, varOption, varOption );


	if ( pBook == NULL )
	{
		Errorf( _T("Failed to open Excel file!") );
		return FALSE;
	}	
	
	pSheet = pBook->Sheets->Item[ 1 ];

	if ( pSheet == NULL )
	{
		Errorf( _T("Failed to get first Worksheet!") );
		return FALSE;
	}			
	
	// Search down spreadsheet to find first available unused row.
	// If spreadsheet is full or corrupted, quit now.
	if (!getCurrentRow(pSheet, &currentRow)) return FALSE;	
		
	for (col = 0; col < TOTAL_COLUMNS; col++)
	{		
		makeRangeString(currentRow, col+1, strRange);	// Data array is numbered beginning at 0, 
														// but first Excel column is 1, so offset is added here
		Excel::RangePtr pRange = pSheet->GetRange( _bstr_t(strRange), _bstr_t(strRange));		
		pRange->Font->Color = arrTestData[col].color;
		pRange->NumberFormat = _bstr_t(DataFormat[col]);		
		pRange->Item [1][1] =  _bstr_t(arrTestData[col].result);
	}

    // Switch off alert prompting to save as 
	pApplication->DisplayAlerts = false;
    // Save the values in book.xml and release resources
    pSheet->SaveAs(CurrentDataFilename);	
	pBook->Close( VARIANT_FALSE );
	
	// Need to quit, otherwise Excel remains active and locks the .xls file.
	pApplication->Quit( );	
	resetTestData();	

	return TRUE;
}


BOOL TestApp::getCurrentRow(Excel::_WorksheetPtr pSheet, int *ptrCurrentRow)
{
	char strStartRange[32], strEndRange[32];
	int row;
	if (*ptrCurrentRow == NULL || pSheet == NULL) return FALSE;
	makeRangeString(1, 1, strStartRange);
	makeRangeString(MAXROWS, LASTCOLUMN, strEndRange);
	Excel::RangePtr pRange = pSheet->GetRange( _bstr_t(strStartRange), _bstr_t(strEndRange));

	if (*ptrCurrentRow < 4) *ptrCurrentRow = 4;
   	for (row = *ptrCurrentRow; row < MAXROWS; row++)
	{
		CString strCell = pRange->Item[row][1];			
		if (strCell.GetLength()== 0) break;
	}
	if (row == MAXROWS) return FALSE;
	*ptrCurrentRow = row;
	return TRUE;
}