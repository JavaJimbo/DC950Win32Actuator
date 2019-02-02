/* Display.cpp - Routines for formatting and displaying text on main dialog box.
 * Written in Visual C++ for DC950 test system
 *
 * 
 */

// NOTE: INCUDES MUST BE IN THIS ORDER!!!
#include "stdafx.h"
#include "DC950Test.h"
#include "DC950TestDialog.h"
#include <string.h>
#include "TestApp.h"
#include "Definitions.h"
#include <iostream>
#include <sstream>
#include <windows.h>
#include "avaspec.h"
//#include "BasicExcel.hpp"
//#include "ExcelFormat.h"
#include <time.h>
#include "ReadOnlyEdit.h"
#include "PasswordEntry_Dlg.h"	

extern struct testData arrTestData[];

BOOL lineOneEmpty = TRUE;
BOOL lineTwoEmpty = TRUE;
BOOL lineThreeEmpty = TRUE;
BOOL lineFourEmpty = TRUE;
BOOL lineFiveEmpty = TRUE;

extern CTestDialog *ptrDialog;
CReadOnlyEdit	*arrayEditBoxTest[LAST_TEST_EDIT_BOX+1];
extern CFont BigFont, SmallFont, MidFont;

char *arrTestTitles[] = {NULL, "Hi-Pot",         "Ground Bond",     "Pot at zero",  "Pot at Full",	"Remote Lamp",  "AC Sweep",    "Actuator",     "Filter ON",     "Filter OFF",      "Final Assembly"};
char strTestLog[MAXLOG] = "";


char strLineOne[MAXLINE] = {'\0'};
char strLineTwo[MAXLINE] = {'\0'};
char strLineThree[MAXLINE] = {'\0'};
char strLineFour[MAXLINE] = {'\0'};
char strLineFive[MAXLINE] = {'\0'};

void TestApp::ConfigureFont(CFont &ptrFont, int fontHeight, int fontWidth, BOOL flgBold) {
	if (flgBold) {
		ptrFont.CreateFont(
			fontHeight,
			fontWidth,
			0,
			FW_BOLD,
			FW_DONTCARE,
			FALSE,
			FALSE,
			FALSE,
			DEFAULT_CHARSET,
			OUT_DEFAULT_PRECIS,
			CLIP_DEFAULT_PRECIS,
			DEFAULT_QUALITY,
			DEFAULT_PITCH,
			NULL
		);
	}
	else
	{
		ptrFont.CreateFont(
			fontHeight,
			fontWidth,
			0,
			FW_NORMAL,
			FW_DONTCARE,
			FALSE,
			FALSE,
			FALSE,
			DEFAULT_CHARSET,
			OUT_DEFAULT_PRECIS,
			CLIP_DEFAULT_PRECIS,
			DEFAULT_QUALITY,
			DEFAULT_PITCH,
			NULL
		);
	}
}

	BOOL TestApp::InitializeFonts(){
		ConfigureFont(BigFont, 24, 12, TRUE);
		ConfigureFont(MidFont, 21, 10, TRUE);
		ConfigureFont(SmallFont, 15, 5, TRUE);
		return TRUE;
	}

	BOOL TestApp::DisplayTestEditBox(int boxNumber, int passFailStatus){
		char TestEditBoxText[MAXSTRING];

		if (boxNumber > LAST_TEST_EDIT_BOX) return FALSE;
		else if (boxNumber == 0) return FALSE;
		if (arrayEditBoxTest[boxNumber] == NULL) return FALSE;
		if (arrTestTitles[boxNumber] == NULL) return FALSE;

		strcpy_s (TestEditBoxText, MAXSTRING, arrTestTitles[boxNumber]);

		if (passFailStatus == NOT_DONE_YET){
			arrayEditBoxTest[boxNumber]->SetBackColor(BLACK);
			arrayEditBoxTest[boxNumber]->SetTextColor(YELLOW);		

		}

		else if (passFailStatus == PASS){
			arrayEditBoxTest[boxNumber]->SetBackColor(GREEN);
			arrayEditBoxTest[boxNumber]->SetTextColor(BLACK);		
			strcat_s (TestEditBoxText, MAXSTRING, ":  PASS");
		}

		else if (passFailStatus == FAIL){
			arrayEditBoxTest[boxNumber]->SetBackColor(RED);
			arrayEditBoxTest[boxNumber]->SetTextColor(BLACK);	
			strcat_s (TestEditBoxText, MAXSTRING, ":  FAIL");
		}

		arrayEditBoxTest[boxNumber]->SetFont(&BigFont, TRUE);
		arrayEditBoxTest[boxNumber]->SetWindowText((LPCTSTR)TestEditBoxText);

		return TRUE;
	}

	BOOL TestApp::InitializeTestEditBoxes()
	{
		arrayEditBoxTest[0] = NULL;
		arrayEditBoxTest[1] = &ptrDialog->m_EditBox_Test1;
		arrayEditBoxTest[2] = &ptrDialog->m_EditBox_Test2;
		arrayEditBoxTest[3] = &ptrDialog->m_EditBox_Test3;
		arrayEditBoxTest[4] = &ptrDialog->m_EditBox_Test4;
		arrayEditBoxTest[5] = &ptrDialog->m_EditBox_Test5;
		arrayEditBoxTest[6] = &ptrDialog->m_EditBox_Test6;
		arrayEditBoxTest[7] = &ptrDialog->m_EditBox_Test7;
		arrayEditBoxTest[8] = &ptrDialog->m_EditBox_Test8;	
		arrayEditBoxTest[9] = &ptrDialog->m_EditBox_Test9;		
		arrayEditBoxTest[10] = &ptrDialog->m_EditBox_Test10;		
		return TRUE;
	}



	BOOL TestApp::InitializeDisplayInstructions() {	
		int i = 0;

		// 0 PRESET				
		testStep[i].lineOne = "TEST SYSTEM PREPARATION";		
		testStep[i].lineTwo = "Before beginning, make sure California Instruments supply,";
		testStep[i].lineThree = "Interface Box, and Multimeter are ON. Then press ENTER and wait";
		testStep[i].lineFour = "while serial ports and hardware start up.";
		testStep[i].lineFive = "This may take up to a minute.";
		testStep[i].stepType = MANUAL;
		testStep[i].enableENTER = TRUE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableRETRY = FALSE;
		testStep[i].enablePREVIOUS = FALSE;		
		testStep[i].enableADMIN = TRUE;
		testStep[i].editBoxNumber = 0;
		i++;

		// 1 SCAN BARCODE			
		testStep[i].lineOne = "BARCODE SCAN";
		testStep[i].lineTwo = "";
		testStep[i].lineThree = "Scan unit barcode. If barcode doesn't appear";
		testStep[i].lineFour = "in box above, try scanning it again.";
		testStep[i].lineFive = "";
		testStep[i].stepType = MANUAL;
		testStep[i].enableENTER = TRUE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableRETRY = FALSE;
		testStep[i].enablePREVIOUS = FALSE;		
		testStep[i].enableADMIN = TRUE;
		testStep[i].editBoxNumber = 0;
		i++;

		// 2 SCAN PART NUMBER		
		testStep[i].lineOne = "PART NUMBER SCAN";
		testStep[i].lineTwo = "";
		testStep[i].lineThree = "Scan 12 digit model number so it appears above.";
		testStep[i].lineFour = "Then plug in DC950 and turn it on. Click ENTER to begin.";
		testStep[i].lineFive = "";
		testStep[i].stepType = MANUAL;
		testStep[i].enableENTER = TRUE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableRETRY = FALSE;
		testStep[i].enablePREVIOUS = TRUE;	
		testStep[i].enableADMIN = TRUE;
		testStep[i].editBoxNumber = 0;
		i++;

		  

		// 3 HI POT TEST				
		testStep[i].lineOne = "HI-POT TEST";
		testStep[i].lineTwo = "";
		testStep[i].lineThree = "Connect tester, then run test.";
		testStep[i].lineFour = "Click PASS or FAIL";
		testStep[i].lineFive = "";
		testStep[i].stepType = MANUAL;
		testStep[i].enableENTER = FALSE;
		testStep[i].enablePASS = TRUE;
		testStep[i].enableFAIL = TRUE;		
		testStep[i].enableRETRY = FALSE;
		testStep[i].enablePREVIOUS = TRUE;		
		testStep[i].enableADMIN = FALSE;
		testStep[i].editBoxNumber = HI_POT_EDIT;
		i++;

		// 4 GROUND BOND TEST			
		testStep[i].lineOne = "GROUND BOND TEST";
		testStep[i].lineTwo = "";
		testStep[i].lineThree = "Connect tester, then run test.";
		testStep[i].lineFour = "Click PASS or FAIL";
		testStep[i].lineFive = "";
		testStep[i].stepType = MANUAL;
		testStep[i].enableENTER = FALSE;
		testStep[i].enablePASS = TRUE;
		testStep[i].enableFAIL = TRUE;
		testStep[i].enableRETRY = FALSE;
		testStep[i].enablePREVIOUS = TRUE;	
		testStep[i].enableADMIN = FALSE;
		testStep[i].editBoxNumber = GROUND_BOND_EDIT;
		i++;

		// 5 LAMP OFF		
		testStep[i].lineOne = "POT AT ZERO TEST";
		testStep[i].lineTwo = "Remove lamp assembly from DC950 and connect test lamp.";
		testStep[i].lineThree = " Insert block inside DC950 to push in black Cherry switch.";
		testStep[i].lineFour = "Set switch on rear of DC950 to LOCAL. Turn POT to zero";
		testStep[i].lineFive = "Turn ON the DC950, and then click ENTER to run test";
		testStep[i].stepType = MANUAL;
		testStep[i].enableENTER = TRUE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableRETRY = FALSE;
		testStep[i].enablePREVIOUS = TRUE;
		testStep[i].enableADMIN = FALSE;
		testStep[i].editBoxNumber = POT_LOW_EDIT;
		i++;

		// 6 LAMP ON		
		testStep[i].lineOne = "POT AT FULL TEST";
		testStep[i].lineTwo = "";
		testStep[i].lineThree = "Turn POT up to maximum.";
		testStep[i].lineFour = "Then click ENTER to run test";
		testStep[i].lineFive = "";
		testStep[i].stepType = MANUAL;
		testStep[i].enableENTER = TRUE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableRETRY = FALSE;
		testStep[i].enablePREVIOUS = TRUE;
		testStep[i].enableADMIN = FALSE;
		testStep[i].editBoxNumber = POT_HIGH_EDIT;
		i++;

		// 7 REMOTE SETUP			
		testStep[i].lineOne = "REMOTE TEST"; 
		testStep[i].lineTwo = "";
		testStep[i].lineThree = "Turn pot to zero. Set switch on rear of DC950 to REMOTE.";
		testStep[i].lineFour = "Plug in DB9 connectors at rear of DC950.";
		testStep[i].lineFive = "Then click ENTER to run test";
		testStep[i].stepType = MANUAL; //  SUBSTEP;
		testStep[i].enableENTER = TRUE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableRETRY = FALSE;
		testStep[i].enablePREVIOUS = TRUE;
		testStep[i].enableADMIN = FALSE;
		testStep[i].editBoxNumber = 0;
		i++;
		
		// 8 REMOTE POT				
		testStep[i].lineOne = "REMOTE TEST";
		testStep[i].lineTwo = "";
		testStep[i].lineThree = "Running test. Please wait...";
		testStep[i].lineFour = "";
		testStep[i].lineFive = "";
		testStep[i].stepType = AUTO;
		testStep[i].enableENTER = FALSE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableRETRY = FALSE; 
		testStep[i].enablePREVIOUS = FALSE;
		testStep[i].enableADMIN = FALSE;
		testStep[i].editBoxNumber = REMOTE_EDIT;
		i++;

		// 9 AC SWEEP		
		testStep[i].lineOne = "AC SWEEP TEST";
		testStep[i].lineTwo = "";
		testStep[i].lineThree = "Running test. Please wait...";
		testStep[i].lineFour = "";
		testStep[i].lineFive = "";
		testStep[i].stepType = AUTO;
		testStep[i].enableENTER = FALSE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableRETRY = FALSE; 
		testStep[i].enablePREVIOUS = FALSE;
		testStep[i].enableADMIN = FALSE;
		testStep[i].editBoxNumber = AC_SWEEP_EDIT;
		i++;

		// 10 DARK SCAN SETUP		
		testStep[i].lineOne = "DARK REFERENCE SCAN";
		testStep[i].lineTwo = "";
		testStep[i].lineThree = "Insert spectrometer probe in reference tube.";
		testStep[i].lineFour = "Then click ENTER to begin spectrometer scan";
		testStep[i].lineFive = "";
		testStep[i].stepType = MANUAL;
		testStep[i].enableENTER = TRUE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableRETRY = FALSE;
		testStep[i].enablePREVIOUS = TRUE;
		testStep[i].enableADMIN = FALSE;
		testStep[i].editBoxNumber = 0;
		i++;

		// 11 DARK SCAN RUN		
		testStep[i].lineOne = "DARK REFERENCE SCAN";
		testStep[i].lineTwo = "Spectrometer is scanning - please wait";
		testStep[i].lineThree = " ";
		testStep[i].lineFour = "While waiting for scan to complete, turn off DC950 and install";
		testStep[i].lineFive = "lamp housing back inside unit. Then turn on DC950";
		testStep[i].stepType = AUTO;
		testStep[i].enableENTER = FALSE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableRETRY = FALSE; 
		testStep[i].enablePREVIOUS = FALSE;
		testStep[i].enableADMIN = FALSE;
		testStep[i].editBoxNumber = 0;
		i++;
		
		// 12 DC950 SCAN SETUP		
		testStep[i].lineOne = "LAMP SCAN ON ASSEMBLED UNIT";
		testStep[i].lineTwo = "";
		testStep[i].lineThree = "Before beginning, make sure lamp housing is installed back inside unit.";
		testStep[i].lineFour =  "Insert spectrometer probe in front of DC950.";
		testStep[i].lineFive = "Turn on power switch and wait for lamp to warm up. Then click ENTER";
		testStep[i].stepType = MANUAL;
		testStep[i].enableENTER = TRUE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableRETRY = FALSE;
		testStep[i].enablePREVIOUS = TRUE;
		testStep[i].enableADMIN = FALSE;
		testStep[i].editBoxNumber = 0;
		i++;

		// 13 ACTUATOR TEST		
		testStep[i].lineOne = "ACTUATOR TEST";
		testStep[i].lineTwo = "";
		testStep[i].lineThree = "Checking TTL input";
		testStep[i].lineFour = "Please wait...";
		testStep[i].lineFive = "";
		testStep[i].stepType = AUTO;
		testStep[i].enableENTER = FALSE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableRETRY = FALSE; 
		testStep[i].enablePREVIOUS = FALSE;
		testStep[i].enableADMIN = FALSE;
		testStep[i].editBoxNumber = ACTUATOR_EDIT;
		i++;

		// 14 SCAN RUN FILTER CLOSED		
		testStep[i].lineOne = "SCAN WITH FILTER ON (CLOSED)";
		testStep[i].lineTwo = "";
		testStep[i].lineThree = "Spectrometer is scanning - please wait";
		testStep[i].lineFour = "";
		testStep[i].lineFive = "";
		testStep[i].stepType = AUTO;
		testStep[i].enableENTER = FALSE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableRETRY = FALSE; 
		testStep[i].enablePREVIOUS = FALSE;
		testStep[i].enableADMIN = FALSE;
		testStep[i].editBoxNumber = FILTER_ON_EDIT;
		i++;

		// 15 SCAN RUN FILTER OPEN		
		testStep[i].lineOne = "SCAN WITH FILTER OFF (OPEN)";
		testStep[i].lineTwo = "";
		testStep[i].lineThree = "Running Spectrometer test";
		testStep[i].lineFour = "Please wait...";
		testStep[i].lineFive = "";
		testStep[i].stepType = AUTO;
		testStep[i].enableENTER = FALSE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableRETRY = FALSE; 
		testStep[i].enablePREVIOUS = FALSE;
		testStep[i].enableADMIN = FALSE;
		testStep[i].editBoxNumber = FILTER_OFF_EDIT;
		i++;

		// 16 FINAL_ASSEMBLY
		testStep[i].lineOne = "FINAL ASSEMBLY";
		testStep[i].lineTwo = "Turn off power switch and install lamp housing back inside DC950.";
		testStep[i].lineThree = "Flip switch on rear of DC950 back to LOCAL, then turn power on.";
		testStep[i].lineFour = "Turn pot up and down to ensure lamp works, then click PASS or FAIL.";
		testStep[i].lineFive = "Turn off power, disconnect test leads, and remove unit.";
		testStep[i].stepType = MANUAL;
		testStep[i].enableENTER = FALSE;
		testStep[i].enablePASS = TRUE;
		testStep[i].enableFAIL = TRUE;
		testStep[i].enableRETRY = FALSE;
		testStep[i].enablePREVIOUS = FALSE;
		testStep[i].enableADMIN = FALSE;
		testStep[i].editBoxNumber = FINAL_ASSEMBLY_EDIT;
		i++;

		// 17 FINAL_PASS: TEST COMPLETE		
		testStep[i].lineOne = "UNIT PASSED ALL TESTS";
		testStep[i].lineTwo = "";
		testStep[i].lineThree = "Turn off power switch on DC950 and unplug power cord.";
		testStep[i].lineFour = "Remove DC950 from test system.";
		testStep[i].lineFive = "Then click ENTER to begin testing next unit.";
		testStep[i].stepType = MANUAL;
		testStep[i].enableENTER = TRUE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableRETRY = FALSE;
		testStep[i].enablePREVIOUS = FALSE;
		testStep[i].enableADMIN = FALSE;
		testStep[i].editBoxNumber = 0;
		i++;

		// 18 TEST COMPLETE FAILED			
		testStep[i].lineTwo = "DC950 FAILED - REMOVE FROM TEST SYSTEM";
		testStep[i].lineOne = "Turn off power switch on DC950, and unplug power cord."; 
		testStep[i].lineThree = "If lamp housing has been removed, reassemble it.";
		testStep[i].lineFour = "Remove DC950 from test system";
		testStep[i].lineFive = "Then click ENTER to begin testing next unit.";
		testStep[i].stepType = MANUAL;
		testStep[i].enableENTER = TRUE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableRETRY = FALSE;
		testStep[i].enablePREVIOUS = FALSE;	
		testStep[i].enableADMIN = FALSE;
		testStep[i].editBoxNumber = 0;
		i++;

		// 19 RETRY_TEST?: RETRY OR QUIT?		
		testStep[i].lineOne = "UNIT FAILED TEST";
		testStep[i].lineTwo = "";
		testStep[i].lineThree = "Click RETRY to run test again,";
		testStep[i].lineFour = "or click FAIL to end test and remove unit.";
		testStep[i].lineFive = "";
		testStep[i].stepType = MANUAL;
		testStep[i].enableENTER = FALSE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = TRUE;
		testStep[i].enableRETRY = TRUE;
		testStep[i].enablePREVIOUS = FALSE;	
		testStep[i].enableADMIN = FALSE;
		testStep[i].editBoxNumber = 0;
		i++;


		// 20 SYSTEM_FAILED	
		testStep[i].lineOne = "SYSTEM FAILURE";
		testStep[i].lineTwo = "";
		testStep[i].lineThree = "Check USB cables and make sure meter, Interface box,";
		testStep[i].lineFour = "and California Instruments supply are turned on.";
		testStep[i].lineFive = "Restart this program if necessary";
		testStep[i].stepType = MANUAL;
		testStep[i].enableENTER = TRUE;
		testStep[i].enablePASS = FALSE;
		testStep[i].enableFAIL = FALSE;
		testStep[i].enableRETRY = FALSE;
		testStep[i].enablePREVIOUS = FALSE;
		testStep[i].enableADMIN = TRUE;
		testStep[i].editBoxNumber = 0;
		return TRUE;
	}

	void TestApp::disableBarcodeScan() {
		ptrDialog->m_BarcodeEditBox.SendMessage(EM_SETREADONLY, 1, 0);
	}

	void TestApp::enablePartNumberScan() {		
		ptrDialog->m_PartNumberEditBox.SetWindowText((LPCTSTR)"");
		ptrDialog->m_PartNumberEditBox.SetFocus();
		ptrDialog->m_PartNumberEditBox.ShowCaret();
		ptrDialog->m_PartNumberEditBox.SendMessage(EM_SETREADONLY, 0, 0);
	}

	void TestApp::disablePartNumberScan() {
		ptrDialog->m_PartNumberEditBox.SendMessage(EM_SETREADONLY, 1, 0);
	}

	void TestApp::ClearLog(){
		strTestLog[0] = '\0';
		ptrDialog->m_EditBox_Log.SetWindowText((LPCTSTR)strTestLog);
		ptrDialog->pEdit->LineScroll (ptrDialog->pEdit->GetLineCount());
	}

	void TestApp::DisplayLog(char *newString){		
		strcat_s(strTestLog, MAXLOG, newString);
		ptrDialog->m_EditBox_Log.SetWindowText((LPCTSTR)strTestLog);
		ptrDialog->pEdit->LineScroll (ptrDialog->pEdit->GetLineCount());			
	}

	int TestApp::stringLength(char *ptrString)
	{
		if (ptrString == NULL) return (0);
		else return strlen (ptrString);
	}

	void TestApp::clearString(char *ptrString)
	{
		if (ptrString != NULL) ptrString[0] = '\0';
	}


	// Displays user instructions for each test, and enables/disables buttons as needed
	void TestApp::DisplayIntructions(int stepNumber, int status)
	{
		char strInstruct[MAXLINE] = "";
		
		if (lineOneEmpty) strcpy_s(strLineOne, MAXLINE, testStep[stepNumber].lineOne);
		if (lineTwoEmpty)	strcpy_s(strLineTwo, MAXLINE, testStep[stepNumber].lineTwo);
		if (lineThreeEmpty)	strcpy_s(strLineThree, MAXLINE, testStep[stepNumber].lineThree);
		if (lineFourEmpty)	strcpy_s(strLineFour, MAXLINE, testStep[stepNumber].lineFour);
		if (lineFiveEmpty)	strcpy_s(strLineFive, MAXLINE, testStep[stepNumber].lineFive);		
		
		strInstruct[0] = '\0';
		strcpy_s(strInstruct, MAXLINE, "");

		strcat_s(strInstruct, MAXLINE, strLineOne);
		strcat_s(strInstruct, MAXLINE, "\r\n");

		strcat_s(strInstruct, MAXLINE, strLineTwo);
		strcat_s(strInstruct, MAXLINE, "\r\n");

		strcat_s(strInstruct, MAXLINE, strLineThree);
		strcat_s(strInstruct, MAXLINE, "\r\n");

		strcat_s(strInstruct, MAXLINE, strLineFour);
		strcat_s(strInstruct, MAXLINE, "\r\n");

		strcat_s(strInstruct, MAXLINE, strLineFive);
		
		ptrDialog->m_EditBox_Instruct.SetFont(&MidFont, TRUE);
		ptrDialog->m_EditBox_Instruct.SetWindowText((LPCTSTR)strInstruct);
		

		// Enable or disable buttons depending on which step will execute next:
		ptrDialog->m_static_ButtonPass.EnableWindow(testStep[stepNumber].enablePASS);
		ptrDialog->m_static_ButtonFail.EnableWindow(testStep[stepNumber].enableFAIL);
		ptrDialog->m_static_ButtonEnter.EnableWindow(testStep[stepNumber].enableENTER);
		ptrDialog->m_static_ButtonPrevious.EnableWindow(testStep[stepNumber].enablePREVIOUS);
		ptrDialog->m_static_ButtonRetry.EnableWindow(testStep[stepNumber].enableRETRY);
		ptrDialog->m_static_ButtonAdmin1.EnableWindow(testStep[stepNumber].enableADMIN);
		ptrDialog->m_static_ButtonAdmin2.EnableWindow(testStep[stepNumber].enableADMIN);

		if (stepNumber == 1) enableBarcodeScan();
		else disableBarcodeScan();
		if (stepNumber == 2) enablePartNumberScan();		
		else disablePartNumberScan();

		strLineOne[0] = '\0';
		strLineTwo[0] = '\0';
		strLineThree[0] = '\0';
		strLineFour[0] = '\0';
		strLineFive[0] = '\0';

		lineOneEmpty = TRUE;
		lineTwoEmpty = TRUE;
		lineThreeEmpty = TRUE;
		lineFourEmpty = TRUE;
		lineFiveEmpty = TRUE;
				
		if (status == FAIL || status == SYSTEM_ERROR)
		{
			ptrDialog->m_EditBox_Instruct.SetBackColor(CYAN);
			ptrDialog->m_EditBox_Instruct.SetTextColor(RED);
		}
		else 
		{
			ptrDialog->m_EditBox_Instruct.SetBackColor(WHITE);
			ptrDialog->m_EditBox_Instruct.SetTextColor(BLACK);
		}		
	}

	
	void TestApp::ClearInstructBox()
	{
		clearString(strLineOne);
		clearString(strLineTwo);
		clearString(strLineThree);
		clearString(strLineFour);
		clearString(strLineFive);
		// strInstruct[0] = '\0';
		ptrDialog->m_EditBox_Instruct.SetWindowText((LPCTSTR)"");
		ptrDialog->pInstruct->LineScroll (ptrDialog->pInstruct->GetLineCount());
		// ptrDialog->m_EditBox_Instruct.SetBackColor(WHITE);
		// ptrDialog->m_EditBox_Instruct.SetTextColor(BLACK);
	}

	void TestApp::writeInstructionLine (int lineNumber, int lineLength, char *ptrText)
{
	switch (lineNumber)
	{
		case 1:	strcpy_s(strLineOne, MAXLINE, ptrText); 
				lineOneEmpty = FALSE;
				break;

		case 2:	strcpy_s(strLineTwo, MAXLINE, ptrText); 
				lineTwoEmpty = FALSE;
				break;

		case 3:	strcpy_s(strLineThree, MAXLINE, ptrText); 
				lineThreeEmpty = FALSE;
				break;

		case 4:	strcpy_s(strLineFour, MAXLINE, ptrText); 
				lineFourEmpty = FALSE;
				break;

		case 5:	strcpy_s(strLineFive, MAXLINE, ptrText); 
				lineFiveEmpty = FALSE;
				break;
		default:
			break;
	}
}

	void TestApp::enableBarcodeScan() {		
		ptrDialog->m_BarcodeEditBox.SetWindowText((LPCTSTR)"");	
		ptrDialog->m_PartNumberEditBox.SetWindowText((LPCTSTR)"");
		ptrDialog->m_BarcodeEditBox.SetFocus();
		ptrDialog->m_BarcodeEditBox.ShowCaret();
		ptrDialog->m_BarcodeEditBox.SendMessage(EM_SETREADONLY, 0, 0);
	}



	void TestApp::resetDisplays(BOOL enableTextBoxes) {
	int i;

		resetTestData();

		ClearLog();
		ptrDialog->m_BarcodeEditBox.SetWindowText((LPCTSTR)"");	
		ptrDialog->m_PartNumberEditBox.SetWindowText((LPCTSTR)"");		
		
		ptrDialog->m_static_ButtonHalt.EnableWindow(TRUE);

		ptrDialog->m_BarcodeEditBox.ShowWindow(TRUE);						
		ptrDialog->m_PartNumberEditBox.ShowWindow(TRUE);
		ptrDialog->m_static_BarcodeLabel.ShowWindow(TRUE);
		ptrDialog->m_static_PartNumber.ShowWindow(TRUE);	
						
		InitializeDisplayInstructions();

		if (enableTextBoxes)
		{
			ptrDialog->m_EditBox_Test1.EnableWindow(TRUE);
			ptrDialog->m_EditBox_Test2.EnableWindow(TRUE);
			ptrDialog->m_EditBox_Test3.EnableWindow(TRUE);
			ptrDialog->m_EditBox_Test4.EnableWindow(TRUE);
			ptrDialog->m_EditBox_Test5.EnableWindow(TRUE);
			ptrDialog->m_EditBox_Test6.EnableWindow(TRUE);
			ptrDialog->m_EditBox_Test7.EnableWindow(TRUE);
			ptrDialog->m_EditBox_Test8.EnableWindow(TRUE);
			ptrDialog->m_EditBox_Test9.EnableWindow(TRUE);
			ptrDialog->m_EditBox_Test10.EnableWindow(TRUE);
		}
		else
		{
			ptrDialog->m_EditBox_Test1.EnableWindow(FALSE);
			ptrDialog->m_EditBox_Test2.EnableWindow(FALSE);
			ptrDialog->m_EditBox_Test3.EnableWindow(FALSE);
			ptrDialog->m_EditBox_Test4.EnableWindow(FALSE);
			ptrDialog->m_EditBox_Test5.EnableWindow(FALSE);
			ptrDialog->m_EditBox_Test6.EnableWindow(FALSE);
			ptrDialog->m_EditBox_Test7.EnableWindow(FALSE);
			ptrDialog->m_EditBox_Test8.EnableWindow(FALSE);
			ptrDialog->m_EditBox_Test9.EnableWindow(FALSE);
			ptrDialog->m_EditBox_Test10.EnableWindow(FALSE);
		}
		for (i = 1; i <= LAST_TEST_EDIT_BOX; i++)  DisplayTestEditBox(i, NOT_DONE_YET);		
	}

