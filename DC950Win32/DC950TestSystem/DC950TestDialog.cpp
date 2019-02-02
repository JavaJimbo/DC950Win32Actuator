// COMMENTED
/*  Project: DC950TestSystem
*   DC950TestDialog.cpp - implementation file for main dialog box 
*
*	This is the top level file for the project.
*	Program execution begins by initializing the DC950TestDialog dialog box
*	and executing the OnInitDialog( ) routine below.
*	This routine opens the files with configuration information
*	with the serial port assignments, test limits, Admin password,
*	and also checks to make sure the test data spreadsheet can be located.
*
*	Once the main dialog box is initailized, the program remains idle until
*	the user clicks the ENTER button, and TestHandler() is called below. 
*	TestHandler( ) is the top level function
*	which calls all routines used in this program after initialization at startup.
*
*	11-26-17 JBS: Converted *ptrDialog to global, removed from all headers.
*	12-2-17 JBS: Converted to Win32 bits for PC in work cell.
*	12-3-17 JBS: Added Spectrometer diagnostics dialog box
*	12-10-17 JBS: All calculations for irradiance, color temp etc are completed
*	12-11-17 JBS: DC950IR P/N: 660001661561
*	12-13-17 JBS: Prepared system for demo.
*	12-14-17 JBS: Fixed pass/fail display bug
*	12-14-17 JBS: Eliminated color temp for filter test, 
*	 modified ADMIN to adjust MIN and MAX peak amplitude for filter test.
*    Make sure that timer doesn't run past pauses for manual steps.
*	12-15-17 JBS: Fixed bug that crashed the program when the spreadsheet was being created.
*	12-16-17 JBS: Eliminated unused edit boxes from main dialog box.
*	12-16-17 JBS: Added part number scan.
*	12-19-17 JBS: Debugging and cleanup at Setra. Part number match added. Spreadsheet location hard coded.
*	12-19-17 JBS: Added integration time, number of scans and averages to Admin 2.
*	12-20-17 JBS: Added stray light correction.
*	12-22-17 JBS: Added edit box for setting spreadsheet name and path in Admin 1.
*	12-23-17 JBS: Fixed strdup() memory leak in loadINIpartNumbers(), LoadExcelFilenames(), LoadScanData()
*	12-28-17 JBS: Includes SpreadsheetIsOpen() and CheckAndCreateDirectory() in Initialize.cpp
*	12-28-17 JBS: Added minor enhancements and bug fixes at Setra, 
*					including spectrometer test label next to part number edit box.
*					HALT Button commented out.
*	12-29-17 JBS:	Added m_static_ButtonEnter.EnableWindow(TRUE) to StopTimer() but haven't tested it yet.
*	01-01-17 JBS:	Substituted INSTRUCT Edit box for LINE 1,2,3. Low level routines can write to box if a system
*					error occurs. Otherwise lines 1,2,3 
*	01-03-17 JBS:	Added SYSTEM_ERROR, SYSTEM_HALT, and RETRY_TEST to Execute().
*					Instruction display box has six lines. 
*					1) Read testStep[] text and strLineOne[] etc. Any text written to strLineOne[] takes precedence.
*					2) After copying to instruct box, clear strLineOne[].
*	01-04-17 JBS:	Eliminated DisplayMessageBox(), substituted INSTRUCTION WINDOW
*	01-05-17 JBS:	Cleaned up new code at home, implemented ENTER_CLICK, PREVIOUS_CLICK, 
*					FAIL_CLICK, PASS_CLICK, HALT_CLICK, RETRY_CLICK, TIMER_INT
*					Tested code at home in demo mode.
*					DEBUGGED at Setra test stand.
*	01-08-17 JBS:	Changed FINAL_PASS to FINAL_ASSEMBLY at home. Cleaned up instruction text.
*	01-08-17 JBS:	Debugged at Setra moved INI files to internal folder and working spreadsheet and backup spreadsheet
*	01-10-17 JBS:	Implements backup file creation and prevents corrupted spreadsheets.
*	01-11-18 JBS:	At Setra, added Admin password change. Renamed INI subfolder to "Startup"
*	01-12-18 JBS:	Home: fully implemented and debugged dated spreadsheet backup scheme
*	01-14-18 JBS:	Added minimize and File/Exit features. Lots of debugging on Primary and Secondary file backup stuff.
*	01-15-18 JBS:	More debugging at home - worked on Execute( ) Startup folder eliminated
*	01-15-18 JBS:	Debugged at Setra
*	01-16-18 JBS:	Added FAULT signal check at home 
*	01-16-18 JBS:	Debugged LAMP FAULT signal at Setra. Questions about TTL input test.
*					Cleaned up system error messages at Setra test stand. Added instruction text.
*	01-23-18 JBS:	Added MinRemoteVoltLimit[] and MaxRemoteVoltLimit[] for remote voltage test limits, modified RunPowerSupplyTest(),
*					and added text boxes to modify limits in AdminDialog2.cpp
*	01-24-18 JBS:	Debugged at Setra, minor tweaks. Created Release version of code.
*	02-15-18 JBS:	Add comments
*	02-28-18 JBS:	
*	03-21-18 JBS:	Commented version, with existing Excel routines.
*	03-22-18 JBS:	Swapped in new spreadsheet read/write routines.
*	03-24-18 JBS:	
*	03-26-18 JBS:	Boxborough - Added MSExcel.cpp - RELEASE VERSION
*	03-28-18 JBS:	Warwick: fixed bug in AdminDialog1: ptrExcelFilename filename pointer gets set only when user clicks OK
*					Also chaged HALT button color to RED, cleaned up stdafx.h, and eliminated unused routines from MSExcel.cpp
*					Also created SaveTestDataToSpreadsheet()
*	04-03-18 JBS:	Released at Setra.
*   01-18-19 JBS:   Need to understand why CheckFaultSignal() is passing when it should fail if no unit is plugged in.
*					Add Diagnostic feature allowing user to actuate actuator.
*	02-02-19 JBS: 
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
#include "ReadOnlyEdit.h"
// #include "PasswordEntry_Dlg.h"	
#include "PasswordEntryDlg.h"	
#include <ole2.h> // OLE2 Definitions
#include "MSExcel.h"

LOGFONT lf;
CFont font1;

bool StartUpFlag = true;
bool EnableManualActuatorControl = true;
BOOL runFilterActuatorTest = FALSE;
BOOL copySpreadsheet = TRUE;
extern int HiPotStatus, GroundBondStatus, FinalAssemblyStatus;
BOOL testOLE();

enum CALLERS {
	ENTER_CLICK = 0,
	PREVIOUS_CLICK,
	FAIL_CLICK,
	PASS_CLICK,
	HALT_CLICK,
	RETRY_CLICK,
	TIMER_INT
};

char AdminPassword[MAXPASSWORD] = "Setra";

const char *SpreadsheetTemplateFilename = "DC950Template.xls";  
const char *CurrentDataFilename = "C:\\Temp\\CurrentTestData.xls";  
const char *CurrentDataPathname = "C:\\Temp";  // MUST BE SAME AS ABOVE PATH!
char ExcelBackupFilename[MAXFILENAME] = "DC950BackupData.xls";
char *ptrExcelFilename = "DC950BackupData.xls";  //$$$$
char ExcelSecondaryFilename[MAXFILENAME] = "DC950TestData[00-00-00].xls";
char ExcelPrimaryFilename[MAXFILENAME] = "DC950TestData.xls";
 

DATAstring arrPartNumbers[MAX_PART_NUMBERS] = {"660001661561"};

int totalPartNumbers = 1;

extern DeviceConfigType* l_pDeviceData;
extern uint16	m_StopPixel;

int stepNumber = 0;

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

TestApp MyTestApp;	
CMSExcel MyExcelApp;

HANDLE m_timerHandle = NULL;
CTestDialog *ptrDialog;

// Dialog constructor
CTestDialog::CTestDialog(CWnd* pParent /*=NULL*/)	: CDialog(CTestDialog::IDD, pParent)
{

}


// Dialog deconstructor: shut down system and copy test data to spreadsheet one last time:
CTestDialog::~CTestDialog()
{	
	MyTestApp.ShutDown(copySpreadsheet);		
}

void CTestDialog::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);				
	DDX_Control(pDX, IDC_BUTTON_ENTER, m_static_ButtonEnter); 
	DDX_Control(pDX, IDC_BUTTON_PASS,  m_static_ButtonPass);
	DDX_Control(pDX, IDC_BUTTON_FAIL, m_static_ButtonFail);
		
	DDX_Control(pDX, IDC_BUTTON_PREVIOUS, m_static_ButtonPrevious);
	DDX_Control(pDX, IDC_BUTTON_HALT, m_static_ButtonHalt);
	DDX_Control(pDX, IDC_EDIT_BARCODE, m_BarcodeEditBox);
	DDX_Control(pDX, IDC_EDIT_PARTNUMBER, m_PartNumberEditBox);
	DDX_Control(pDX, IDC_EDIT_TEST1, m_EditBox_Test1);
	DDX_Control(pDX, IDC_EDIT_TEST2, m_EditBox_Test2);
	DDX_Control(pDX, IDC_EDIT_TEST3, m_EditBox_Test3);
	DDX_Control(pDX, IDC_EDIT_TEST4, m_EditBox_Test4);
	DDX_Control(pDX, IDC_EDIT_TEST5, m_EditBox_Test5);
	DDX_Control(pDX, IDC_EDIT_TEST6, m_EditBox_Test6);
	DDX_Control(pDX, IDC_EDIT_TEST7, m_EditBox_Test7);
	DDX_Control(pDX, IDC_EDIT_TEST8, m_EditBox_Test8);
	DDX_Control(pDX, IDC_EDIT_TEST10, m_EditBox_Test9);
	DDX_Control(pDX, IDC_EDIT_TEST11, m_EditBox_Test10);
	DDX_Control(pDX, IDC_EDIT_INSTRUCT, m_EditBox_Instruct);

	DDX_Control(pDX, IDC_EDIT_LOG, m_EditBox_Log);	
	DDX_Control(pDX, IDC_STATIC_SN, m_static_BarcodeLabel);	
	DDX_Control(pDX, IDC_STATIC_FILENAME, m_static_SpreadsheetFilename);
	DDX_Control(pDX, IDC_STATIC_PARTNO, m_static_PartNumber);
	DDX_Control(pDX, IDC_BUTTON_RETRY, m_static_ButtonRetry);
			
	DDX_Control(pDX, IDC_BUTTON_ADMIN1,  m_static_ButtonAdmin1);	
	DDX_Control(pDX, IDC_BUTTON_ADMIN2,  m_static_ButtonAdmin2);
}

BEGIN_MESSAGE_MAP(CTestDialog, CDialog)
	ON_WM_SYSCOMMAND()
	
	ON_BN_CLICKED(IDC_BUTTON_ENTER, &CTestDialog::OnClickedButtonEnter)
	ON_BN_CLICKED(IDC_BUTTON_HALT, &CTestDialog::OnClickedButtonHalt)
	ON_BN_CLICKED(IDC_BUTTON_PREVIOUS, &CTestDialog::OnClickedButtonPrevious)
	ON_BN_CLICKED(IDC_BUTTON_PASS, &CTestDialog::OnClickedButtonPass)
	ON_BN_CLICKED(IDC_BUTTON_FAIL, &CTestDialog::OnClickedButtonFail)
	ON_BN_CLICKED(IDC_BUTTON_RETRY, &CTestDialog::OnClickedButtonRetry)		
	ON_BN_CLICKED(IDC_BUTTON_ADMIN1, &CTestDialog::OnClickedButtonAdmin1)
	ON_WM_ERASEBKGND()
	ON_BN_CLICKED(IDC_BUTTON_ADMIN2, &CTestDialog::OnClickedButtonAdmin2)
	ON_COMMAND(ID_FILE_EXIT, &CTestDialog::OnClickFileExit)
	ON_COMMAND(ID_ACTUATOR_OPEN, &CTestDialog::OnClickActuatorOPEN)
	ON_COMMAND(ID_ACTUATOR_CLOSED, &CTestDialog::OnClickActuatorCLOSED)
	END_MESSAGE_MAP()


// This routine prevents the ENTER or ESCAPE keys from closing the main dialog box
// and thereby shutting down the program. It is necessary because 
// the barcode scanner otherwise closes the dialog box 
// when it sends a carriage return after the serial number.
BOOL CTestDialog::PreTranslateMessage(MSG* pMsg)
{
	if (pMsg->message == WM_KEYDOWN)
	{
		if (pMsg->wParam == VK_RETURN || pMsg->wParam == VK_ESCAPE)
		{
			return TRUE;                // Do not process further
		}
	}

	return CWnd::PreTranslateMessage(pMsg);
}


// Called when Admin2 is clicked
void CTestDialog::OnClickedButtonAdmin2()
{		
	if (enableAdmin) m_SpecDiag.DoModal();
	else {
		CString s = (CString)"";  
		m_PED.DoModal(&s);	
		if(s != AdminPassword)	
			AfxMessageBox((LPCTSTR) "Incorrect password entered!\r\nPlease see system adminstrator for access to administrator utilities.");	
		else
		{
			enableAdmin = TRUE;
			m_SpecDiag.DoModal();
		}
	}	
}

// Called when Admin1 is clicked
void CTestDialog::OnClickedButtonAdmin1()
{	
	if (enableAdmin) m_Admin1.DoModal();
	else {
		CString s = (CString)"";
		m_PED.DoModal(&s);	
		if(s != AdminPassword)	
			AfxMessageBox((LPCTSTR)"Incorrect password entered!\r\nPlease see system adminstrator for access to administrator utilities.");	
		else
		{
			enableAdmin = TRUE;
			m_Admin1.DoModal();
		}
	}
	char filenameText[MAXFILENAME+16];
	sprintf_s(filenameText, MAXFILENAME+16, "Spreadsheet: %s", ptrExcelFilename);
	m_static_SpreadsheetFilename.SetWindowText(filenameText);	
}

// Initializes main dialog box at startup
BOOL CTestDialog::OnInitDialog()
{		
	int length = 0;
	int filenameStatus = 0;
	int dummy = 0;
	BOOL loadError = FALSE;
	CDialog::OnInitDialog();	

	

	// Load file "PasswordFile.txt" to get Admin password
	if (!MyTestApp.LoadAdminPassword())
	{
		MessageBox((LPCTSTR)"Could not load password\r\nSee system administrator", (LPCTSTR)"PASSWORD FILE ERROR", MB_ICONWARNING | MB_OK | MB_DEFBUTTON2);
		loadError = true;
	}
	 
	enableAdmin = FALSE;

	// The ptrDialog pointer is initialized here to "this" for the main dialog box
	// It permits functions to access the text boxes and controls when displaying text.
	ptrDialog = this;
	stepNumber = 0;
		
	runFilterActuatorTest = FALSE;				

	// Initialize fonts to use when needed.
	// Three different text sizes are used: 
	// BigFont, MidFont, and SmallFont
	pEdit = (CEdit*) GetDlgItem(IDC_EDIT_LOG);
	pInstruct = (CReadOnlyEdit*) GetDlgItem(IDC_EDIT_INSTRUCT);
	MyTestApp.InitializeFonts();
	
	stepNumber = 0;		// Initialize stepNumber to start the test sequence from the beginning
	MyTestApp.InitializeDisplayInstructions();	// Initialize the user instruction box
	MyTestApp.DisplayIntructions(stepNumber, NOT_DONE_YET);
		
	m_EditBox_Log.SetLimitText(MAXLOG);   // Initialize the test log box so it can display tons of test data
	m_EditBox_Instruct.SetLimitText(MAXLOG);
	MyTestApp.InitializeTestEditBoxes();  // Initialize Edit boxes which will display pass/fail status of each test
	
	// Open configuration files and initialize test parameters:
	if (!MyTestApp.loadINIfileBinary())	// This file has test limits and serial port assignments
	{
		MessageBox((LPCTSTR)"Could not load StartupFile.bin\r\nSee system administrator", (LPCTSTR)"FILE LOAD ERROR", MB_ICONWARNING | MB_OK | MB_DEFBUTTON2);
		loadError = true;
	}
	if (!MyTestApp.loadINIpartNumbers())	// This file has the part numbers for the filter actuator models
	{
		MessageBox((LPCTSTR)"Could not load PartNumber.txt\r\nSee system administrator", (LPCTSTR)"FILE LOAD ERROR", MB_ICONWARNING | MB_OK | MB_DEFBUTTON2);
		loadError = true;
	}

	MyTestApp.loadConfigFile();		// This initializes the spectrometer parameters
	
	
	// Fetch the name and pathof the Primary spreadsheet which stores test data for all units.
	// If there was a problem locating the Primary during a previous test run,
	// then a Secondary spreadsheet is used instead.
	filenameStatus = MyTestApp.LoadExcelFilenames(ExcelPrimaryFilename, ExcelSecondaryFilename);
	if (filenameStatus == FILE_LOAD_ERROR)
	{
		MessageBox((LPCTSTR)"Could not load OutFile.txt\r\nSee system administrator", (LPCTSTR)"FILE LOAD ERROR", MB_ICONWARNING | MB_OK | MB_DEFBUTTON2);
		m_static_ButtonEnter.EnableWindow(FALSE);
		return false;			
	}		
	else if (filenameStatus == USING_PRIMARY_SPREADSHEET) ptrExcelFilename = ExcelPrimaryFilename;
	else if (filenameStatus == USING_SECONDARY_SPREADSHEET) 
		ptrExcelFilename = ExcelSecondaryFilename;
	else ptrExcelFilename = ExcelBackupFilename;
		
	if (!MyTestApp.InitializeSpreadsheets())
	{
		MessageBox((LPCTSTR)"Could not load spreadsheet template\r\nSee system administrator", (LPCTSTR)"SPREADSHEET ERROR", MB_ICONWARNING | MB_OK | MB_DEFBUTTON2);	
		loadError = true;
	}	

	if (loadError) {
		m_static_ButtonEnter.EnableWindow(FALSE);
		return false;
	}
	else {
		m_static_ButtonEnter.EnableWindow(TRUE);	
		return true;
	}

	// $$$$ New: enabled actuator Menu buttons
	//m_static_ButtonMenuActuatorON.EnableWindow(TRUE);
	//m_static_ButtonMenuActuatorOFF.EnableWindow(TRUE);
}





// TestHandler()- the main top level routine for the entire program.
// All test sequences and other features are controlled by TestHandler
// with the exception of the Admin dialog boxes.	
//
// MyTestApp.Execute() is used below to run each test.
// Some tests are performed manually by the operator,
// in which case TestHandler() is called when a pushbutton is clicked:
// either ENTER, PASS, or FAIL.
// Otherwise, for automated test sequences,
// TestHandler() is called by the timer routine timerHandler() 
//
// The global variable stepNumber keeps track of what test is performed.

void CTestDialog::TestHandler(int caller)
{	
static int previousStepNumber = 0;
	
	// Store current stepNumber here in case a test needs to be repeated later
	if (stepNumber != RETRY_TEST) 
		previousStepNumber = stepNumber;

	// If TestHandler was called by the user clicking the PREVIOUS button,
	// decrement stepNumber here to go backward to the previous test.
	// Then display instructions for that test below:
	if (caller == PREVIOUS_CLICK)
	{
		int  stepType, boxNum;
		do 
		{
			if (stepNumber <= BARCODE_SCAN) break;
			stepNumber--;
			boxNum = MyTestApp.testStep[stepNumber].editBoxNumber;
			if (boxNum) MyTestApp.DisplayTestEditBox(boxNum, NOT_DONE_YET);
			stepType = MyTestApp.testStep[stepNumber].stepType;
		} while (stepType == AUTO);

		if (stepNumber == 1)
		{
			MyTestApp.resetDisplays(TRUE); 
			MyTestApp.enableBarcodeScan();
		}	
		else MyTestApp.disableBarcodeScan();

		if (stepNumber == 2)
		{		
			MyTestApp.enablePartNumberScan();
		}
		else MyTestApp.disablePartNumberScan();
		
		MyTestApp.DisplayIntructions(stepNumber, NOT_DONE_YET);	

		return;
	} 
	// If TestHandler was called by the user clicking the FAIL button, process that here:
	else if (caller == FAIL_CLICK)	
	{		
		if (stepNumber == HI_POT_TEST)
			HiPotStatus = FAIL;
		else if (stepNumber == GROUND_BOND_TEST)
			GroundBondStatus = FAIL;
		else if (stepNumber == FINAL_ASSEMBLY)
			FinalAssemblyStatus = FAIL;
		MyTestApp.Execute(stepNumber);

		if (stepNumber == HI_POT_TEST)
			stepNumber = RETRY_TEST;
		else if (stepNumber == GROUND_BOND_TEST)
			stepNumber = RETRY_TEST;
		else stepNumber = FINAL_FAIL;

		MyTestApp.DisplayIntructions(stepNumber, FAIL);
		return;
	}
	// If TestHandler was called by the user clicking the RETRY button, process that here:
	else if (caller == RETRY_CLICK) 	
	{
		stepNumber = previousStepNumber;
		if (stepNumber > DARK_SCAN_SETUP) 
		{
			stepNumber = DARK_SCAN_SETUP;
			MyTestApp.DisplayTestEditBox(ACTUATOR_EDIT, NOT_DONE_YET);
			MyTestApp.DisplayTestEditBox(FILTER_ON_EDIT, NOT_DONE_YET);
			MyTestApp.DisplayTestEditBox(FILTER_OFF_EDIT, NOT_DONE_YET);
		}		  
		else if (stepNumber > REMOTE_TEST_SETUP) 
		{
			stepNumber = REMOTE_TEST_SETUP;
			MyTestApp.DisplayTestEditBox(REMOTE_EDIT, NOT_DONE_YET);
			MyTestApp.DisplayTestEditBox(AC_SWEEP_EDIT, NOT_DONE_YET);
		}
		else 
		{
			int	EditBoxNum = MyTestApp.testStep[stepNumber].editBoxNumber;
			if (EditBoxNum) MyTestApp.DisplayTestEditBox(EditBoxNum, NOT_DONE_YET);
		}
		MyTestApp.DisplayIntructions(stepNumber, NOT_DONE_YET);
		return;
	}
	// If TestHandler was called by the user clicking the PASS button, process that here:
	else if (caller == PASS_CLICK)
	{
		if (stepNumber == HI_POT_TEST)
			HiPotStatus = PASS;
		else if (stepNumber == GROUND_BOND_TEST)
			GroundBondStatus = PASS;
		else if (stepNumber == FINAL_ASSEMBLY)
			FinalAssemblyStatus = PASS;
	}
	
	// Now run current test here. 
	// Execute() will also diplay test results and record them in the spreadsheet:
	int testResult = MyTestApp.Execute(stepNumber);	
	
	if (testResult != NOT_DONE_YET) {		
		StopTimer();	
		MyTestApp.resetSubStepNumber();

		if (testResult == SYSTEM_ERROR)
		{
			if (stepNumber > 0) 
			{
				stepNumber = SYSTEM_FAILED;							
				#ifndef NOPOWERSUPPLY
					MyTestApp.SetPowerSupplyVoltage(0, 60);
				#endif
				#ifndef NOSERIAL
					MyTestApp.closeAllSerialPorts();
				#endif
			}
		}

		// If unit passed the current test, then advance to next step.		
		// If unit isn't a filter actuator model, 
		// and the STANDARD test sequence has been completed,
		// then skip the ACTUATOR and SPECTROMETER tests and go to FINAL PASS:
		else if (testResult == PASS)
		{			
			if (stepNumber == FINAL_PASS)
				stepNumber = 1;
			else stepNumber++; 				
			if (runFilterActuatorTest == FALSE){
				if (stepNumber == END_STANDARD_UNIT_TESTS)
					stepNumber = FINAL_ASSEMBLY;					
			}				
		}
		// Otherwise if DC950 just failed a test,
		// not a barcode or part number scan,		
		// give user option to RETRY,
		// or if unit keeps failing a retries,
		// then go to FINAL FAIL.
		else if (stepNumber > 2) 
		{								
			if (stepNumber == RETRY_TEST)
				stepNumber = FINAL_FAIL;
			else if (stepNumber == FINAL_FAIL)	
			{
				stepNumber = 1;
				testResult = NOT_DONE_YET;
			}
			// A system failure is a special case.
			// NOT DONE YET is used here instead of FAIL
			// because the malfunction is with the system, not the DC950:
			else if (stepNumber == SYSTEM_FAILED) 
			{
				stepNumber = 0;
				testResult = NOT_DONE_YET;
			}
			else stepNumber = RETRY_TEST;				
		}
	}
	
	// If this is the start of an automated sequence
	// or if sequence is already in progress, start AUTO timer.
	// If next step is manual, make sure timer is off:				
	if (MyTestApp.testStep[stepNumber].stepType == AUTO) 
		StartTimer();
	else {
		StopTimer();
		MyTestApp.resetSubStepNumber(); 
	}

	// Now display user instructions to perform the  next step.
	// If testResult is FAIL, then instruction text will turn RED
	// and background will be CYAN:
	MyTestApp.DisplayIntructions(stepNumber, testResult);
}

// If user clicks PASS button, run TestHandler( ) with PASS_CLICK option:
void CTestDialog::OnClickedButtonPass()
{
	m_static_ButtonPass.EnableWindow(FALSE);
	m_static_ButtonFail.EnableWindow(FALSE);
	TestHandler(PASS_CLICK);
}

// If user clicks FAIL button, run TestHandler( ) with FAIL_CLICK option:
void CTestDialog::OnClickedButtonFail()
{
	m_static_ButtonPass.EnableWindow(FALSE);
	m_static_ButtonFail.EnableWindow(FALSE);
	TestHandler(FAIL_CLICK);
}

// If user clicks ENTER button, run TestHandler( ) with ENTER_CLICK option:
void CTestDialog::OnClickedButtonEnter() 
{	
static BOOL startup = TRUE;

	// The first time the ENTER button is clicked, the HALT button is turned RED here.
	// This is purely for cosmetic reasons - it otherwise looks weird before it gets enabled.
	if (startup)
	{
		EnableManualActuatorControl = false;
		startup = FALSE;		
		// All of the following code is necessary just to turn the HALT button RED:
		memset(&lf, 0, sizeof(LOGFONT));    // clear out structure 
		lf.lfHeight = 20;   // Set pixel size
		lf.lfWeight = FW_BOLD;    
		_tcsncpy_s(lf.lfFaceName, LF_FACESIZE, _T("Arial"), 20);  // Set font
		font1.CreateFontIndirect(&lf);  // create the font
		m_static_ButtonHalt.SetFont(&font1);	
	
		m_static_ButtonHalt.m_nFlatStyle = CMFCButton::BUTTONSTYLE_3D;//required for flatering and use bg color	
		m_static_ButtonHalt.m_bTransparent = false;//reg for use bg color
		m_static_ButtonHalt.SetFaceColor( RGB(255, 0, 0), true);
		m_static_ButtonHalt.SetTextColor( RGB(0, 0, 0));		
	}

	// For normal operation, the ENTER button is temporarily disabled whenever it is clicked.
	// Otherwise, TestHandler() might run twice, which could be disasterous.
	m_static_ButtonEnter.EnableWindow(FALSE);
	TestHandler(ENTER_CLICK);

}

// If user clicks RETRY button, run TestHandler( ) with RETRY_CLICK option:
void CTestDialog::OnClickedButtonRetry()
{	
	TestHandler(RETRY_CLICK);
}


// If user clicks PREVIOUS button, run TestHandler( ) with PREVIOUS_CLICK option:
void CTestDialog::OnClickedButtonPrevious()
{	
	TestHandler(PREVIOUS_CLICK);
}

// The next four routines implement the AUTO timer for running automated test sequences.
// If a test sequence is in progress and the timeout interval has ellapsed,
// this CALLBACK executes which calls timerHandler( ) below which in turn calls
// TestHandler() to run the next step in the test sequence: 
void CALLBACK TimerProc(CTestDialog *lpParametar, BOOLEAN TimerOrWaitFired)
{
	// This is used only to call QueueTimerHandler
	// Typically, this function is static member of CTimersDlg	
	CTestDialog *ptrTimer;
	ptrTimer = lpParametar;
	ptrTimer->timerHandler();			
}
// See above  notes 
void CTestDialog::timerHandler() {
	if (MyTestApp.testStep[stepNumber].stepType == AUTO) 
		TestHandler(TIMER_INT);
	else StopTimer();	
}

// This starts the AUTO timer when an automated test sequence is being started
void CTestDialog::StartTimer() {
	
	if (m_timerHandle == NULL) {
		DWORD elapsedTime = 3000;
		BOOL success = ::CreateTimerQueueTimer(  // Create new timer
			&m_timerHandle,
			NULL,
			(WAITORTIMERCALLBACK)TimerProc,
			this,
			0,
			elapsedTime,
			WT_EXECUTEINTIMERTHREAD);
	}	
}

// This stops the AUTO timer when an automated test sequence is completed or interrupted.
void CTestDialog::StopTimer() {	
	if (m_timerHandle != NULL) {
		DeleteTimerQueueTimer(NULL, m_timerHandle, NULL);
		m_timerHandle = NULL;
	}
	// MyTestApp.resetSubStepNumber();	
}


// If HALT button is clicked, shut doiwn everything and close program:
// Also test data is saved in the spreadsheet one last time:
void CTestDialog::OnClickedButtonHalt()
{
	m_static_ButtonHalt.EnableWindow(FALSE);
	MyTestApp.SetPowerSupplyVoltage(0, 60);
	MyTestApp.ShutDown(copySpreadsheet);
	exit(0);
}

// If user clicks on FILE/EXIT on the MENU toolbar,
// this routine shuts the system down and closes the program.
// It also saves the test data in the spreadsheet one last time:
void CTestDialog::OnClickFileExit()
{
	MyTestApp.SetPowerSupplyVoltage(0, 60);
	MyTestApp.ShutDown(copySpreadsheet);
	exit(0);
}


void CTestDialog::OnClickActuatorOPEN()
{
	int testResult;
	if (EnableManualActuatorControl)
	{
		m_static_ButtonEnter.EnableWindow(FALSE);
		m_static_ButtonAdmin1.EnableWindow(FALSE);
		m_static_ButtonAdmin2.EnableWindow(FALSE);
		
		if (StartUpFlag)
		{			
			if (MyTestApp.InitializeSystem(false)) StartUpFlag = false;
			MyTestApp.SetPowerSupplyVoltage(120, 60);
		}

		if (!StartUpFlag) 
		{
			MyTestApp.SetInterfaceBoardActuatorOutput(FILTER_OPEN, &testResult);		
			m_EditBox_Instruct.SetWindowText((LPCTSTR) "ACTUATOR OPEN.\r\n");
		}
	}
	else MessageBox((LPCTSTR)"If you wish to use this feature, close program and restart", (LPCTSTR)"ACTUATOR CONTROLLER DISABLED", MB_ICONWARNING | MB_OK | MB_DEFBUTTON2);
}
	
void CTestDialog::OnClickActuatorCLOSED()
{
	int testResult;
	if (EnableManualActuatorControl)
	{
		m_static_ButtonEnter.EnableWindow(FALSE);
		m_static_ButtonAdmin1.EnableWindow(FALSE);
		m_static_ButtonAdmin2.EnableWindow(FALSE);

		if (StartUpFlag)
		{			
			if (MyTestApp.InitializeSystem(false)) StartUpFlag = false;
			MyTestApp.SetPowerSupplyVoltage(120, 60);
		}

		if (!StartUpFlag) 
		{
			MyTestApp.SetInterfaceBoardActuatorOutput(FILTER_CLOSED, &testResult);		
			m_EditBox_Instruct.SetWindowText((LPCTSTR) "ACTUATOR CLOSED");
		}
	}
	else MessageBox((LPCTSTR)"If you wish to use this feature, close program and restart", (LPCTSTR)"ACTUATOR CONTROLLER DISABLED", MB_ICONWARNING | MB_OK | MB_DEFBUTTON2);
}


