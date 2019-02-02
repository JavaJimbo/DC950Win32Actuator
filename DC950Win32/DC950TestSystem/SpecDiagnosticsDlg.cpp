// SpecDiagnosticsDlg.cpp : implementation file for SpecDiagnostics Settings dialog box

#include "stdafx.h"
#include "Definitions.h"
#include <windows.h>
#include <fstream>
#include <sstream>
#include <string.h>
#include "SpecDiagnosticsDlg.h"
#include "TestApp.h"
#include "avaspec.h"
#include "ReadOnlyEdit.h"
#include "TestApp.h"
#include "irrad.h"

// const int MAX_PIXELS = 3648;


// SPECTROMETER GLOBAL VARIABLES
extern DeviceConfigType* l_pDeviceData;
extern int m_NrDevices;
extern long handleSpectrometer;
extern unsigned int numberOfPixels;
extern double Lambda[];
extern float currentIntegrationTime;
extern uint16	m_StartPixel;
extern uint16	m_StopPixel;
extern uint16	m_NumberOfScans;
extern BOOL enableLinearCorrection;
extern BOOL enableStrayLightCorrection;

HANDLE m_specTimerHandle = NULL;
char arrSpectrometerLog[MAXLOG] = {'\0'};
int intSeconds = 0;
extern TestApp MyTestApp;
BOOL filterState = OFF;
BOOL darkDataReady = FALSE;
BOOL scopeDataReady = FALSE;
BOOL useDarkData = TRUE;
float *ptrData;
double l_pSpectrum[PIXEL_ARRAY_SIZE];
double l_pCorrSpectrum[PIXEL_ARRAY_SIZE];

int intScanType = FILTER_OPEN; // FILTER_CLOSED, DARK_SCAN

void OnRunStatic();
extern TestApp MyTestApp;	

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// SpecDiagnostics dialog

SpecDiagnostics::SpecDiagnostics(CWnd* pParent /*=NULL*/) : CDialog(SpecDiagnostics::IDD, pParent)
{
	
}

SpecDiagnostics::~SpecDiagnostics()
{

}


void SpecDiagnostics::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);	
	DDX_Control(pDX, IDC_EDIT_SPECTROMETER, m_EditBox_Spectrometer);  
	DDX_Control(pDX, IDC_BUTTON_RUN_TEST, m_static_ButtonTestSpectrometer);
	DDX_Control(pDX, IDC_BUTTON_SCAN, m_static_ButtonScan);
	DDX_Control(pDX, IDC_BUTTON_DARK_SCAN, m_static_ButtonDarkScan);
	DDX_Control(pDX, IDC_BUTTON_ACTUATOR, m_static_ButtonActuator);
	DDX_Control(pDX, IDC_CHECK_SLC, m_static_checkStrayLightCorrection);
}

BEGIN_MESSAGE_MAP(SpecDiagnostics, CDialog)
	ON_WM_CLOSE()		
	ON_BN_CLICKED(IDOK, &SpecDiagnostics::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &SpecDiagnostics::OnBnClickedCancel)
	ON_BN_CLICKED(IDC_BUTTON_RUN_TEST, &SpecDiagnostics::OnBnClickedButtonRunTest)
	ON_BN_CLICKED(IDC_BUTTON_SCAN, &SpecDiagnostics::OnBnClickedButtonScan)
	ON_BN_CLICKED(IDC_BUTTON_DARK_SCAN, &SpecDiagnostics::OnBnClickedButtonDarkScan)
	ON_BN_CLICKED(IDC_BUTTON_ACTUATOR, &SpecDiagnostics::OnBnClickedButtonActuator)
	ON_BN_CLICKED(IDC_CHECK_SLC, &SpecDiagnostics::OnCheckedBoxSLC)
	ON_BN_CLICKED(IDC_BUTTON_HALT_TEST, &OnBnClickedButtonHaltTest)	
	ON_BN_CLICKED(IDC_BUTTON_LOAD_EEPROM, &SpecDiagnostics::OnBnClickedButtonLoadEeprom)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// SpecDiagnostics message handlers

//int SpecDiagnostics::DoModal() 
//{
	// TODO: Add your specialized code here and/or call the base class
//	return (int) CDialog::DoModal();
//}

void SpecDiagnostics::OnClose() 
{
	// TODO: Add your message handler code here and/or call default
	CDialog::OnClose();
}

void SpecDiagnostics::OnNMThemeChangedScrollComInterface(NMHDR *pNMHDR, LRESULT *pResult)
{
	// This feature requires Windows XP or greater.
	// The symbol _WIN32_WINNT must be >= 0x0501.
	// TODO: Add your control notification handler code here
	*pResult = 0;
}


void SpecDiagnostics::OnDeltaposSpin1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMUPDOWN pNMUpDown = reinterpret_cast<LPNMUPDOWN>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0;
}

void SpecDiagnostics::OnPaint() 
{	
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0); 

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		// dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
	
}

void SpecDiagnostics::OnBnClickedOk()
{
	int UserResponse = MessageBox((LPCTSTR)"Click OK to save new settings", (LPCTSTR)"SAVE NEW SETTINGS?", MB_ICONWARNING | MB_OKCANCEL | MB_DEFBUTTON2);
	if (UserResponse == IDOK) {	
		CDialog::OnOK();
	}
}

void SpecDiagnostics::OnBnClickedCancel()
{
	int UserResponse = MessageBox((LPCTSTR)"Click OK to quit without saving new settings", (LPCTSTR)"DISGARD NEW SETTINGS?", MB_ICONWARNING | MB_OKCANCEL | MB_DEFBUTTON2);
	if (UserResponse == IDOK) CDialog::OnCancel();
}


BOOL writeTABfile(){
const char *InputFilename = "Spectrum.txt";
const char *OutputFilename = "SpectrumTAB.txt";

	float var1, var2;
	char *ptrLine;

	std::string strLine;	
	std::ifstream inFile(InputFilename);
	std::ofstream outFile(OutputFilename, std::ofstream::out);

	while (std::getline(inFile, strLine)){
		ptrLine = strdup(strLine.c_str());
		sscanf(ptrLine, "%f %f", &var1, &var2);
		outFile << var1 << "\t" << var2 << "\n";
	}
	outFile.close();
	return TRUE;
}




void SpecDiagnostics::StopTimer() {
	if (m_specTimerHandle != NULL) {
		DeleteTimerQueueTimer(NULL, m_specTimerHandle, NULL);
		m_specTimerHandle = NULL;
	}
}

void CALLBACK SpecTimerProc(SpecDiagnostics *lpParametar, BOOLEAN TimerOrWaitFired)
{
	// This is used only to call QueueTimerHandler
	// Typically, this function is static member of CTimersDlg	
	SpecDiagnostics *ptrTimer;
	ptrTimer = lpParametar;
	ptrTimer->timerHandler();			
}

void SpecDiagnostics::StartTimer() {
	if (m_specTimerHandle == NULL) {
		DWORD elapsedTime = 1000;
		BOOL success = ::CreateTimerQueueTimer(  // Create new timer
			&m_specTimerHandle,
			NULL,
			(WAITORTIMERCALLBACK)SpecTimerProc,
			this,
			0,
			elapsedTime,
			WT_EXECUTEINTIMERTHREAD);
	}
}


BOOL TestApp::fetchSpectrometerData(float *ptrScopeData)
{		
	return TRUE;
}

BOOL SpecDiagnostics::startSpectrometerTest(BOOL filterState, int intPWM)
{	
	return TRUE;
}



BOOL SpecDiagnostics::OnInitDialog() 
{
	CDialog::OnInitDialog();
	pEdit = (CEdit*) GetDlgItem(IDC_EDIT_SPECTROMETER);
	m_EditBox_Spectrometer.SetLimitText(MAXLOG); 
	return TRUE;  // return TRUE unless you set the focus to a control
}	


void SpecDiagnostics::OnCheckedBoxSLC(){
	enableStrayLightCorrection = ((CButton*)GetDlgItem(IDC_CHECK_SLC))->GetCheck();
}

void SpecDiagnostics::OnBnClickedButtonRunTest()
{
}


void SpecDiagnostics::timerHandler() {
char strSeconds[32];

		if (MyTestApp.RunSpectrometerScan(intScanType) == NOT_DONE_YET)\
		{		
			sprintf_s (strSeconds, 32, "\r\nTime seconds: %d", intSeconds++);
			DisplayLog(strSeconds);
		}
		else 
		{
			StopTimer();
			DisplayLog("\r\nSCAN COMPLETE");
		}
	
}

// Start static Excel Automation demo
void OnRunStatic() 
{
}

void SpecDiagnostics::OnBnClickedButtonHaltTest()
{	
	StopTimer();
}


void SpecDiagnostics::OnBnClickedButtonLoadEeprom()
{
}



void SpecDiagnostics::ClearLog(){
	arrSpectrometerLog[0] = '\0';
	m_EditBox_Spectrometer.SetWindowText(arrSpectrometerLog);
	pEdit->LineScroll (pEdit->GetLineCount());
}

void SpecDiagnostics::DisplayLog(char *newString){		
	strcat_s(arrSpectrometerLog, MAXLOG, newString);
	m_EditBox_Spectrometer.SetWindowText(arrSpectrometerLog);
	pEdit->LineScroll (pEdit->GetLineCount());			
}


void SpecDiagnostics::OnBnClickedButtonActuator()
{
	if (intScanType == FILTER_OPEN){
		intScanType = FILTER_CLOSED;
		m_static_ButtonActuator.SetWindowText((LPCTSTR)"FILTER CLOSED");		
	}
	else if (intScanType == FILTER_CLOSED){
		intScanType = DARK_SCAN;
		m_static_ButtonActuator.SetWindowText((LPCTSTR)"DARK SCAN");
	}
	else {
		intScanType = FILTER_OPEN;
		m_static_ButtonActuator.SetWindowText((LPCTSTR)"FILTER OPEN");
	}	
}

void SpecDiagnostics::OnBnClickedButtonDarkScan()
{
	//ClearLog();	
	//intScanType = DARK_SCAN;
	//StartTimer();
}

void SpecDiagnostics::OnBnClickedButtonScan()
{
	ClearLog();	
	StartTimer();
}
