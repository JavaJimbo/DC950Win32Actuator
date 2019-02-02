// SpecDiagnosticsDlg.cpp : implementation file for Admin2 Settings dialog box

#include "stdafx.h"
#include "Definitions.h"
#include <windows.h>
#include <fstream>
#include <iostream>
#include <sstream>
#include <string.h>
#include "AdminDialog2.h"
#include "TestApp.h"
#include "avaspec.h"
#include "ReadOnlyEdit.h"
#include "TestApp.h"
#include "irrad.h"

extern TestApp MyTestApp;
extern float MinRemoteVoltLimit[];
extern float MaxRemoteVoltLimit[];

extern float AllowableLampVoltageError;
extern float AllowableVrefError;
extern float MinClosedFilterPeakAmplitude;
extern float MaxClosedFilterPeakAmplitude;
extern float MinClosedFilterWavelength;
extern float MaxClosedFilterWavelength;
extern float MinClosedFilterIrradiance;
extern float MaxClosedFilterIrradiancep;
extern float MinClosedFilterFWHM;
extern float MaxClosedFilterFWHM;
extern float SpareIniValue;

extern float MinOpenColorTemp;
extern float MaxOpenColorTemp;
extern float MinOpenIrradiance;
extern float MaxOpenIrradiance;

// SPECTROMETER GLOBAL VARIABLES
extern DeviceConfigType* l_pDeviceData;
extern int m_NrDevices;
extern long handleSpectrometer;
extern unsigned int numberOfPixels;

extern double Lambda[];
extern float openFilterData[];
extern float darkData[];
extern float closedFilterData[];
extern float irradianceIntensity[];
extern float m_currentIntegrationTime;

extern uint16	m_StartPixel;
extern uint16	m_StopPixel;
extern uint16	m_NumberOfScans;
extern BOOL m_enableLinearCorrection;
extern BOOL m_enableStrayLightCorrection;

extern CTestDialog *ptrDialog;
extern DATAstring arrPartNumbers[];

extern int totalPartNumbers;


// extern float arrINIconfigValues[MAXINIVALUES];

extern float openColorTemp, openIrradianceIntegral, closedIrradianceIntegral;	
extern float centerWavelength, wavelengthAmplitude, FWHM;

// HANDLE handleInterfaceBoard = NULL, handleHPmultiMeter = NULL, handleACpowerSupply = NULL;
// CFont BigFont, SmallFont, MidFont;
int testSubStepNumber = 0;


// SPECTROMETER GLOBAL VARIABLES
extern DeviceConfigType* l_pDeviceData;
extern int m_NrDevices;
extern long handleSpectrometer;
extern unsigned int numberOfPixels;
extern double Lambda[];

extern uint16	m_StartPixel;
extern uint16	m_StopPixel;

extern float	m_currentIntegrationTime;
extern unsigned long m_numberOfAverages;
extern uint16	m_NumberOfScans;

extern BOOL		m_enableLinearCorrection;
extern BOOL		m_enableStrayLightCorrection;

char strPartNumberList[MAXLOG] = {'\0'};

extern TestApp MyTestApp;
extern DATAstring arrPartNumbers[];
extern int totalPartNumbers;
HANDLE m_specTimerHandle = NULL;
int intSeconds = 0;

extern char AdminPassword[];

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// Admin2 dialog

Admin2::Admin2(CWnd* pParent /*=NULL*/) : CDialog(Admin2::IDD, pParent)
{
	
}

Admin2::~Admin2()
{

}

void Admin2::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);	
	 	
	DDX_Control(pDX, IDC_CHECK_SLC, m_checkEnableStrayLightCorrection);
	DDX_Control(pDX, IDC_CHECK_LINEAR, m_checkEnableLinearCorrection);

	DDX_Control(pDX, IDC_SPIN_INT_TIME, m_SpinIntegrationTime);
	DDX_Control(pDX, IDC_SPIN_NUM_AVERAGES, m_SpinNumAverages);
	DDX_Control(pDX, IDC_SPIN_NUM_SCANS, m_SpinNumScans);

	DDX_Control(pDX, IDC_EDIT_INT_TIME, m_EditIntegrationTime);
	DDX_Control(pDX, IDC_EDIT_NUM_AVERAGES, m_EditNumAverages);
	DDX_Control(pDX, IDC_EDIT_NUM_SCANS, m_EditNumScans);

	DDX_Control(pDX, IDC_EDIT_PASSWORD, m_EditBox_Password);
	DDX_Control(pDX, IDC_EDIT_PARTNUMBERS, m_EditBox_PartNumber);
	
	DDX_Control(pDX, IDC_EDIT_MIN1, m_EditBox_Min1);
	DDX_Control(pDX, IDC_EDIT_MIN2, m_EditBox_Min2);
	DDX_Control(pDX, IDC_EDIT_MIN3, m_EditBox_Min3);
	DDX_Control(pDX, IDC_EDIT_MIN4, m_EditBox_Min4);
	DDX_Control(pDX, IDC_EDIT_MIN5, m_EditBox_Min5);
	DDX_Control(pDX, IDC_EDIT_MIN6, m_EditBox_Min6);

	
	DDX_Control(pDX, IDC_EDIT_MAX1, m_EditBox_Max1);
	DDX_Control(pDX, IDC_EDIT_MAX2, m_EditBox_Max2);
	DDX_Control(pDX, IDC_EDIT_MAX3, m_EditBox_Max3);
	DDX_Control(pDX, IDC_EDIT_MAX4, m_EditBox_Max4);
	DDX_Control(pDX, IDC_EDIT_MAX5, m_EditBox_Max5);
	DDX_Control(pDX, IDC_EDIT_MAX6, m_EditBox_Max6);
}

BEGIN_MESSAGE_MAP(Admin2, CDialog)
	ON_WM_CLOSE()		
	ON_BN_CLICKED(IDOK, &Admin2::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &Admin2::OnBnClickedCancel)	
	ON_BN_CLICKED(IDC_CHECK_SLC, &Admin2::OnCheckedBoxSLC)	
	ON_BN_CLICKED(IDC_CHECK_LINEAR, &Admin2::OnBnClickedCheckLinear)
	ON_BN_CLICKED(IDC_BUTTON_PASSWORD, &Admin2::OnBnClickedButtonPassword)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// Admin2 message handlers

//int Admin2::DoModal() 
//{
	// TODO: Add your specialized code here and/or call the base class
//	return (int) CDialog::DoModal();
//}

void Admin2::OnClose() 
{
	// TODO: Add your message handler code here and/or call default
	CDialog::OnClose();
}






BOOL Admin2::CopyTextFromPartNumberEditBox()
{
	int i = 0;
	int length = 0;
	char *ptrPartNumber = NULL;
	char delimiters[] = " \r\n";
	int UserResponse = 0;

	m_EditBox_PartNumber.GetWindowText((LPTSTR)strPartNumberList, MAXLOG);
	
	ptrPartNumber = strtok(strPartNumberList, delimiters);
	i = 0;
	while (ptrPartNumber != NULL && i < MAX_PART_NUMBERS){
		length = strlen(ptrPartNumber);		
		if (length == VALID_PART_NUMBER_LENGTH)
		{
			strcpy_s(arrPartNumbers[i].String, MAXSTRING, ptrPartNumber);
			i++;
			if (i >= MAX_PART_NUMBERS) break;
		}
		else return FALSE;
		//{
		//	UserResponse = MessageBox((LPCTSTR)"Invalid part number length./r/nSave valid part numbers?", (LPCTSTR)"INVALID PART NUMBER", MB_ICONWARNING | MB_OKCANCEL | MB_DEFBUTTON2);
		//	if (UserResponse != IDOK) return FALSE;
		//}
		ptrPartNumber = strtok(NULL, delimiters);
	}
	totalPartNumbers = i;
	return TRUE;
}

void Admin2::OnBnClickedCancel()
{
	int UserResponse = MessageBox((LPCTSTR)"Click OK to quit without saving new settings", (LPCTSTR)"DISGARD NEW SETTINGS?", MB_ICONWARNING | MB_OKCANCEL | MB_DEFBUTTON2);
	if (UserResponse == IDOK) 
	{	
		CDialog::OnCancel();
	}
}

// This routine prevents the ENTER or ESCAPE keys from closing the main dialog box
// and thereby shutting down the program. It is necessary because 
// the barcode scanner otherwise closes the dialog box 
// when it sends a carriage return after the serial number.
BOOL Admin2::PreTranslateMessage(MSG* pMsg)
{
	if (pMsg->message == WM_KEYDOWN)
	{
		if (pMsg->wParam == VK_ESCAPE)
		{
			return TRUE;                // Do not process further
		}
		else if (pMsg->wParam == VK_RETURN)
		{
			CEdit* ptrEdit = (CEdit*)GetFocus();
			if (ptrEdit == (CEdit*)GetDlgItem(IDC_EDIT_PARTNUMBERS)){
				m_EditBox_PartNumber.GetWindowText((LPTSTR)strPartNumberList, MAXLOG);
				strcat_s(strPartNumberList, MAXLOG, "\r\n");	
				m_EditBox_PartNumber.SetWindowText(strPartNumberList);
				ptrEdit->GetDlgItem(IDC_EDIT_PARTNUMBERS);
				ptrEdit->SetSel(-1);
				ptrEdit->SetFocus();				
				ptrEdit->ShowCaret();
				ptrEdit->SendMessage(EM_SETREADONLY, 0, 0);
			}
			return TRUE;                
		}
	}

	return CWnd::PreTranslateMessage(pMsg);
}



void Admin2::OnCheckedBoxSLC(){
	m_enableStrayLightCorrection = ((CButton*)GetDlgItem(IDC_CHECK_SLC))->GetCheck();
}

void Admin2::ClearLog(){
	strPartNumberList[0] = '\0';
	m_EditBox_PartNumber.SetWindowText(strPartNumberList);
	pEdit->LineScroll (pEdit->GetLineCount());
}

void Admin2::DisplayLog(char *newString){		
	strcat_s(strPartNumberList, MAXLOG, newString);
	m_EditBox_PartNumber.SetWindowText(strPartNumberList);
	pEdit->LineScroll (pEdit->GetLineCount());			
}


void Admin2::OnBnClickedCheckLinear()
{
	// TODO: Add your control notification handler code here
}


// m_SpinNumScans;
BOOL Admin2::OnInitDialog() 
{
	CDialog::OnInitDialog();
	int i;
	char strValue[MAXSTRING];

	// Min and Max Remote voltage test limits	
	sprintf_s(strValue, MAXSTRING, "%.2f", MinRemoteVoltLimit[1]);
	m_EditBox_Min1.SetWindowText((LPTSTR)strValue);
	sprintf_s(strValue, MAXSTRING, "%.2f", MinRemoteVoltLimit[2]);
	m_EditBox_Min2.SetWindowText((LPTSTR)strValue);
	sprintf_s(strValue, MAXSTRING, "%.2f", MinRemoteVoltLimit[3]);
	m_EditBox_Min3.SetWindowText((LPTSTR)strValue);
	sprintf_s(strValue, MAXSTRING, "%.2f", MinRemoteVoltLimit[4]);
	m_EditBox_Min4.SetWindowText((LPTSTR)strValue);
	sprintf_s(strValue, MAXSTRING, "%.2f", MinRemoteVoltLimit[5]);
	m_EditBox_Min5.SetWindowText((LPTSTR)strValue);
	sprintf_s(strValue, MAXSTRING, "%.2f", MinRemoteVoltLimit[6]);
	m_EditBox_Min6.SetWindowText((LPTSTR)strValue);

	sprintf_s(strValue, MAXSTRING, "%.2f", MaxRemoteVoltLimit[1]);
	m_EditBox_Max1.SetWindowText((LPTSTR)strValue);
	sprintf_s(strValue, MAXSTRING, "%.2f", MaxRemoteVoltLimit[2]);
	m_EditBox_Max2.SetWindowText((LPTSTR)strValue);
	sprintf_s(strValue, MAXSTRING, "%.2f", MaxRemoteVoltLimit[3]);
	m_EditBox_Max3.SetWindowText((LPTSTR)strValue);
	sprintf_s(strValue, MAXSTRING, "%.2f", MaxRemoteVoltLimit[4]);
	m_EditBox_Max4.SetWindowText((LPTSTR)strValue);
	sprintf_s(strValue, MAXSTRING, "%.2f", MaxRemoteVoltLimit[5]);
	m_EditBox_Max5.SetWindowText((LPTSTR)strValue);
	sprintf_s(strValue, MAXSTRING, "%.2f", MaxRemoteVoltLimit[6]);
	m_EditBox_Max6.SetWindowText((LPTSTR)strValue);


	// Initilize Numerical Spin Controls
	m_SpinIntegrationTime.SetBuddy(&m_EditIntegrationTime);
	m_SpinIntegrationTime.SetRange(0, 1000);
	m_SpinIntegrationTime.SetPos((float) m_currentIntegrationTime);
	m_SpinIntegrationTime.SetDelta(1);		

	m_SpinNumAverages.SetBuddy(&m_EditNumAverages);
	m_SpinNumAverages.SetRange(0, 1000);
	m_SpinNumAverages.SetPos((float) m_numberOfAverages);
	m_SpinNumAverages.SetDelta(1);		

	m_SpinNumScans.SetBuddy(&m_EditNumScans);
	m_SpinNumScans.SetRange(0, 1000);
	m_SpinNumScans.SetPos((float) m_NumberOfScans);
	m_SpinNumScans.SetDelta(1);		

	if (m_enableLinearCorrection) m_checkEnableLinearCorrection.SetCheck(TRUE);
	else m_checkEnableLinearCorrection.SetCheck(FALSE);
		
	if (m_enableStrayLightCorrection) m_checkEnableStrayLightCorrection.SetCheck(TRUE);
	else m_checkEnableStrayLightCorrection.SetCheck(FALSE);

	m_EditBox_PartNumber.SetLimitText(MAXLOG); 
	if (totalPartNumbers > MAX_PART_NUMBERS)
		totalPartNumbers = MAX_PART_NUMBERS;
	strPartNumberList[0] = '\0';
	for (i = 0; i < totalPartNumbers; i++)
	{
		strcat_s(strPartNumberList, MAXLOG, arrPartNumbers[i].String);
		if (i < totalPartNumbers-1) strcat_s(strPartNumberList, MAXLOG, "\r\n");
	}
	m_EditBox_PartNumber.SetWindowText(strPartNumberList);
	m_EditBox_Password.SetWindowText(AdminPassword);

	return TRUE;  // return TRUE unless you set the focus to a control
}	


void Admin2::OnBnClickedOk()
{
	char strValue[MAXSTRING];
	int UserResponse = MessageBox((LPCTSTR)"Click OK to save new settings", (LPCTSTR)"SAVE NEW SETTINGS?", MB_ICONWARNING | MB_OKCANCEL | MB_DEFBUTTON2);
	
	if (UserResponse == IDOK) {	
		if (m_SpinIntegrationTime.GetPos() == (float) 0 || m_SpinNumAverages.GetPos() == (float) 0 || m_SpinNumScans.GetPos() == (float) 0)
			MessageBox((LPCTSTR)"Integration time, number of averages, and number of scans must be non-zero", (LPCTSTR)"INVALID VALUE", MB_ICONWARNING | MB_OK | MB_DEFBUTTON2);		
		else if (CopyTextFromPartNumberEditBox() == FALSE) 
			MessageBox((LPCTSTR)"Part numbers must all be 12 digits long", (LPCTSTR)"INVALID PART NUMBER", MB_ICONWARNING | MB_OK | MB_DEFBUTTON2);			
		else {
			MyTestApp.storeINIpartNumbers();	

			m_currentIntegrationTime = (float) m_SpinIntegrationTime.GetPos();	
			m_numberOfAverages = (unsigned long) m_SpinNumAverages.GetPos();
			m_NumberOfScans = (uint16) m_SpinNumScans.GetPos();
			
			m_EditBox_Min1.GetWindowText((LPTSTR)strValue, MAXSTRING);
			MinRemoteVoltLimit[1] = (float) atof(strValue);
			m_EditBox_Min2.GetWindowText((LPTSTR)strValue, MAXSTRING);
			MinRemoteVoltLimit[2] = (float) atof(strValue);
			m_EditBox_Min3.GetWindowText((LPTSTR)strValue, MAXSTRING);
			MinRemoteVoltLimit[3] = (float) atof(strValue);
			m_EditBox_Min4.GetWindowText((LPTSTR)strValue, MAXSTRING);
			MinRemoteVoltLimit[4] = (float) atof(strValue);
			m_EditBox_Min5.GetWindowText((LPTSTR)strValue, MAXSTRING);
			MinRemoteVoltLimit[5] = (float) atof(strValue);
			m_EditBox_Min6.GetWindowText((LPTSTR)strValue, MAXSTRING);
			MinRemoteVoltLimit[6] = (float) atof(strValue);

			m_EditBox_Max1.GetWindowText((LPTSTR)strValue, MAXSTRING);
			MaxRemoteVoltLimit[1] = (float) atof(strValue);
			m_EditBox_Max2.GetWindowText((LPTSTR)strValue, MAXSTRING);
			MaxRemoteVoltLimit[2] = (float) atof(strValue);
			m_EditBox_Max3.GetWindowText((LPTSTR)strValue, MAXSTRING);
			MaxRemoteVoltLimit[3] = (float) atof(strValue);
			m_EditBox_Max4.GetWindowText((LPTSTR)strValue, MAXSTRING);
			MaxRemoteVoltLimit[4] = (float) atof(strValue);
			m_EditBox_Max5.GetWindowText((LPTSTR)strValue, MAXSTRING);
			MaxRemoteVoltLimit[5] = (float) atof(strValue);
			m_EditBox_Max6.GetWindowText((LPTSTR)strValue, MAXSTRING);
			MaxRemoteVoltLimit[6] = (float) atof(strValue);

			if (m_checkEnableLinearCorrection.GetCheck() == 0) m_enableLinearCorrection = FALSE;
			else m_enableLinearCorrection = TRUE;

			if (m_checkEnableStrayLightCorrection.GetCheck() == 0) m_enableStrayLightCorrection = FALSE;
			else m_enableStrayLightCorrection = TRUE;
			MyTestApp.storeINIfileBinary();

			CDialog::OnOK();	
		}
	}
}

void Admin2::OnPaint() 
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

void Admin2::OnNMThemeChangedScrollComInterface(NMHDR *pNMHDR, LRESULT *pResult)
{
	// This feature requires Windows XP or greater.
	// The symbol _WIN32_WINNT must be >= 0x0501.
	// TODO: Add your control notification handler code here
	*pResult = 0;
}


void Admin2::OnDeltaposSpin1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMUPDOWN pNMUpDown = reinterpret_cast<LPNMUPDOWN>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0;
}



void Admin2::StopTimer() {
	if (m_specTimerHandle != NULL) {
		DeleteTimerQueueTimer(NULL, m_specTimerHandle, NULL);
		m_specTimerHandle = NULL;
	}
}

void CALLBACK SpecTimerProc(Admin2 *lpParametar, BOOLEAN TimerOrWaitFired)
{
	// This is used only to call QueueTimerHandler
	// Typically, this function is static member of CTimersDlg	
	Admin2 *ptrTimer;
	ptrTimer = lpParametar;
	ptrTimer->timerHandler();			
}

void Admin2::StartTimer() {
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

void Admin2::timerHandler() {
char strSeconds[32];

	sprintf_s (strSeconds, 32, "\r\nTime seconds: %d", intSeconds++);
	DisplayLog(strSeconds);
}

void Admin2::OnBnClickedButtonPassword()
{
	char strTemp[MAXPASSWORD];

	m_EditBox_Password.GetWindowText((LPTSTR)strTemp, MAXPASSWORD);
	if (MyTestApp.isValidPassword(strTemp))
	{
		strcpy_s(AdminPassword, MAXPASSWORD, strTemp);
		MyTestApp.storeAdminPassword();
		MessageBox((LPCTSTR)"New password saved to PasswordFile.txt", (LPCTSTR)"PASSWORD OK", MB_ICONWARNING | MB_OK | MB_DEFBUTTON2);
	}
	else MessageBox((LPCTSTR)"Password must have at least 5 characters\r\nand may use any combination of\r\nletters, numbers, and/or punctuation", (LPCTSTR)"INVALID PASSWORD", MB_ICONWARNING | MB_OK | MB_DEFBUTTON2);
}



