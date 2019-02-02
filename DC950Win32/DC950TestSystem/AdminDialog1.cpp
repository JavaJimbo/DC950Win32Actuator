// AdminDlg.cpp : implementation file for Admin1 Settings dialog box

#include "stdafx.h"
#include "Definitions.h"
#include <windows.h>
#include <iostream>
#include <sstream>
#include <string.h>
#include "AdminDialog1.h"
#include "TestApp.h"

extern TestApp MyTestApp;	
extern char *ptrExcelFilename;
extern char ExcelSecondaryFilename[];
extern char ExcelPrimaryFilename[];
char		EditExcelFilename[MAXFILENAME] = "";
BOOL		usingSecondaryExcelFile = FALSE;

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// Admin1 dialog

extern long portNumberInterfaceBoard, portNumberACpowerSupply, portNumberMultiMeter;
extern float AllowableLampVoltageError, AllowableVrefError,
MinClosedFilterPeakAmplitude, MaxClosedFilterPeakAmplitude,
MinClosedFilterWavelength, MaxClosedFilterWavelength,
MinClosedFilterIrradiance, MaxClosedFilterIrradiance,
MinClosedFilterFWHM, MaxClosedFilterFWHM,
SpareIniValue, MinOpenColorTemp,
MaxOpenColorTemp, MinOpenIrradiance,
MaxOpenIrradiance;



Admin1::Admin1(CWnd* pParent /*=NULL*/) : CDialog(Admin1::IDD, pParent)
{
	
}




void Admin1::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);	
	DDX_Control(pDX, IDC_SPIN_COM_INTERFACE, m_SpinCOMInterface);
	DDX_Control(pDX, IDC_EDIT_COM_INTERFACE, m_EditCOMInterface);	
	DDX_Control(pDX, IDC_SPIN_COM_MULTIMETER, m_SpinCOMMultiMeter);
	DDX_Control(pDX, IDC_EDIT_COM_MULTIMETER, m_EditCOMMultiMeter);	
	DDX_Control(pDX, IDC_SPIN_COM_POWER_SUPPLY, m_SpinCOMPowerSupply);
	DDX_Control(pDX, IDC_EDIT_COM_POWER_SUPPLY, m_EditCOMPowerSupply);	

	DDX_Control(pDX, IDC_SPIN_LAMP_VOLTAGE,	m_SpinAllowableLampVoltageError);
	DDX_Control(pDX, IDC_SPIN_REF_VOLTAGE, 	m_SpinAllowableVrefError);
	DDX_Control(pDX, IDC_SPIN_MIN_WAVELENGTH_FILTER, m_SpinMinClosedFilterWavelength);
	DDX_Control(pDX, IDC_SPIN_MAX_WAVELENGTH_FILTER, m_SpinMaxClosedFilterWavelength);
	DDX_Control(pDX, IDC_SPIN_MIN_IRRADIANCE_FILTER, m_SpinMinClosedFilterIrradiance);
	DDX_Control(pDX, IDC_SPIN_MAX_IRRADIANCE_FILTER, m_SpinMaxClosedFilterIrradiance);
	DDX_Control(pDX, IDC_SPIN_MIN_FWHM_FILTER, m_SpinMinClosedFilterFWHM);
	DDX_Control(pDX, IDC_SPIN_MAX_FWHM_FILTER, m_SpinMaxClosedFilterFWHM);

	DDX_Control(pDX, IDC_SPIN_MIN_COLOR_TEMP_OPEN, m_SpinMinOpenColorTemp);
	DDX_Control(pDX, IDC_SPIN_MAX_COLOR_TEMP_OPEN, m_SpinMaxOpenColorTemp);
	DDX_Control(pDX, IDC_SPIN_MIN_IRRADIANCE_OPEN, m_SpinMinOpenIrradiance);
	DDX_Control(pDX, IDC_SPIN_MAX_IRRADIANCE_OPEN, m_SpinMaxOpenIrradiance);

	DDX_Control(pDX, IDC_EDIT_MIN_COLOR_TEMP_OPEN, m_EditMinOpenColorTemp);
	DDX_Control(pDX, IDC_EDIT_MAX_COLOR_TEMP_OPEN, m_EditMaxOpenColorTemp);
	DDX_Control(pDX, IDC_EDIT_MIN_IRRADIANCE_OPEN, m_EditMinOpenIrradiance);
	DDX_Control(pDX, IDC_EDIT_MAX_IRRADIANCE_OPEN, m_EditMaxOpenIrradiance);


	DDX_Control(pDX, IDC_EDIT_LAMP_VOLTAGE,	m_EditAllowableLampVoltageError);
	DDX_Control(pDX, IDC_EDIT_REF_VOLTAGE, 	m_EditAllowableVrefError);
	DDX_Control(pDX, IDC_EDIT_MIN_WAVELENGTH_FILTER, m_EditMinClosedFilterWavelength);
	DDX_Control(pDX, IDC_EDIT_MAX_WAVELENGTH_FILTER, m_EditMaxClosedFilterWavelength);
	DDX_Control(pDX, IDC_EDIT_MIN_IRRADIANCE_FILTER, m_EditMinClosedFilterIrradiance);
	DDX_Control(pDX, IDC_EDIT_MAX_IRRADIANCE_FILTER, m_EditMaxClosedFilterIrradiance);
	DDX_Control(pDX, IDC_EDIT_MIN_FWHM_FILTER, m_EditMinClosedFilterFWHM);
	DDX_Control(pDX, IDC_EDIT_MAX_FWHM_FILTER, m_EditMaxClosedFilterFWHM);

	DDX_Control(pDX, IDC_SPIN_MIN_CLOSED_FILTER_PEAK_AMPLITUDE, m_SpinMinClosedFilterPeakAmplitude);
	DDX_Control(pDX, IDC_SPIN_MAX_CLOSED_FILTER_PEAK_AMPLITUDE, m_SpinMaxClosedFilterPeakAmplitude);
	DDX_Control(pDX, IDC_EDIT_MIN_CLOSED_FILTER_PEAK_AMPLITUDE, m_EditMinClosedFilterPeakAmplitude);
	DDX_Control(pDX, IDC_EDIT_MAX_CLOSED_FILTER_PEAK_AMPLITUDE, m_EditMaxClosedFilterPeakAmplitude);

	DDX_Control(pDX, IDC_EDIT_OUTFILE, m_EditBox_OutputFileName);
	DDX_Control(pDX, IDC_RADIO_PRIMARY, m_static_ButtonPrimary);
	DDX_Control(pDX, IDC_RADIO_SECONDARY, m_static_ButtonSecondary);
	DDX_Control(pDX, IDC_STATIC_SPREADSELECT, m_Group_SpreadsheetSelect);
}
	

BEGIN_MESSAGE_MAP(Admin1, CDialog)
	ON_WM_CLOSE()		
	ON_BN_CLICKED(IDOK, &Admin1::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &Admin1::OnBnClickedCancel)
	ON_BN_CLICKED(IDC_RADIO_PRIMARY, &Admin1::OnBnClickedPrimary)
	ON_BN_CLICKED(IDC_RADIO_SECONDARY, &Admin1::OnBnClickedSecondary)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// Admin1 message handlers


void Admin1::OnClose() 
{
	// TODO: Add your message handler code here and/or call default
	CDialog::OnClose();
}



BOOL Admin1::OnInitDialog() 
{
	CDialog::OnInitDialog();	
	strcpy_s(EditExcelFilename, MAXFILENAME, ptrExcelFilename);
	m_EditBox_OutputFileName.SetWindowText ((LPCTSTR)EditExcelFilename);

	if (ptrExcelFilename == ExcelSecondaryFilename)
	{
		usingSecondaryExcelFile = TRUE;
		m_static_ButtonSecondary.EnableWindow(TRUE);
		m_static_ButtonSecondary.SetCheck(TRUE);
		m_static_ButtonPrimary.EnableWindow(TRUE);
		m_Group_SpreadsheetSelect.EnableWindow(TRUE);
	}
	else
	{
		usingSecondaryExcelFile = FALSE;
		m_static_ButtonSecondary.EnableWindow(FALSE);
		m_static_ButtonPrimary.SetCheck(TRUE);
		m_static_ButtonPrimary.EnableWindow(FALSE);
		m_Group_SpreadsheetSelect.EnableWindow(FALSE);
	}

	// Initilize Numerical Spin Controls	
	m_SpinCOMInterface.SetBuddy(&m_EditCOMInterface);
	m_SpinCOMInterface.SetRange(1, MAX_PORTNUMBER);
	m_SpinCOMInterface.SetPos((float) portNumberInterfaceBoard);
	m_SpinCOMInterface.SetDelta(1);	
	
	m_SpinCOMMultiMeter.SetBuddy(&m_EditCOMMultiMeter);
	m_SpinCOMMultiMeter.SetRange(1, MAX_PORTNUMBER);
	m_SpinCOMMultiMeter.SetPos((float) portNumberMultiMeter);
	m_SpinCOMMultiMeter.SetDelta(1);		

	m_SpinCOMPowerSupply.SetBuddy(&m_EditCOMPowerSupply);
	m_SpinCOMPowerSupply.SetRange(1, MAX_PORTNUMBER);
	m_SpinCOMPowerSupply.SetPos((float) portNumberACpowerSupply);
	m_SpinCOMPowerSupply.SetDelta(1);		

	m_SpinAllowableLampVoltageError.SetBuddy(&m_EditAllowableLampVoltageError);
	m_SpinAllowableLampVoltageError.SetRange((float) 0, (float) 10.000);
	m_SpinAllowableLampVoltageError.SetPos(AllowableLampVoltageError);
	m_SpinAllowableLampVoltageError.SetDelta((float) 0.01);

	m_SpinAllowableVrefError.SetBuddy(&m_EditAllowableVrefError);
	m_SpinAllowableVrefError.SetRange((float) 0, (float) 1.000);
	m_SpinAllowableVrefError.SetPos(AllowableVrefError);
	m_SpinAllowableVrefError.SetDelta((float) 0.01);

	m_SpinMinClosedFilterPeakAmplitude.SetBuddy(&m_EditMinClosedFilterPeakAmplitude);
	m_SpinMinClosedFilterPeakAmplitude.SetRange((float)0, 10000);
	m_SpinMinClosedFilterPeakAmplitude.SetPos(MinClosedFilterPeakAmplitude);
	m_SpinMinClosedFilterPeakAmplitude.SetDelta(1);

	m_SpinMaxClosedFilterPeakAmplitude.SetBuddy(&m_EditMaxClosedFilterPeakAmplitude);
	m_SpinMaxClosedFilterPeakAmplitude.SetRange((float)0, 10000);
	m_SpinMaxClosedFilterPeakAmplitude.SetPos(MaxClosedFilterPeakAmplitude);
	m_SpinMaxClosedFilterPeakAmplitude.SetDelta(1);

	m_SpinMinClosedFilterWavelength.SetBuddy(&m_EditMinClosedFilterWavelength);
	m_SpinMinClosedFilterWavelength.SetRange((float)0, 10000);
	m_SpinMinClosedFilterWavelength.SetPos(MinClosedFilterWavelength); 
	m_SpinMinClosedFilterWavelength.SetDelta(1);

	m_SpinMaxClosedFilterWavelength.SetBuddy(&m_EditMaxClosedFilterWavelength);
	m_SpinMaxClosedFilterWavelength.SetRange((float)0, 10000);
	m_SpinMaxClosedFilterWavelength.SetPos(MaxClosedFilterWavelength); 
	m_SpinMaxClosedFilterWavelength.SetDelta(1);

	m_SpinMinClosedFilterIrradiance.SetBuddy(&m_EditMinClosedFilterIrradiance);
	m_SpinMinClosedFilterIrradiance.SetRange((float)0, 10000);
	m_SpinMinClosedFilterIrradiance.SetPos(MinClosedFilterIrradiance); 
	m_SpinMinClosedFilterIrradiance.SetDelta(1);

	m_SpinMaxClosedFilterIrradiance.SetBuddy(&m_EditMaxClosedFilterIrradiance);
	m_SpinMaxClosedFilterIrradiance.SetRange((float)0, 10000);
	m_SpinMaxClosedFilterIrradiance.SetPos(MaxClosedFilterIrradiance); 
	m_SpinMaxClosedFilterIrradiance.SetDelta(1);

	m_SpinMinClosedFilterFWHM.SetBuddy(&m_EditMinClosedFilterFWHM);
	m_SpinMinClosedFilterFWHM.SetRange((float)0, 10000);
	m_SpinMinClosedFilterFWHM.SetPos(MinClosedFilterFWHM); 
	m_SpinMinClosedFilterFWHM.SetDelta(1);

	m_SpinMaxClosedFilterFWHM.SetBuddy(&m_EditMaxClosedFilterFWHM);
	m_SpinMaxClosedFilterFWHM.SetRange((float)0, 10000);
	m_SpinMaxClosedFilterFWHM.SetPos(MaxClosedFilterFWHM); 
	m_SpinMaxClosedFilterFWHM.SetDelta(1);

	m_SpinMinOpenIrradiance.SetBuddy(&m_EditMinOpenIrradiance);
	m_SpinMinOpenIrradiance.SetRange((float)0, 10000);
	m_SpinMinOpenIrradiance.SetPos(MinOpenIrradiance); 
	m_SpinMinOpenIrradiance.SetDelta(1);

	m_SpinMaxOpenIrradiance.SetBuddy(&m_EditMaxOpenIrradiance);
	m_SpinMaxOpenIrradiance.SetRange((float)0, 10000);
	m_SpinMaxOpenIrradiance.SetPos(MaxOpenIrradiance); 
	m_SpinMaxOpenIrradiance.SetDelta(1);

	m_SpinMinOpenColorTemp.SetBuddy(&m_EditMinOpenColorTemp);
	m_SpinMinOpenColorTemp.SetRange((float)0, 10000);
	m_SpinMinOpenColorTemp.SetPos(MinOpenColorTemp); 
	m_SpinMinOpenColorTemp.SetDelta(1);

	m_SpinMaxOpenColorTemp.SetBuddy(&m_EditMaxOpenColorTemp);
	m_SpinMaxOpenColorTemp.SetRange((float)0, 10000);
	m_SpinMaxOpenColorTemp.SetPos(MaxOpenColorTemp); 
	m_SpinMaxOpenColorTemp.SetDelta(1);
	

	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}


void Admin1::OnNMThemeChangedScrollComInterface(NMHDR *pNMHDR, LRESULT *pResult)
{
	// This feature requires Windows XP or greater.
	// The symbol _WIN32_WINNT must be >= 0x0501.
	// TODO: Add your control notification handler code here
	*pResult = 0;
}


void Admin1::OnDeltaposSpin1(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMUPDOWN pNMUpDown = reinterpret_cast<LPNMUPDOWN>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0;
}

void Admin1::OnPaint() 
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

void Admin1::OnBnClickedOk()
{
	int UserResponse = 0;

	UserResponse = MessageBox((LPCTSTR)"Click OK to save new settings", (LPCTSTR)"SAVE NEW SETTINGS?", MB_ICONWARNING | MB_OKCANCEL | MB_DEFBUTTON2);
	if (UserResponse == IDOK) {	
		portNumberInterfaceBoard = (long) m_SpinCOMInterface.GetPos();
		portNumberMultiMeter = (long) m_SpinCOMMultiMeter.GetPos();
		portNumberACpowerSupply = (long) m_SpinCOMPowerSupply.GetPos();
	
		AllowableLampVoltageError = m_SpinAllowableLampVoltageError.GetPos();
		AllowableVrefError = m_SpinAllowableVrefError.GetPos();
		MinClosedFilterPeakAmplitude = m_SpinMinClosedFilterPeakAmplitude.GetPos();
		MaxClosedFilterPeakAmplitude = m_SpinMaxClosedFilterPeakAmplitude.GetPos();
		MinClosedFilterWavelength = m_SpinMinClosedFilterWavelength.GetPos();
		MaxClosedFilterWavelength = m_SpinMaxClosedFilterWavelength.GetPos();
		MinClosedFilterIrradiance = m_SpinMinClosedFilterIrradiance.GetPos();
		MaxClosedFilterIrradiance = m_SpinMaxClosedFilterIrradiance.GetPos();
		MinClosedFilterFWHM = m_SpinMinClosedFilterFWHM.GetPos();
		MaxClosedFilterFWHM = m_SpinMaxClosedFilterFWHM.GetPos(); 
		MinOpenColorTemp = m_SpinMinOpenColorTemp.GetPos(); 
		MaxOpenColorTemp = m_SpinMaxOpenColorTemp.GetPos(); 
		MinOpenIrradiance = m_SpinMinOpenIrradiance.GetPos();
		MaxOpenIrradiance = m_SpinMaxOpenIrradiance.GetPos();

		m_EditBox_OutputFileName.GetWindowText((LPTSTR)EditExcelFilename, MAXSTRING);		
		
		// IF USER HAS SELECTED PRIMARY SPREADSHEET:
		if (m_static_ButtonPrimary.GetCheck() == TRUE)
		{			
			// Make sure filename is valid:
			if (!MyTestApp.validateSpreedsheetFilename(EditExcelFilename))
				MessageBox((LPCTSTR)"Spreadsheet name must end with 'XLS'\r\n and may not contain the characters [ ] < > ? / \\ * \" : |", (LPCTSTR)"SPREADSHEET FILENAME ERROR", MB_ICONWARNING | MB_OK | MB_DEFBUTTON2);
			// and make sure path for filename exists:
			else if (!checkOutputDirectory(EditExcelFilename))					
				MessageBox((LPCTSTR) "Be sure path for spreadsheet is correctly entered", (LPCTSTR)"ERROR: CANNOT LOCATE SPREADSHEET DIRECTORY", MB_ICONWARNING | MB_OK | MB_DEFBUTTON2);		
			else
			{
				// Store any settings that may have been changed (other than the spreadsheet)
				MyTestApp.storeINIfileBinary();

				int userEditedPrimaryName = strcmp(ExcelPrimaryFilename, EditExcelFilename);

				// Now deal with the spreadsheet: If Secondary was being used 
				// or Primary filename is being changed:
				if (usingSecondaryExcelFile ||  userEditedPrimaryName)
				{						
					// Check whether a file with new name and path already exists:
					if (GetFileAttributes(EditExcelFilename) != INVALID_FILE_ATTRIBUTES)									
						// If so, then prompt the operator to decide whether to overwrite it
						UserResponse = MessageBox((LPCTSTR)"A file with this name already exists\r\nDo you wish to overwrite it?", (LPCTSTR)"OVERWRITE FILE?", MB_ICONWARNING | MB_OKCANCEL | MB_DEFBUTTON2);
				}

				// If user clicks OK to use PrimaryFile / or no errors exist, save and exit:						
				if (UserResponse == IDOK) 
				{
					if (usingSecondaryExcelFile)									// If Secondary file was being used up until now,
						CopyFile(ExcelSecondaryFilename, EditExcelFilename, FALSE);	// then copy data over from Secondary to New file
					else if (userEditedPrimaryName)										// Otherwise if user changed to new Primary file,
						CopyFile(ExcelPrimaryFilename, EditExcelFilename, FALSE);	// then copy data from olde to new Primary file

					if (userEditedPrimaryName) strcpy_s(ExcelPrimaryFilename, MAXFILENAME, EditExcelFilename);  // If name is new, copy to Primary.
					strcpy_s(ExcelSecondaryFilename, MAXFILENAME, "");				// Delete Secondary filename since it is no longer needed.
					MyTestApp.storeExcelFilenames(ExcelPrimaryFilename, ExcelSecondaryFilename);		// Store filenames.
					ptrExcelFilename = ExcelPrimaryFilename;						// Filename pointer gets set here so data will now be stored in Primary file.

					 
					CDialog::OnOK();										// Close Admin dialog box.
				}				
			}
		}
		// Otherwise, if user has selected secondary and attempts to rename it, display error message:
		else if (strcmp(ExcelSecondaryFilename, EditExcelFilename) != 0)		
			UserResponse = MessageBox((LPCTSTR)"If you wish to rename spreadsheet,\r\nplease select Using Primary", (LPCTSTR)"SECONDARY SPREADSHEET CANNOT BE RENAMED", MB_ICONWARNING | MB_OK | MB_DEFBUTTON2);
		// Otherwise if user wishes to continue using Secondary with existing name,
		// save non-spreadsheet settings and close Admin box:
		else 
		{
			MyTestApp.storeINIfileBinary();
			CDialog::OnOK();
		}
	}
}

void Admin1::OnBnClickedCancel()
{
	int UserResponse = MessageBox((LPCTSTR)"Click OK to quit without saving new settings", (LPCTSTR)"DISGARD NEW SETTINGS?", MB_ICONWARNING | MB_OKCANCEL | MB_DEFBUTTON2);
	if (UserResponse == IDOK) CDialog::OnCancel();
}


BOOL Admin1::checkOutputDirectory(char *ptrFilename)
{
	int i, j, length;
	char ptrPathName[MAXFILENAME];
	
	length = strlen (ptrFilename);
	if (length >= MAXFILENAME) return FALSE;
	for (i = length-1; i > 0; i--)		
		if (ptrFilename[i] == '\\') break;
	for (j = 0; j < i; j++)
		ptrPathName[j] = ptrFilename[j];
	ptrPathName[j] = '\0';	
	
	if (GetFileAttributes(ptrPathName) == INVALID_FILE_ATTRIBUTES) 				
		return FALSE;	
	else return TRUE;
}


void Admin1::OnBnClickedPrimary()
{
	m_EditBox_OutputFileName.SetWindowText ((LPCTSTR)ExcelPrimaryFilename);	
}

void Admin1::OnBnClickedSecondary()
{
	m_EditBox_OutputFileName.SetWindowText ((LPCTSTR)ExcelSecondaryFilename);	
}
