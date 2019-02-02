/* Initialize.cpp - routines for initializing hardware at startup and also reading/writing INI file.
 * Includes functions for initializing interface box, multimeter, and power supply.
 * Written in Visual C++ for DC950 test system
 *
 * 
 */

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
// #include "BasicExcel.hpp"
// #include "ExcelFormat.h"
#include <time.h>
#include "ReadOnlyEdit.h"
#include "PasswordEntry_Dlg.h"	
#include <direct.h>
#include <stdlib.h>
#include <stdio.h>
#include <fstream>

using namespace std;

char *INIpartNumberFile = "PartNumber.txt";
const char *INIoutFileName = "OutFile.txt";
const char *PasswordFilename = "PasswordFile.txt";
extern char AdminPassword[];

extern char *ptrExcelFilename;
extern char ExcelPrimaryFilename[];
extern char ExcelSecondaryFilename[];
extern char ExcelBackupFilename[];

extern DATAstring arrPartNumbers[];
extern int totalPartNumbers;
extern CTestDialog *ptrDialog;


extern DeviceConfigType* l_pDeviceData;

extern float AllowableLampVoltageError;
extern float AllowableVrefError;
extern float MinClosedFilterPeakAmplitude;
extern float MaxClosedFilterPeakAmplitude;
extern float MinClosedFilterWavelength;
extern float MaxClosedFilterWavelength;
extern float MinClosedFilterIrradiance;
extern float MaxClosedFilterIrradiance;
extern float MinClosedFilterFWHM;
extern float MaxClosedFilterFWHM;
extern float SpareIniValue;
extern float MinRemoteVoltLimit[];
extern float MaxRemoteVoltLimit[];

extern float MinOpenColorTemp;
extern float MaxOpenColorTemp;
extern float MinOpenIrradiance;
extern float MaxOpenIrradiance;

extern float	m_currentIntegrationTime;
extern uint16	m_NumberOfScans;
extern BOOL		m_enableLinearCorrection;
extern BOOL		m_enableStrayLightCorrection;
extern unsigned long m_numberOfAverages;

extern const char *INIbinaryFilename;
extern float arrINIconfigValues[];
extern long portNumberInterfaceBoard, portNumberACpowerSupply, portNumberMultiMeter;
extern float centerWavelength, amplitude, irradiance, FWHM, colorTemperature;
extern HANDLE handleInterfaceBoard, handleHPmultiMeter, handleACpowerSupply;

extern int subStepNumber;

extern BOOL lineOneEmpty;
extern BOOL lineTwoEmpty;
extern BOOL lineThreeEmpty;
extern BOOL lineFourEmpty;
extern BOOL lineFiveEmpty;

extern const char *CurrentDataFilename;
extern const char *CurrentDataPathname;
extern char *ptrExcelFilename;

int HiPotStatus, GroundBondStatus, FinalAssemblyStatus;
extern struct testData arrTestData[];

	// Resets flags and data array for tests. Used before starting test sequence on a new unit.
	void TestApp::resetTestData()
	{
		int i;
		for (i = 0; i < TOTAL_COLUMNS; i++)
		{
			arrTestData[i].result.SetString("");
			arrTestData[i].color = 0;
		}
		HiPotStatus = GroundBondStatus = FinalAssemblyStatus = NOT_DONE_YET;
	}
	
// This routine initializes the HP multimeter and is called once at startup
// If there is a problem with serial communication or the multimeter 
// doesn't respond to commands, FALSE is returned.
BOOL TestApp::InitializeHP34401() {
		char strReset[BUFFERSIZE] = "*RST\r\n";
		char strEnableRemote[BUFFERSIZE] = ":SYST:REM\r\n";
		char strMeasure[BUFFERSIZE] = ":MEAS?\r\n";

		// 1) Send RESET command to HP34401:
		if (!sendReceiveSerial(HP_METER, strReset, NULL, FALSE)) return FALSE;

		msDelay(500);

		// 2) Enable RS232 remote control on HP34401:
		if (!sendReceiveSerial(HP_METER, strEnableRemote, NULL, FALSE)) return FALSE;

		msDelay(500);

		// 3) Now try getting a measurement from the HP34401:
		if (!sendReceiveSerial(HP_METER, strMeasure, NULL, TRUE)) return FALSE;

		return TRUE;
}

// This routine initializes the Interface box by sending the RESET command
// This command sets the relays to default off states, 
// sets the PWM remote voltage control output to zero,
// and opens the filter actuator. 
// If there is a problem with serial communication, FALSE is returned.
BOOL TestApp::InitializeInterfaceBoard() {
#ifdef NOSERIAL
		return TRUE;
#endif
		char strReset[BUFFERSIZE] = "$RESET";
		char strResponse[BUFFERSIZE] = "";

		// Send RESET command to interface box:
		if (!sendReceiveSerial(INTERFACE_BOARD, strReset, strResponse, TRUE)) return FALSE;

		if (!strstr(strResponse, "OK")) return FALSE;		

		return TRUE;
}

// Ths routine sends commands to the programmable power supply to initialize it.
// The AC output voltage is set to zero, the remote RS232 control is enabled, 
// and a MEASure VOLTage command is sent to make sure supply is turned on and 
// serial port communication is working. If at any point the serial port 
// doesn't work, it returns FALSE to inidcate a system error.
BOOL TestApp::InitializePowerSupply() {
		char strCommand[BUFFERSIZE];
		BOOL deviceOK = TRUE;

		strcpy_s(strCommand, BUFFERSIZE, "*RST\r\n");
		if (!sendReceiveSerial(AC_POWER_SUPPLY, strCommand, NULL, FALSE)) return FALSE;
		msDelay(500);
		
		strcpy_s(strCommand, BUFFERSIZE, "*CLS\r\n");
		if (!sendReceiveSerial(AC_POWER_SUPPLY, strCommand, NULL, FALSE)) return FALSE;
		msDelay(500);

		strcpy_s(strCommand, BUFFERSIZE, "*SRE 128\r\n");
		if (!sendReceiveSerial(AC_POWER_SUPPLY, strCommand, NULL, FALSE)) return FALSE;
		msDelay(500);
			
		strcpy_s(strCommand, BUFFERSIZE, "*ESE 0\r\n");
		if (!sendReceiveSerial(AC_POWER_SUPPLY, strCommand, NULL, FALSE)) return FALSE;
		msDelay(500);		

		strcpy_s(strCommand, BUFFERSIZE, "VOLT:RANG 272\r\n");
		if (!sendReceiveSerial(AC_POWER_SUPPLY, strCommand, NULL, FALSE)) return FALSE;		
		msDelay(500);

		strcpy_s(strCommand, BUFFERSIZE, "OUTP 1\r\n");
		if (!sendReceiveSerial(AC_POWER_SUPPLY, strCommand, NULL, FALSE)) return FALSE;		
		msDelay(500);
		
		strcpy_s(strCommand, BUFFERSIZE, "SYST:REM\r\n");
		if (!sendReceiveSerial(AC_POWER_SUPPLY, strCommand, NULL, FALSE)) return FALSE;		
		msDelay(500);
		
		strcpy_s(strCommand, BUFFERSIZE, "VOLT 0\r\n");
		if (!sendReceiveSerial(AC_POWER_SUPPLY, strCommand, NULL, FALSE)) return FALSE;		
		msDelay(500);

		// Now send a MEASURE command to power supply
		// The TRUE input tells sendReceiveSerial() to expect a response
		// If there is no repsonse then sendReceiveSerial() returns FALSE:
		strcpy_s(strCommand, BUFFERSIZE, "MEAS:VOLT?\r\n");
		if (!sendReceiveSerial(AC_POWER_SUPPLY, strCommand, NULL, TRUE)) 
			return FALSE;		

		// Return TRUE to indicate power supply is communicating and functioning:		
		return TRUE;
}
	
// Thsi routine is called at startup. It opens the three serial ports
// and initializes the Interface box, the HP multimeter, the programmable power supply,
// and the spectrometer. If any serial port cannot be opened, or if any 
// of these devices doesn't communicate, it displays an error message
// in the main window and returns FALSE.
BOOL TestApp::InitializeSystem (bool UseSpectrometer) 
{
		static int retry = 1;
#ifdef NOSERIAL
		return TRUE;
#endif
		char portNameInterfaceBoard[MAXSTRING], portNameACpowerSupply[MAXSTRING], portNameMultiMeter[MAXSTRING];

		portNumberInterfaceBoard = (long) arrINIconfigValues[0];
		if (portNumberInterfaceBoard < 1 || portNumberInterfaceBoard > MAX_PORTNUMBER)
		return FALSE;

		portNumberACpowerSupply = (long) arrINIconfigValues[1];
		if (portNumberACpowerSupply < 1 || portNumberACpowerSupply > MAX_PORTNUMBER)
		return FALSE;

		portNumberMultiMeter = (long) arrINIconfigValues[2];
		if (portNumberMultiMeter < 1 || portNumberMultiMeter > MAX_PORTNUMBER)
		return FALSE;
		
		sprintf_s(portNameInterfaceBoard, MAXSTRING, "\\\\.\\COM%d", (int) portNumberInterfaceBoard);
		sprintf_s(portNameACpowerSupply, MAXSTRING, "\\\\.\\COM%d", (int) portNumberACpowerSupply);
		sprintf_s(portNameMultiMeter, MAXSTRING, "\\\\.\\COM%d", (int) portNumberMultiMeter);

		if (!openSerialPort(portNameInterfaceBoard, &handleInterfaceBoard))
		{
			writeInstructionLine(1, MAXLINE, "INTERFACE BOX SERIAL COM ERROR");			
			writeInstructionLine(2, MAXLINE, "Check USB connection to interface box");
			writeInstructionLine(3, MAXLINE, "Make sure interface box is plugged in");
			writeInstructionLine(4, MAXLINE, "and turned on. Then press ENTER.");
			writeInstructionLine(5, MAXLINE, "If error still occurs, close program then restart.");
			return FALSE;
		}
		else if (!openSerialPort(portNameMultiMeter, &handleHPmultiMeter))
		{
			writeInstructionLine(1, MAXLINE, "MULTIMETER SERIAL COM ERROR");
			writeInstructionLine(2, MAXLINE, "Check USB connection to multimeter");
			writeInstructionLine(3, MAXLINE, "Make sure multimeter is turned on.");
			writeInstructionLine(4, MAXLINE, "Press ENTER to try again.");
			writeInstructionLine(5, MAXLINE, "If error still occurs, close program then restart.");
			return FALSE;
		}
		else if (!openSerialPort(portNameACpowerSupply, &handleACpowerSupply))
		{			
			writeInstructionLine(1, MAXLINE, "POWER SUPPLY SERIAL COM ERROR");
			writeInstructionLine(2, MAXLINE, "Make sure power switch on back of supply");
			writeInstructionLine(3, MAXLINE, "is turned on. Also check serial cable.");
			writeInstructionLine(4, MAXLINE, "Press ENTER to try again.");
			writeInstructionLine(5, MAXLINE, "If error still occurs, close program then restart.");
			return FALSE;
		}
		else if (!InitializeHP34401()) return FALSE;
		else if (!InitializeInterfaceBoard()) return FALSE;
			
#ifndef NOPOWERSUPPLY
		else if (!InitializePowerSupply())
		{			
			writeInstructionLine(1, MAXLINE, "POWER SUPPLY SERIAL COM ERROR");
			writeInstructionLine(2, MAXLINE, "Make sure power switch on back of supply");
			writeInstructionLine(3, MAXLINE, "is turned on. Also check serial cable.");
			writeInstructionLine(4, MAXLINE, "Press ENTER to try again.");
			writeInstructionLine(5, MAXLINE, "If error still occurs, close program then restart.");
			return FALSE;
		}
#endif
		else if (UseSpectrometer)
		{
			if (!InitializeSpectrometer()) 
			{				
				writeInstructionLine(1, MAXLINE, "SPECTROMETER COMMUNICATION ERROR");
				writeInstructionLine(2, MAXLINE, "Make sure USB cable is connected to spectrometer");
				writeInstructionLine(3, MAXLINE, "If problem still occurs, unplug then plug in spectroemeter USB.");								
				writeInstructionLine(4, MAXLINE, "");
				writeInstructionLine(5, MAXLINE, "");
				return FALSE;
			}
		}
		return TRUE;				
	}


//  This routine stores the variables which set test limits and also serial port assignments.
//  These variables are stored in the files INIbinaryFilename which is "StartupFile.bin".
//	
void TestApp::storeINIfileBinary(){	
	float INIconfigValue = 0;
	int i = 0;

	arrINIconfigValues[0] = (float) portNumberInterfaceBoard;
	arrINIconfigValues[1] = (float) portNumberACpowerSupply;
	arrINIconfigValues[2] = (float) portNumberMultiMeter;

	arrINIconfigValues[3] = AllowableLampVoltageError;
	arrINIconfigValues[4] = AllowableVrefError;
	arrINIconfigValues[5] = MinClosedFilterPeakAmplitude;
	arrINIconfigValues[6] = MaxClosedFilterPeakAmplitude;
	arrINIconfigValues[7] = MinClosedFilterWavelength;
	arrINIconfigValues[8] = MaxClosedFilterWavelength;
	arrINIconfigValues[9] = MinClosedFilterIrradiance;
	arrINIconfigValues[10] = MaxClosedFilterIrradiance;
	arrINIconfigValues[11] = MinClosedFilterFWHM;
	arrINIconfigValues[12] = MaxClosedFilterFWHM;
	arrINIconfigValues[13] = SpareIniValue;
	arrINIconfigValues[14] = MinOpenColorTemp;
	arrINIconfigValues[15] = MaxOpenColorTemp;
	arrINIconfigValues[16] = MinOpenIrradiance;
	arrINIconfigValues[17] = MaxOpenIrradiance;

	arrINIconfigValues[18] = (float) m_currentIntegrationTime;
	arrINIconfigValues[19] = (float) m_NumberOfScans;

	if (m_enableLinearCorrection) arrINIconfigValues[20] = (float) 9999;
	else arrINIconfigValues[20] = (float) 0.0;

	arrINIconfigValues[21] = (float) m_numberOfAverages;

	if (m_enableStrayLightCorrection) arrINIconfigValues[22] = (float) 9999;
	else arrINIconfigValues[22] = (float) 0.0;

		// Min and Max Remote voltage test limits
	arrINIconfigValues[23] = MinRemoteVoltLimit[0];
	arrINIconfigValues[24] = MinRemoteVoltLimit[1];
	arrINIconfigValues[25] = MinRemoteVoltLimit[2];
	arrINIconfigValues[26] = MinRemoteVoltLimit[3];
	arrINIconfigValues[27] = MinRemoteVoltLimit[4];
	arrINIconfigValues[28] = MinRemoteVoltLimit[5];
	arrINIconfigValues[29] = MinRemoteVoltLimit[6];
	arrINIconfigValues[30] = MinRemoteVoltLimit[7];

	arrINIconfigValues[31] = MaxRemoteVoltLimit[0];
	arrINIconfigValues[32] = MaxRemoteVoltLimit[1];
	arrINIconfigValues[33] = MaxRemoteVoltLimit[2];
	arrINIconfigValues[34] = MaxRemoteVoltLimit[3];
	arrINIconfigValues[35] = MaxRemoteVoltLimit[4];
	arrINIconfigValues[36] = MaxRemoteVoltLimit[5];
	arrINIconfigValues[37] = MaxRemoteVoltLimit[6];
	arrINIconfigValues[38] = MaxRemoteVoltLimit[7];


	std::ofstream outFile;
	// Open file to create it
	outFile.open(INIbinaryFilename, ios::out|ios::binary|ios::trunc);
	if (!outFile.is_open()) 
		throw ios::failure(string("Error opening INI file ") + 
	string(INIbinaryFilename) +
	string("  in main()"));
	i = 0;
	do {
		INIconfigValue = arrINIconfigValues[i];
		writeBinaryFile(outFile, INIconfigValue);
		i++;
	} while (i < MAXINIVALUES);
	outFile.close();				
}



BOOL TestApp::loadINIfileBinary(){
	std::ifstream inFile;
	float INIconfigValue = 0;

	inFile.open(INIbinaryFilename, ios::in|ios::binary); // Open INI file		
	if (inFile.is_open()){
		int i = 0;
		do {
			readBinaryFile(inFile, INIconfigValue);
			arrINIconfigValues[i] = INIconfigValue;
			i++;
		} while (!inFile.eof() && i < MAXINIVALUES);
		inFile.close();

		portNumberInterfaceBoard = (long) arrINIconfigValues[0];
		portNumberACpowerSupply = (long) arrINIconfigValues[1];
		portNumberMultiMeter = (long) arrINIconfigValues[2];	
		AllowableLampVoltageError = arrINIconfigValues[3];
		AllowableVrefError = arrINIconfigValues[4];
		MinClosedFilterPeakAmplitude = arrINIconfigValues[5];
		MaxClosedFilterPeakAmplitude = arrINIconfigValues[6];
		MinClosedFilterWavelength = arrINIconfigValues[7];
		MaxClosedFilterWavelength = arrINIconfigValues[8];
		MinClosedFilterIrradiance = arrINIconfigValues[9];
		MaxClosedFilterIrradiance = arrINIconfigValues[10];
		MinClosedFilterFWHM = arrINIconfigValues[11];
		MaxClosedFilterFWHM = arrINIconfigValues[12];
		SpareIniValue = arrINIconfigValues[13];
		MinOpenColorTemp = arrINIconfigValues[14];
		MaxOpenColorTemp = arrINIconfigValues[15];
		MinOpenIrradiance = arrINIconfigValues[16];
		MaxOpenIrradiance = arrINIconfigValues[17];

		m_currentIntegrationTime = (float) arrINIconfigValues[18];
		m_NumberOfScans = (uint16) arrINIconfigValues[19];

		if (arrINIconfigValues[20] == (float) 0.0) m_enableLinearCorrection = FALSE;
		else m_enableLinearCorrection = TRUE;

		m_numberOfAverages = (unsigned long) arrINIconfigValues[21];

		if (arrINIconfigValues[22] == (float) 0.0) m_enableStrayLightCorrection = FALSE;
		else m_enableStrayLightCorrection = TRUE;

		if (i >= MAXINIVALUES)
		{
			// Min and Max Remote voltage test limits
			MinRemoteVoltLimit[0] = arrINIconfigValues[23];
			MinRemoteVoltLimit[1] = arrINIconfigValues[24];
			MinRemoteVoltLimit[2] = arrINIconfigValues[25];
			MinRemoteVoltLimit[3] = arrINIconfigValues[26];
			MinRemoteVoltLimit[4] = arrINIconfigValues[27];
			MinRemoteVoltLimit[5] = arrINIconfigValues[28];
			MinRemoteVoltLimit[6] = arrINIconfigValues[29];
			MinRemoteVoltLimit[7] = arrINIconfigValues[30];

			MaxRemoteVoltLimit[0] = arrINIconfigValues[31];
			MaxRemoteVoltLimit[1] = arrINIconfigValues[32];
			MaxRemoteVoltLimit[2] = arrINIconfigValues[33];
			MaxRemoteVoltLimit[3] = arrINIconfigValues[34];
			MaxRemoteVoltLimit[4] = arrINIconfigValues[35];
			MaxRemoteVoltLimit[5] = arrINIconfigValues[36];
			MaxRemoteVoltLimit[6] = arrINIconfigValues[37];
			MaxRemoteVoltLimit[7] = arrINIconfigValues[38];			
		}	
		else storeINIfileBinary();		
	}
	else {
		//storeINIfileBinary();
		//AfxMessageBox((LPCTSTR)"SYSTEM ERROR - No INI file, creating default INI FILE");
		//return TRUE;
		return FALSE;
	}
	return TRUE;
}

// Read binary float value
void TestApp::readBinaryFile(std::istream& in, float& value){
	in.read(reinterpret_cast<char *>(&value), sizeof(value));
}
// Write binary float value
void TestApp::writeBinaryFile(std::ostream& out, float value){
	out.write(reinterpret_cast<char *>(&value), sizeof(value));
}


BOOL TestApp::loadINIpartNumbers()
{
std::string strLine;
char *ptrLine;
int i = 0, j = 0;

	if (INIpartNumberFile == NULL) return FALSE;
	std::ifstream inFile(INIpartNumberFile);	
	if (inFile) {	
		i = 0;
		while (std::getline(inFile, strLine))
		{
			ptrLine = strdup(strLine.c_str());						
			memcpy(arrPartNumbers[i].String, ptrLine, VALID_PART_NUMBER_LENGTH);
			free (ptrLine);
			arrPartNumbers[i].String[VALID_PART_NUMBER_LENGTH] = '\0';
			i++;
			if (i >= MAX_PART_NUMBERS) break;
		}	
		totalPartNumbers = i;		
		return TRUE;
	}
	else return FALSE;
}

BOOL TestApp::storeINIpartNumbers()
{
int i;

ofstream myfile;

	if (INIpartNumberFile == NULL) return FALSE;
	if (totalPartNumbers > MAX_PART_NUMBERS) totalPartNumbers = MAX_PART_NUMBERS;

	myfile.open(INIpartNumberFile);
	for (i = 0; i < totalPartNumbers; i++) myfile << arrPartNumbers[i].String << "\r\n";
	
	myfile.close();
	return TRUE;
}


   
BOOL TestApp::CheckAndCreateDirectory (char *path){
	if (GetFileAttributes(path) == INVALID_FILE_ATTRIBUTES) {
		CreateDirectory(path,NULL);
		return TRUE;
	}
	else {
		return FALSE;
	}
}


void TestApp::ShutDown(BOOL copySpreadsheet)
{
	SaveSpreadsheet();		
	
	if (copySpreadsheet)
	{
		CopyFile(CurrentDataFilename, ExcelBackupFilename, FALSE);
		CopyFile(CurrentDataFilename, ptrExcelFilename,  FALSE);
	}
	DeleteFile((LPCSTR)CurrentDataFilename);
	closeAllSerialPorts();		
	CloseSpectrometer();
	delete l_pDeviceData;
}




BOOL TestApp::isValidPassword(char *ptrPassword)
{
std::string strLine;
char ch, strTemp[MAXPASSWORD];
int i = 0, j = 0;
	
	if (ptrPassword == NULL) return FALSE;
	i = 0; j = 0;

	do {
		ch = ptrPassword[i];
		if (ch != ' ') break;
		i++;
	} while (i < MAXPASSWORD);

	do {
		ch = ptrPassword[i];
		if (ch == ' ') return FALSE;
		else if (isalnum(ch) || ispunct(ch) || ch == '_') strTemp[j++] = ch;
		else if (ch =='\0' || ch == '\r' || ch == '\n') break;
		else return FALSE;
		i++;
	} while (i < MAXPASSWORD);

	if (j < MAXPASSWORD && j >= MINPASSWORD) 
	{
		strTemp[j] = '\0';
		strcpy_s(ptrPassword, MAXPASSWORD, strTemp);
		return TRUE;
	}
	return FALSE;
}

BOOL TestApp::LoadAdminPassword()
{
std::string strLine;
char *ptrLine;
BOOL result = TRUE;
char strTemp[MAXPASSWORD];
	
	if (PasswordFilename == NULL) return FALSE;

	std::ifstream inFile(PasswordFilename);		
	
	
	if (inFile)
	{
		if (std::getline(inFile, strLine))
		{
			ptrLine = strdup(strLine.c_str());
			strcpy_s(strTemp, MAXPASSWORD, ptrLine);
			free (ptrLine);

			if (!isValidPassword(strTemp)) result = FALSE;
			else 
			{
				strcpy_s(AdminPassword, MAXPASSWORD, strTemp);	
				result = TRUE; // Success - password is loaded and is valid
			}
		}	
		else result = FALSE;
		inFile.close();		
	}
	else result = FALSE;
	return result;
}

BOOL TestApp::storeAdminPassword()
{
ofstream myfile;

	if (PasswordFilename == NULL) return FALSE;

	myfile.open(PasswordFilename);
	myfile << AdminPassword;
	
	myfile.close();
	return TRUE;
}

// This routine is called once after startup to initialize the spreadsheets.
// All test data is stored in the Primary spreadsheet and also the Backup.
// If the Primary couldn't be located, a Secondary spreadsheet is used instead.
//
// Once the primary spreadsheet is located, all existing test data is copied to a new spreadsheet.		
// This is the working copy which records new test data, 
// appending it to the existing data after each new unit is tested.
// The working copy then overwrites the primary spreadsheet.
// In the event that the PC crashes or loses power 
// while writing data to the working spreadsheet,
// it becomes corrupted and unusable.
// This is why only the working spreadsheet is opened while program is running.
BOOL TestApp::InitializeSpreadsheets()
{
	if (!validateTempPath()) return FALSE;

	// Try to locate spreadsheet:	
	if (GetFileAttributes(ptrExcelFilename) == INVALID_FILE_ATTRIBUTES)
	{		
		// If spreadsheet can't be located either because it doesn't exist
		// or because directory can't be accessed i.e. network is down,
		// then spreadsheet gets renamed here to be stored 
		// in local subdirectory along with program executable.
		// Filename will now include date to help trace problem later:
		createSecondarySpreadsheetName(ExcelPrimaryFilename, ExcelSecondaryFilename);
		ptrExcelFilename = ExcelSecondaryFilename;
		storeExcelFilenames(ExcelPrimaryFilename, ExcelSecondaryFilename);

		// Now try to get data from backup spreadsheet:
		if (GetFileAttributes(ExcelBackupFilename) == INVALID_FILE_ATTRIBUTES)
		{
			// If backup spreadsheet also can't be found,
			// then a new spreadsheet is created here from scratch,
			// and is copied to both main and backup spreadsheets.
			// Unfortunately this means that all data from previous tests is lost!
			//
			// The new file is also the working spreadsheet,
			// which is used to record data as test progresses.
			// The working spreadsheet always has the same name,
			// which is CurrentDataFilename:
			if (!createNewSpreadsheet()) return FALSE;			
			CopyFile(CurrentDataFilename, ExcelBackupFilename, FALSE);
			CopyFile(CurrentDataFilename, ptrExcelFilename, FALSE);			
		}
		else 
		{
			// Otherwise if backup spreadsheet is located, 
			// copy it to original and working spreadsheets:
			CopyFile(ExcelBackupFilename, ptrExcelFilename, FALSE);
			CopyFile(ExcelBackupFilename, CurrentDataFilename, FALSE);
		}
	}
	// Otherwise, if original spreadsheet exists, back it up.
	// Also create working copy "CurrentDataFilename" with same data:
	else 
	{
		CopyFile(ptrExcelFilename, ExcelBackupFilename, FALSE);
		CopyFile(ptrExcelFilename, CurrentDataFilename, FALSE);
	}

	
	// If Primary spreadsheet isn't being used, something is wrong, so notify system administrator:
	if (ptrExcelFilename != ExcelPrimaryFilename)
	{
		char strMessage[MAXFILENAME+64];
		sprintf_s(strMessage, MAXFILENAME+64, "Warning: a backup spreadsheet is being used\r\nin place of the primary spreadsheet.\r\nPlease notify system administrator.", ptrExcelFilename);
		MessageBox((LPCTSTR) strMessage, (LPCTSTR)"BACKUP SPREADSHEET WARNING", MB_ICONWARNING | MB_OKCANCEL | MB_DEFBUTTON2);
	}

	char filenameText[MAXFILENAME+16];
	sprintf_s(filenameText, MAXFILENAME+16, "Spreadsheet: %s", ptrExcelFilename);
	ptrDialog->m_static_SpreadsheetFilename.SetWindowText(filenameText);
	return TRUE;
}

int TestApp::LoadExcelFilenames(char *ptrPrimary, char *ptrSecondary)
{
std::string strLine;
char *ptrLine = NULL, ch;
char filename[MAXFILENAME];
int numFilenames = 0, i = 0;	
	
	if (INIoutFileName == NULL) return 0;  // Normally this should never happen!

	ptrSecondary[0] = '\0';

	std::ifstream inFile(INIoutFileName);		
	
	if (inFile)
	{
		numFilenames = 0;
		while (std::getline(inFile, strLine) && numFilenames < 2)
		{
			ptrLine = strdup(strLine.c_str());			
			if (ptrLine != NULL)
			{
				for (i = 0; i < MAXFILENAME; i++)
				{
					ch = ptrLine[i];
					if (ch == '\0' || ch == '\r' || ch == '\n')
					{
						filename[i] = '\0';
						break;
					}					
					filename[i] = ch;
					if (ptrLine[i] == '.' && ptrLine[i+1] == 'x' && ptrLine[i+2] == 'l' && ptrLine[i+3] == 's')
					{
						filename[i+1] = 'x';
						filename[i+2] = 'l';
						filename[i+3] = 's';
						filename[i+4] = '\0';
						numFilenames++;
						if (numFilenames == 1) strcpy_s(ptrPrimary, MAXFILENAME, filename);
						else if (numFilenames == 2) strcpy_s(ptrSecondary, MAXFILENAME, filename);					
						break;						
					}
				}
				std::free(ptrLine);	
			}				
		}
		inFile.close();
	}	
	else return(0);
	return numFilenames;
}


BOOL TestApp::storeExcelFilenames(char *ptrPrimaryFilename,  char *ptrSecondaryFilename)
{
ofstream myfile;

	if (INIoutFileName == NULL) return FALSE;

	myfile.open(INIoutFileName);
	myfile << ptrPrimaryFilename << "\r\n";
	myfile << ptrSecondaryFilename << "\r\n";
	
	myfile.close();
	return TRUE;
}


BOOL TestApp::createSecondarySpreadsheetName(char *ptrPrimary, char *ptrSecondary)
{
char dateBuffer[32], ch;
struct tm  tstruct;	
int i, j;

	// Make sure filename array exists. If it doesn't, something is wrong!
	if (ptrPrimary == NULL) return FALSE;

	// Otherwise if filename isn't valid, create default name:
	if (!strstr(ptrPrimary, ".xls")) 
		strcpy_s(ptrPrimary, MAXFILENAME, "DC950TestData.xls");

	// If secondary filename already exists, don't create a new one:
	if (strstr(ptrSecondary, ".xls")) return TRUE;

	j = 0;
	for (i = 0; i < MAXFILENAME; i++)	
	{
		ch = ptrPrimary[i];
		if (ch == '\0' || ch == '\r' || ch == '\n') break;
		else if (ch == '\\') j = i + 1;
	}	
	for (i = 0; i <MAXFILENAME; i++)
	{
		if (j > MAXFILENAME-5) {
			ptrSecondary[0] = '\0';
			return FALSE;
		}
		if (ptrPrimary[j] == '.' && ptrPrimary[j+1] == 'x' && ptrPrimary[j+2] == 'l' && ptrPrimary[j+3] == 's')
			break;
		else ptrSecondary[i] = ptrPrimary[j++];
	}
	if (i == MAXFILENAME)
	{
		ptrSecondary[0] = '\0';
		return FALSE;
	}
	ptrSecondary[i] = '\0';
	time_t now = time(0);
	tstruct = *localtime(&now);
	strftime(dateBuffer, sizeof(dateBuffer), "[%m-%d-%y].xls", &tstruct);
	strcat_s(ptrSecondary, MAXFILENAME, dateBuffer);

	return TRUE;
}

BOOL TestApp::validateTempPath()
{
	if (GetFileAttributes(CurrentDataPathname) != INVALID_FILE_ATTRIBUTES) 
		return TRUE;
	if (_mkdir(CurrentDataPathname) == 0 ) 
		return TRUE;
	else 
		return FALSE;
}