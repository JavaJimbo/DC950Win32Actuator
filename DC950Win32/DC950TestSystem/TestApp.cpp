// COMMENTED
/* TestApp.cpp - for DC950 test system
 * 
 * This file includes Execute(), which implements the test sequence
 * 2-9-2018 JBS: Added comments
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
//#include "BasicExcel.hpp"
//#include "ExcelFormat.h"
#include <time.h>
#include "ReadOnlyEdit.h"
#include "PasswordEntry_Dlg.h"	
#include <ole2.h> // OLE2 Definitions
#include "MSExcel.h"

extern CMSExcel MyExcelApp;

float AllowableLampVoltageError = (float) 0.1;
float AllowableVrefError = (float) 0.1;
float MinClosedFilterPeakAmplitude = (float) 100.0;
float MaxClosedFilterPeakAmplitude = (float) 200.0;
float MinClosedFilterWavelength = (float) 938.0;
float MaxClosedFilterWavelength = (float) 943.0;
float MinClosedFilterIrradiance = (float) 200.0;
float MaxClosedFilterIrradiance = (float) 300.0;
float MinClosedFilterFWHM = (float) 10.0;
float MaxClosedFilterFWHM = (float) 18.0;
float SpareIniValue = (float) 5;

//0  	0.0-0.2
//1		3.8-4.4
//2		8.2-8.5
//3		12.5-12.7
//4		16.7-16.9
//5 	20.9-21.1

float MinRemoteVoltLimit[NUM_REMOTE_TEST_VALUES] = {(float)0.0, (float)0.0, (float)3.8, (float)8.2, (float)12.5, (float)16.7, (float)20.9, (float)20.9};
float MaxRemoteVoltLimit[NUM_REMOTE_TEST_VALUES] = {(float)0.2, (float)0.2, (float)4.4, (float)8.5, (float)12.7, (float)16.9, (float)21.1, (float)21.1};

float MinOpenColorTemp = (float) 100.0;
float MaxOpenColorTemp = (float) 200.0;
float MinOpenIrradiance = (float) 200.0;
float MaxOpenIrradiance = (float) 300.0;

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

const char *INIbinaryFilename = "StartupFile.bin";
float arrINIconfigValues[MAXINIVALUES];
long portNumberInterfaceBoard = 3, portNumberACpowerSupply = 4, portNumberMultiMeter = 7;

float openColorTemp, openIrradianceIntegral, closedIrradianceIntegral;	
float centerWavelength, wavelengthAmplitude, FWHM;

extern BOOL runFilterActuatorTest;

extern BOOL lineOneEmpty;
extern BOOL lineTwoEmpty;
extern BOOL lineThreeEmpty;
extern BOOL lineFourEmpty;
extern BOOL lineFiveEmpty;

extern const char *CurrentDataFilename;
extern char *ptrExcelFilename;
extern char ExcelBackupFilename[];


HANDLE handleInterfaceBoard = NULL, handleHPmultiMeter = NULL, handleACpowerSupply = NULL;
CFont BigFont, SmallFont, MidFont;
int subStepNumber = 0;
extern int HiPotStatus, GroundBondStatus, FinalAssemblyStatus;

	// CONSTRUCTORS FOR TestApp
	TestApp::~TestApp() {
	}

	TestApp::TestApp(CWnd* pParent) {				
	}

	
	//  RunPowerSupplyTest() - implements sequence for testing DC950 lamp voltage output
	//	while adjusting variable power supply to AC test voltages: 
	//	90, 115, 132, 254, 230, 180 volts AC.
	//  The first three voltages are at 60 Hz, the second three are at 50 Hz.
	//
	//  The multimeter measures the lamp voltage each time through
	//  and the voltage is checked to see if it is within limits.
	//  It returns NOT_DONE_YET each time through until
	//  the full sequence has completed, at which point it returns PASS.
	//	If lamp voltage is outside of limits at any point,
	//  it returns FAIL.

	int TestApp::RunPowerSupplyTest()
	{
#ifdef NOSERIAL
		return PASS;
#endif

#define NUM_TEST_VOLTAGES 6
		int supplyVoltage, frequency;
		int arrSupplyVoltage[NUM_TEST_VOLTAGES] = { 90, 115, 132, 254, 230, 180}; // Array of output AC test voltages 
		int arrSupplyFrequency[NUM_TEST_VOLTAGES] = {60, 60, 60, 50, 50, 50}; // Array of output AC frequencies in Hz
		float lampVoltage, errorVoltage, expectedVoltage, actualControlVoltage;
		int testResult = PASS;
		char strAddLog[MAXSTRING];

		// If this is the first time through (subStepNumber = 0) then
		// set the remote control voltage input to the DC950
		// to five volts (MAX PWM) for full lamp instensity
		// Also check the Fault Signal output on the DC950.
		// It should always be low when the EDC950 is on:
		if (subStepNumber == 0) {
			if (!InitializeInterfaceBoard()) return SYSTEM_ERROR;
			msDelay(100);
			if (!SetInterfaceBoardPWM(MAX_PWM)) return SYSTEM_ERROR;
			msDelay(100);
			if (!CheckFaultSignal(&testResult)) return SYSTEM_ERROR; 
			if (testResult == PASS)	storeTestResult(FAULT_TEST, PASS, NULL);
			else
			{
				storeTestResult(FAULT_TEST, FAIL, "LAMP FAULT");
				strcpy_s(strAddLog, MAXSTRING, "Lamp fault signal test FAIL - lamp error detected\r\n");
				DisplayLog(strAddLog);	
				return FAIL;
			}
			strcpy_s(strAddLog, MAXSTRING, "Lamp fault signal test: PASS\r\n");
			DisplayLog(strAddLog);				
			msDelay(100);
		}

		if (subStepNumber > NUM_TEST_VOLTAGES) return SYSTEM_ERROR;
		if (subStepNumber < 0) return SYSTEM_ERROR;

		// Set programmable power supply voltage and frequency to next setpoint:
		supplyVoltage = arrSupplyVoltage[subStepNumber];
		frequency = arrSupplyFrequency[subStepNumber];
		if (!SetPowerSupplyVoltage(supplyVoltage, frequency)) return SYSTEM_ERROR;

		msDelay(500);
		// Check the remote control voltage going out to the DC950.
		// It should be close to 5 volts:
		if (!ReadVoltage(CONTROL_VOLTAGE, &actualControlVoltage)) return SYSTEM_ERROR;

			errorVoltage = getAbs(actualControlVoltage - (float) 5.0);
			if (errorVoltage > MAX_CONTROL_VOLTAGE_ERROR)
			{			
				writeInstructionLine(3, MAXLINE, "Either Multimeter or Interface box isn't working.");
				writeInstructionLine(4, MAXLINE, "Check connection on meter and interface box");
				writeInstructionLine(5, MAXLINE, "Also switch multimeter off, then on.");
				return SYSTEM_ERROR;
			}

		// The expected DC950 lamp voltage should be 4.2 times the remote control voltage:
		expectedVoltage = actualControlVoltage * (float) 4.2;
		msDelay(100);
		
		if (!ReadVoltage(LAMP, &lampVoltage)) return SYSTEM_ERROR;
		errorVoltage = getAbs(lampVoltage - expectedVoltage);
		// If DC950 lamp voltage isn't within limits, test FAILs:
		if (errorVoltage > AllowableLampVoltageError)
			testResult = FAIL;
		
		float minVoltage = expectedVoltage - AllowableLampVoltageError;
		float maxVoltage = expectedVoltage + AllowableLampVoltageError;			
		
		// Display test results in test log window:
		sprintf_s(strAddLog, MAXSTRING, "AC Sweep %d volt: MIN %.2f < %.2f < MAX %.2f:", supplyVoltage, minVoltage, lampVoltage, maxVoltage);
		if (testResult == PASS) strcat_s(strAddLog, MAXSTRING, "    PASS\r\n");
		else strcat_s(strAddLog, MAXSTRING, "    FAIL\r\n");
		DisplayLog(strAddLog);			

		char strMeasuredVoltage[MAXSTRING];
		sprintf_s(strMeasuredVoltage, MAXSTRING, "%.2f", lampVoltage);
		// Record test results in spreadsheet:
		storeTestResult(subStepNumber + SWEEP_0, testResult, strMeasuredVoltage);

		if (testResult == FAIL) return (FAIL);

		// Increment the subStep index to test another voltage next time through.
		// If all six voltages have been tested, return PASS.
		// Otherwise if there are more voltages to test, return NOT DONE YET:
		subStepNumber++;
		if (subStepNumber >= NUM_TEST_VOLTAGES) 
		{			
			if (!SetPowerSupplyVoltage(120, 60)) return SYSTEM_ERROR;
			if (!InitializeInterfaceBoard()) return SYSTEM_ERROR;
			return (PASS);
		}
		else return NOT_DONE_YET;
	}


	// Ths routine tests three remote control features on the DC950:
	// the 5V reference voltage output, 
	// the remote pot voltage control input,
	// and the remote relay inhibit input.
	//
	// For the latter two tests, the lamp voltage is measured 
	// to ensure it is proportional to the pot input voltage. 
	// Each time through, a different remote voltage  
	// is used until the test sequence is completed
	// for control voltages from 0-5 VDC.
	// 
	// The first and last voltage in the sequence are repeated:
	// a dummy voltage of zero volts is used first time through
	// when the 5V reference can be checked,
	// and five volts is used the last time through
	// when the relay inhibit signal is checked. 
	//
	// The multimeter measures the lamp voltage each time 
	// and and test passes if it is within limits.
	//
	// This routine returns NOT_DONE_YET each time through until
	// the full sequence has completed, at which point it returns PASS.
	// However, unit fails at any point, FAIL is returned.
	int TestApp::RunRemoteTests()
	{

#define LAST_TEST (NUM_REMOTE_TEST_VALUES - 1)

#ifdef NOSERIAL
		return PASS;
#endif

		char strAddLog[MAXSTRING];
		int testPWM;
		// Array of PWM values to produce the desired remote pot voltage
		// PWM = 1243 yields 1 volt approximately
		// PWM = 2486 yields 2 volts
		// PWM = 3729 yields 3 volts
		// PWM = 4972 yields 4 volts
		// PWM = 6215 yields 5 volts
		// int PWMvalues[NUM_REMOTE_TEST_VALUES] = { 0, 0, 1243, 2486, 3729, 4972, 6215, 6215 };  
		int PWMvalues[NUM_REMOTE_TEST_VALUES] = { 0, 0, 1238, 2501, 3764, 5026, 6215, 6215 };  
		float expectedControlVoltage[NUM_REMOTE_TEST_VALUES] = { 0, 0, 1,2,3,4,5,5};
		float measuredVoltage, actualControlVoltage, expectedVoltage, errorVoltage, maxVoltage, minVoltage;
		float offsetVoltage;
		int testResult = PASS;

		if (subStepNumber > NUM_REMOTE_TEST_VALUES) return SYSTEM_ERROR;
		if (subStepNumber < 0) return SYSTEM_ERROR;
		// Get next PWM value to set Interface box voltage output:
		testPWM = PWMvalues[subStepNumber];

		// If this is the first Remote voltage test,
		// make sure the INHIBIT relay on the interface box is OFF,
		// and set the expected voltage to 5.0 volts
		// so that VREF can be checked
		if (subStepNumber == 0) {
			if (!SetInhibitRelay(FALSE)) return SYSTEM_ERROR;
		}

		msDelay(100);
		// Set remote control voltage output on the Interface box:
		if (!SetInterfaceBoardPWM(testPWM)) return SYSTEM_ERROR;		

		// FIRST TEST IN SEQUENCE: CHECK DC950 5V VREF VOLTAGE OUTUPUT:
		if (subStepNumber == 0) 
		{
			msDelay(500);

			expectedVoltage = (float) 5.0;
			minVoltage = expectedVoltage - AllowableVrefError;
			maxVoltage = expectedVoltage + AllowableVrefError;	

			if (!ReadVoltage(VREF, &measuredVoltage)) return SYSTEM_ERROR;
			errorVoltage = getAbs(measuredVoltage - expectedVoltage);
			if (errorVoltage > AllowableVrefError) testResult = FAIL;
				
			sprintf_s(strAddLog, MAXSTRING, "5V Reference:      MIN %.2f < %.2f < MAX %.2f:", minVoltage, measuredVoltage, maxVoltage);
			if (testResult == PASS) strcat_s(strAddLog, MAXSTRING, "    PASS\r\n");
			else strcat_s(strAddLog, MAXSTRING, "    FAIL\r\n");
			DisplayLog(strAddLog);			
		}
		// TEST LAMP OUTPUT FOR REMOTE POT CONTROL VOLTAGES 0-5 VOLTS:
		else if (subStepNumber < LAST_TEST) 
		{
			msDelay(500);  
			// First make sure Interface box is working. Check remote control voltage output:
			if (!ReadVoltage(CONTROL_VOLTAGE, &actualControlVoltage)) return SYSTEM_ERROR;
			errorVoltage = getAbs(actualControlVoltage - expectedControlVoltage[subStepNumber]);
			if (errorVoltage > MAX_CONTROL_VOLTAGE_ERROR)
			{					
				writeInstructionLine(3, MAXLINE, "Either Multimeter or Interface box isn't working.");
				writeInstructionLine(4, MAXLINE, "Check connection on meter and interface box");
				writeInstructionLine(5, MAXLINE, "Also switch multimeter off, then on.");
				return SYSTEM_ERROR;
			}
			offsetVoltage = (actualControlVoltage - expectedControlVoltage[subStepNumber]) * (float) 4.2;
			minVoltage = MinRemoteVoltLimit[subStepNumber] + offsetVoltage;
			maxVoltage = MaxRemoteVoltLimit[subStepNumber] + offsetVoltage;
			msDelay(100);
			// Now measure DC950 lamp voltage:
			if (!ReadVoltage(LAMP, &measuredVoltage)) return SYSTEM_ERROR;
			if (subStepNumber == 1) minVoltage = (float) 0.0;
			
			// If DC950 lamp voltage is outside of test limits, return FAIL:
			if (measuredVoltage < minVoltage || measuredVoltage > maxVoltage) testResult = FAIL;
			
			// Display results in test log box:
			sprintf_s(strAddLog, MAXSTRING, "Remote %d volt:      MIN %.2f < %.2f < MAX %.2f:", subStepNumber-1, minVoltage, measuredVoltage, maxVoltage);
			if (testResult == PASS) strcat_s(strAddLog, MAXSTRING, "    PASS\r\n");
			else strcat_s(strAddLog, MAXSTRING, "    FAIL\r\n");
			DisplayLog(strAddLog);	
		}
		// LAST TEST IN SEQUENCE - CHECK INHIBIT RELAY INPUT TO DC950.
		// WITH POT CONTROL INPUT AT 5 VOLTS AND INHIBIT RELAY ON,
		// LAMP SHOULD BE OFF:
		else {
			if (!SetInhibitRelay(TRUE)) return SYSTEM_ERROR;
			msDelay(500);		

			expectedVoltage = (float) 0.0;
			if (!ReadVoltage(LAMP, &measuredVoltage)) return SYSTEM_ERROR;			
			errorVoltage = getAbs(measuredVoltage - expectedVoltage);
			if (errorVoltage > AllowableLampVoltageError)
				testResult = FAIL;

			maxVoltage = expectedVoltage + AllowableLampVoltageError;
			sprintf_s(strAddLog, MAXSTRING, "Inhibit ON, lamp OFF:      %.2f volts  <   MAX %.2f:", measuredVoltage, maxVoltage);
			if (testResult == PASS) strcat_s(strAddLog, MAXSTRING, "    PASS\r\n");
			else strcat_s(strAddLog, MAXSTRING, "    FAIL\r\n");
			DisplayLog(strAddLog);
		}


		char strMeasuredVoltage[MAXSTRING];
		sprintf_s(strMeasuredVoltage, MAXSTRING, "%.2f", measuredVoltage);
		// Record test results in spreadsheet:
		storeTestResult(subStepNumber + VREF_TEST, testResult, strMeasuredVoltage);

		// Return PASS, FAIL, or NOT DONE YET:
		if (testResult == FAIL) return (FAIL);

		subStepNumber++;
		if (subStepNumber >= NUM_REMOTE_TEST_VALUES) {
			if (!InitializeInterfaceBoard()) return SYSTEM_ERROR;
			return (PASS);
		}
		else return NOT_DONE_YET;
	}



	// CHECK TTL FAULT_TEST SIGNAL - OPEN AND CLOSE ACTUATOR A FEW TIMES 
	// AND MAKE SURE DC950 TTL OUTPUT REMAINS HIGH:
	int TestApp::TTL_inputTest(){	
		int testResult = PASS;

		if (!SetInterfaceBoardActuatorOutput(FILTER_CLOSED, &testResult)) return (SYSTEM_ERROR);
		if (testResult == FAIL) return FAIL;
		msDelay(300);
		if (!SetInterfaceBoardActuatorOutput(FILTER_OPEN, &testResult)) return (SYSTEM_ERROR);
		if (testResult == FAIL) return FAIL;
		msDelay(300);
		if (!SetInterfaceBoardActuatorOutput(FILTER_CLOSED, &testResult)) return (SYSTEM_ERROR);
		if (testResult == FAIL) return FAIL;
		msDelay(300);
		if (!SetInterfaceBoardActuatorOutput(FILTER_OPEN, &testResult)) return (SYSTEM_ERROR);
		if (testResult == FAIL) return FAIL;
		msDelay(300);
		if (!SetInterfaceBoardActuatorOutput(FILTER_CLOSED, &testResult)) return (SYSTEM_ERROR);		
		if (testResult == FAIL) return FAIL;
		return testResult;
	}	
	  

	// EXECUTE(): THS ROUTINE IS THE MAIN FUNCTION WHICH EXECUTES ALL TESTS.
	// It is called by TestHandler() in DC950TestDialog. The global variable 
	// "stepNumber" determines which test in the sequence is executed. 
	// 
	// Returns: "testStatus" which equals NOT_DONE_YET, PASS, or FAIL
	int TestApp::Execute(int stepNumber) {
		int testStatus = PASS; 			
		float fltVoltage = (float) 0.0;	
		static char chrSerialNumber[MAXSTRING] = "";
		static char chrPartNumber[MAXSTRING] = "";
		char lstPartNumber[MAXSTRING] = "";
		char strTest[MAXSTRING];
		float minVoltage, maxVoltage;
		int i = 0;
		static BOOL startUpFlag = TRUE;
		
		switch (stepNumber) {
		case 0:			// 0 START UP: used once after power up, when user clicks ENTER first time:		

			if (!InitializeSystem(true)){
				closeAllSerialPorts();			
				testStatus = SYSTEM_ERROR;		
			}
			else {			
				resetDisplays(TRUE);			
				testStatus = PASS;
				ptrDialog->m_static_ButtonHalt.EnableWindow(TRUE);
			}
			break;

		case 1:			// 1 SCAN BARCODE: This is the first step in the test sequence for each unit
						// If no barcode is scanned and user clicks ENTER,
						// then INVALID warning is displayed:			
			msDelay(100);	
			if (!InitializeInterfaceBoard()) return (SYSTEM_ERROR);
			if (!SetPowerSupplyVoltage(120, 60)) return (SYSTEM_ERROR);
			ptrDialog->m_BarcodeEditBox.GetWindowText((LPTSTR)chrSerialNumber, MAXSTRING);
			ClearLog();
			
			if (0 == strcmp(chrSerialNumber, "")) {
				writeInstructionLine (1, MAXLINE, "BARCODE SCAN INVALID. Please try again");
				testStatus = NOT_DONE_YET;
				break;
			}			
			testStatus = PASS;			
			break;

		case 2:			// 2 SCAN PART NUMBER
						// If no bad partnumber is scanned then INVALID warning is displayed:
			msDelay(100);				
			if (!InitializeInterfaceBoard()) return (SYSTEM_ERROR);
			ptrDialog->m_PartNumberEditBox.GetWindowText((LPTSTR)chrPartNumber, MAXSTRING);
			ClearLog();
			runFilterActuatorTest = FALSE;
			testStatus = PASS;
			if (0 == strcmp(chrPartNumber, "")) {
				testStatus = FAIL;
				break;
			}
			else {				
				if (strlen(chrPartNumber) == VALID_PART_NUMBER_LENGTH)
				{
					i = 0;
					do {
						strcpy_s(lstPartNumber, MAXSTRING, arrPartNumbers[i].String);
						// If DC950 partnumber is for filter actuator unit, display text in log box:
						if (memcmp(chrPartNumber, lstPartNumber, VALID_PART_NUMBER_LENGTH) == 0)
						{															 
							DisplayLog("This unit has a filter actuator\r\n");
							DisplayLog("Test sequence will include spectrometer scan.\r\n");
							runFilterActuatorTest = TRUE;
							break;
						}
						i++;
					} while (i < totalPartNumbers && i < MAX_PART_NUMBERS);

					if (i == totalPartNumbers)
					{											
						DisplayLog("This unit is standard model\r\n");
						DisplayLog("Spectrometer scan will NOT run.\r\n");							
					}
				}
				else {
					writeInstructionLine(1, MAXLINE, "PART NUMBER SCAN INVALID. Please try again");
					writeInstructionLine(2, MAXLINE, "Part Number must have 12 digits");
					testStatus = NOT_DONE_YET;
				}
			}
			// If both barcode and part number are valid, then both are recorded in spreadsheet
			// along with the test date:
			storeSerialNumberAndPartNumber(chrSerialNumber, chrPartNumber); // TODO
			if (runFilterActuatorTest == FALSE){
				ptrDialog->m_EditBox_Test7.EnableWindow(FALSE);
				ptrDialog->m_EditBox_Test8.EnableWindow(FALSE);
				ptrDialog->m_EditBox_Test9.EnableWindow(FALSE);
			}
			else 
			{
				ptrDialog->m_EditBox_Test7.EnableWindow(TRUE);
				ptrDialog->m_EditBox_Test8.EnableWindow(TRUE);	
				ptrDialog->m_EditBox_Test9.EnableWindow(TRUE);
			}									
			break;

		case 3:			// 3 HI POT TEST 			
			if (!InitializeInterfaceBoard()) return (SYSTEM_ERROR);
			testStatus = HiPotStatus;
			DisplayTestEditBox(HI_POT_EDIT, testStatus);
			if (testStatus == PASS) DisplayLog("Hi Pot test:    PASS\r\n");
			else DisplayLog("Hi Pot test:    FAIL\r\n");
			storeTestResult(HI_POT, testStatus, NULL);			
			break;

		case 4:			// 4 GROUND BOND TEST	
			if (!InitializeInterfaceBoard()) return (SYSTEM_ERROR);
			testStatus = GroundBondStatus;
			DisplayTestEditBox(GROUND_BOND_EDIT, testStatus);
			if (testStatus == PASS) DisplayLog("Ground Bond test:    PASS\r\n");
			else DisplayLog("Ground Bond test:    FAIL\r\n");
			storeTestResult(GROUND_BOND, testStatus, NULL);
			break;

		case 5:			// 5 POT TEST LOW = LAMP OFF				
			if (!ReadVoltage(LAMP, &fltVoltage)) return(SYSTEM_ERROR);
			
			if (fltVoltage  > AllowableLampVoltageError) testStatus = FAIL;
			else testStatus = PASS;
			DisplayTestEditBox(POT_LOW_EDIT, testStatus);

			maxVoltage = AllowableLampVoltageError;		
			sprintf_s(strTest, MAXSTRING, "Pot turned OFF: MIN 0  <  %.2f volts  <   MAX %.2f:", fltVoltage, maxVoltage);
			if (testStatus == PASS) strcat_s (strTest, MAXSTRING, "    PASS\r\n");
			else strcat_s (strTest, MAXSTRING, "    FAIL\r\n");
			DisplayLog(strTest);			
			storeTestResult(LAMP_OFF, testStatus, NULL);
			break;

		case 6:			// 6 POT TEST HIGH = LAMP ON						
			if (!ReadVoltage(LAMP, &fltVoltage)) return(SYSTEM_ERROR);
#ifdef NOSERIAL
			fltVoltage = (float) 21.0;
#endif
			if (getAbs(fltVoltage - (float) 21.0)  > AllowableLampVoltageError) testStatus = FAIL;
			else testStatus = PASS;
			DisplayTestEditBox(POT_HIGH_EDIT, testStatus);

			minVoltage = (float) 21.0 - AllowableLampVoltageError;
			maxVoltage = (float) 21.0 + AllowableLampVoltageError;			
			sprintf_s(strTest, MAXSTRING, "Pot at FULL: MIN %.2f  <  %.2f volts  <  MAX %.2f:", minVoltage, fltVoltage, maxVoltage);
			if (testStatus == PASS) strcat_s (strTest, MAXSTRING, "    PASS\r\n");
			else strcat_s (strTest, MAXSTRING, "    FAIL\r\n");
			DisplayLog(strTest);
			storeTestResult(LAMP_ON, testStatus, NULL);
			break;

		case 7:			// 7 REMOTE TEST SETUP	
			if (!InitializeInterfaceBoard()) return (SYSTEM_ERROR);
			break;

		case 8:			// 8 REMOTE POT TEST
			testStatus = RunRemoteTests();	
			if (testStatus == SYSTEM_ERROR) return (SYSTEM_ERROR);
			DisplayTestEditBox(REMOTE_EDIT, testStatus);									
			break;

		case 9:			// 9 AC POWER SWEEP
			testStatus = RunPowerSupplyTest();		
			if (testStatus == SYSTEM_ERROR) return (SYSTEM_ERROR);
			DisplayTestEditBox(AC_SWEEP_EDIT, testStatus);	
			break;

		case 10:		// 10 DARK SCAN SETUP	
			if (!SetPowerSupplyVoltage(120, 60)) return (SYSTEM_ERROR);			
			break;

		case 11:    // 11 DARK SCAN RUN						
			testStatus = RunSpectrometerScan(DARK_SCAN);
			break;

		case 12:	// 12 DC950 SCAN SETUP					
			break;		

		case 13:	// 13 ACTUATOR TEST			
			testStatus = TTL_inputTest();
			if (testStatus == SYSTEM_ERROR) break;
			DisplayTestEditBox(ACTUATOR_EDIT, testStatus);		
			 
			if (testStatus == PASS) 
			{
				DisplayLog("TTL check:    PASS\r\n");
				storeTestResult(FAULT_TEST, PASS, NULL);
			}
			else 
			{
				DisplayLog("TTL check:    FAIL\r\n");
				storeTestResult(FAULT_TEST, FAIL, "ACTUATOR ERROR");
			}			
			DisplayTestEditBox(ACTUATOR_EDIT, testStatus);			
			break;	

		case 14:			// 14 SPECTROMETER SCAN - FILTER CLOSED
			testStatus = RunSpectrometerScan(FILTER_CLOSED);
			if (testStatus == SYSTEM_ERROR) break;
			DisplayTestEditBox(FILTER_ON_EDIT, testStatus);
			break;

		case 15:			// 15 SPECTROMETER SCAN - FILTER OPEN
			testStatus = RunSpectrometerScan(FILTER_OPEN);
			if (testStatus == SYSTEM_ERROR) break;
			DisplayTestEditBox(FILTER_OFF_EDIT, testStatus);
			break;
			
		case 16:		// 16 FINAL_ASSEMBLY - TEST COMPLETE											
			testStatus = FinalAssemblyStatus;
			storeTestResult(FINAL_RESULT, testStatus, NULL);
			if (!SaveTestDataToSpreadsheet()) 
			{
				testStatus = SYSTEM_ERROR;
				SetPowerSupplyVoltage(0, 60);
				break;
			}
			SetPowerSupplyVoltage(0, 60);
			DisplayTestEditBox(FINAL_ASSEMBLY_EDIT, testStatus);
			InitializeInterfaceBoard();
			if (testStatus == PASS) DisplayLog("UNIT PASSED\r\n");
			else DisplayLog("UNIT FAILED\r\n");
			break;

		case 17:		// 17 FINAL PASS - UNIT PASSED ALL TESTS
			resetTestData();
			resetDisplays(TRUE);
			testStatus = PASS;
			break;

		case 18:		// 18 FINAL_FAIL - UNIT FAILED			
			testStatus = FAIL;
			if (FinalAssemblyStatus == NOT_DONE_YET)
			{												
				storeTestResult(FINAL_RESULT, testStatus, NULL);
				if (!SaveTestDataToSpreadsheet()) 
				{
					testStatus = SYSTEM_ERROR;
					SetPowerSupplyVoltage(0, 60);
					break;
				}				
				DisplayLog("UNIT FAILED\r\n");
				SetPowerSupplyVoltage(0, 60);
				InitializeInterfaceBoard();
			}						
			resetTestData();
			resetDisplays(TRUE);												
			break;

		case 19:		// 18 RETRY_TEST?: RETRY OR QUIT?
			break;

		case 20:			// 19 SYSTEM_FAILED		
			if (!SaveTestDataToSpreadsheet()) 
			{
				testStatus = SYSTEM_ERROR;
				SetPowerSupplyVoltage(0, 60);
				break;
			}	
			resetTestData();
			resetDisplays(FALSE);
			testStatus = FAIL;
			break;
		default:
			stepNumber = 1; // This state should never occur
			break;
		}
		return testStatus;
	}

	BOOL TestApp::SaveTestDataToSpreadsheet()
	{
		if (!SaveSpreadsheet())
		{
			if (!InitializeSpreadsheets()) 
			{
				CreateSpreadsheetErrorText(); 
				return FALSE;
			}
			if (!SaveSpreadsheet()) 
			{
				CreateSpreadsheetErrorText(); 
				return FALSE;
			}
		}
		if (!CopyFile(CurrentDataFilename, ExcelBackupFilename, FALSE))
		{
			CreateSpreadsheetErrorText(); 
			return FALSE;
		}
		
		if (!CopyFile(CurrentDataFilename, ptrExcelFilename,  FALSE)) 
		{
			if (!InitializeSpreadsheets()) 
			{
				CreateSpreadsheetErrorText(); 
				return FALSE;
			}		
		}
		return TRUE;
	}

	void TestApp::CreateSpreadsheetErrorText()
	{
		writeInstructionLine(1, MAXLINE, "SPREADSHEET ERROR");
		writeInstructionLine(2, MAXLINE, "Cannot record test data in spreadsheets.");
		writeInstructionLine(3, MAXLINE, "Do not attempt to continue testing.");
		writeInstructionLine(4, MAXLINE, "Notify System Administrator.");
		writeInstructionLine(5, MAXLINE, "");
	}

	// SPECTROMETER TEST SEQUENCE - executes all necessary steps
	// required for spectrometer scan. The input variable
	// "scanType" determines whether scan is DARK, FILTER ON or OFF.
	// 
	// Returns: PASS, FAIL, or NOT DONE YET
	// Also global variables are set when scan and calculations are complete:
	// centerWavelength, amplitude, irradiance, FWHM, colorTemperature
	int TestApp::RunSpectrometerScan(int scanType)
	{
#ifdef NOSERIAL
		return PASS;
#endif
		char strText[MAXSTRING];
		int testResult = NOT_DONE_YET;
		int tempResult = PASS;
		static int numCompletedScans = 0;
		static int elapsedScanTime = 0;

#define MAXSECONDS 120

		// If this is the first step in a scan sequence, do the following preliminary steps:
		if (subStepNumber == 0) 
		{
			int intPWM, actuatorStatus = PASS;

			// Set the remote control voltage to 
			// zero for DARK scan, MAX for FILTER ON scan
			// so DC950 is at full brightness,
			// or half brightness if FILTER OFF:
			if (scanType == FILTER_CLOSED) intPWM = MAX_PWM;
			else if (scanType == DARK_SCAN) intPWM = 0;
			else intPWM = MAX_PWM / 2;
			if (!SetInterfaceBoardPWM(intPWM)) return SYSTEM_ERROR;	
			msDelay(100);

			// Set filter actuator to ON or OFF depending on scan type:
			if (!SetInterfaceBoardActuatorOutput(scanType, &actuatorStatus)) return SYSTEM_ERROR;
			else if (actuatorStatus == FAIL) 
			{
				writeInstructionLine(3, MAXLINE, "");
				writeInstructionLine(4, MAXLINE, "Actuator didn't function properly.");
				writeInstructionLine(5, MAXLINE, "Try running scan sequence again");
				return FAIL;
			}
			msDelay(100);
			numCompletedScans = 0;
			// To initiate scan, send Activate and Start Measurement commansd to spectrometer:
			if (!ActivateSpectrometer())
			{
				displaySpectrometerCOMerror();
				return SYSTEM_ERROR;
			}
			if (!startMeasurement()) 
			{
				displaySpectrometerCOMerror();
				return SYSTEM_ERROR;
			}
		}
		// if scan has gone on for too long, then a SYSTEM ERROR has occured,
		// so send Stop Measurment command to spectrometer and quit test here:
		else if (subStepNumber > MAXSECONDS) {
			if (!stopMeasurement()) return SYSTEM_ERROR;
			writeInstructionLine(3, MAXLINE, "Spectrometer malfunction. Scan exceeded maximum time.");
			writeInstructionLine(4, MAXLINE, "Check spectrometer USB cable.");
			writeInstructionLine(5, MAXLINE, "If problem occurs again, contact system administrator.");
			DisplayLog("ERROR: Scan timeout\r\n");
			return SYSTEM_ERROR;
		}
		// Otherwise if a scan is in progress, keep polling spectrometer
		// once each program loop until it sends Data Ready response
		else if (isSpectrometerDataReady()) {	
			elapsedScanTime = 0;
			numCompletedScans++;
			// If desired number of scans have occured, send Stop Measurement
			// command to spectrometer. DARK scan is done only once unit:
			if (numCompletedScans >= m_NumberOfScans || scanType == DARK_SCAN) {
				if (!stopMeasurement()) 
				{
					displaySpectrometerCOMerror();
					return SYSTEM_ERROR;
				}				
				// Assume success, change to FAIL if any of the limits are exceeded	in the calculations below
				testResult = PASS; 
			}
			else {
				sprintf_s(strText, MAXSTRING, "Pre scan #%d done\r\n", numCompletedScans);
				writeInstructionLine(2, MAXSTRING, strText);
			}
			
			// Reference scan - probe inserted in black tube, nothing to calculate,
			// but scan data is stored:
			if (scanType == DARK_SCAN) {				
#ifdef NOSPECTROMETER
				LoadScanData("\\DarkScan.txt", Lambda, darkData);
#else
				// For DARK scan, data is stored in darkData[ ] array:
				// If spectrometer doesn't respond, halt test for SYSTEM ERROR:
				if (!getScanData(darkData))
				{
					displaySpectrometerCOMerror();
					return SYSTEM_ERROR;
				}
#endif
				DisplayLog("DARK SCAN DONE\r\n");
			}
			// Scan with actuator closed and filter is on, lamp turned up full:
			else if (scanType == FILTER_CLOSED) {
#ifdef NOSPECTROMETER
				LoadScanData("\\ClosedScan.txt", Lambda, closedFilterData);
				copyToIrradianceSpectrum (closedFilterData);
#else
				// For FILTE CLOSED scan, data is stored in closedFilterData[ ] array:
				// If spectrometer doesn't respond, halt test for SYSTEM ERROR.				
				if (!getScanData(closedFilterData))
				{
					displaySpectrometerCOMerror();
					return SYSTEM_ERROR;
				}
#endif
				DisplayLog("CLOSED FILTER SCAN DONE\r\n");
			}
			// Actuator open, light is unfiltered
			else  
			{
#ifdef NOSPECTROMETER
				LoadScanData("\\OpenScan.txt", Lambda, openFilterData);
				copyToIrradianceSpectrum (openFilterData);
#else
				// For FILTER OFF scan, data is stored in openFilterData[ ] array:
				if (!getScanData(openFilterData)) 
				{
					displaySpectrometerCOMerror();
					return SYSTEM_ERROR;
				}
#endif				 
				DisplayLog("OPEN FILTER SCAN DONE\r\n");
			}					

			// NOW that scan data has been acquired, do calculations for FILTER ON or OFF.
			// If all calcluated data is within test limits, results is PASS.
			// If one or more limit is exceeded, result is FAIL.
			// Test data and result is copied in DisplayLog box onscreen 
			// and also recorded in spreadsheet below:

			// SPECTROMETER TEST FOR CLOSED FILTER:
			if (scanType == FILTER_CLOSED) 
			{
				convertScopeDataToIrradiance(closedFilterData, irradianceIntensity, m_enableLinearCorrection, m_enableStrayLightCorrection);
				copyToIrradianceSpectrum (irradianceIntensity);				
				getIrradiance ((float)175.0, (float)1099.0, &closedIrradianceIntegral);
				
				getCenterWavelengthAndFWHM ((float)175.0, (float)1099.0, (int)100, &centerWavelength, &wavelengthAmplitude, &FWHM);
				// CENTER WAVELNGTH CLOSED:
				sprintf_s(strText, MAXSTRING, "%.2f", centerWavelength);
				if (centerWavelength < MinClosedFilterWavelength || centerWavelength > MaxClosedFilterWavelength ) {
					storeTestResult(WAVELENGTH_CLOSED, FAIL, strText);
					tempResult = testResult = FAIL;
				}
				else {
					tempResult = PASS;
					storeTestResult(WAVELENGTH_CLOSED, PASS, strText);
				}

				sprintf_s(strText, MAXSTRING, "Wavelength closed:     MIN %.2f < %.2f < MAX %.2f:", MinClosedFilterWavelength, centerWavelength, MaxClosedFilterWavelength);
				if (tempResult == PASS) strcat_s(strText, MAXSTRING, "    PASS\r\n");
				else strcat_s(strText, MAXSTRING, "    FAIL\r\n");	
				DisplayLog(strText);

				// IRRADIANCE CLOSED:
				sprintf_s(strText, MAXSTRING, "%.2f", closedIrradianceIntegral);
				if (closedIrradianceIntegral > MaxClosedFilterIrradiance || closedIrradianceIntegral < MinClosedFilterIrradiance) {
					tempResult = testResult = FAIL;
					storeTestResult(IRRADIANCE_CLOSED, FAIL, strText);
				}
				else {
					tempResult = PASS;
					storeTestResult(IRRADIANCE_CLOSED, PASS, strText);
				}

				sprintf_s(strText, MAXSTRING, "Irradiance filter closed:      MIN %.2f < %.2f < MAX %.2f:", MinClosedFilterIrradiance, closedIrradianceIntegral, MaxClosedFilterIrradiance);
				if (tempResult == PASS) strcat_s(strText, MAXSTRING, "    PASS\r\n");
				else strcat_s(strText, MAXSTRING, "    FAIL\r\n");	
				DisplayLog(strText);

				// MIN_AMPLITUDE_CLOSED:
				sprintf_s(strText, MAXSTRING, "%.2f", wavelengthAmplitude);
				if (wavelengthAmplitude < MinClosedFilterPeakAmplitude || wavelengthAmplitude > MaxClosedFilterPeakAmplitude) {
					tempResult = testResult = FAIL;
					storeTestResult(MIN_AMPLITUDE_CLOSED, FAIL, strText);
				}
				else {
					tempResult = PASS;
					storeTestResult(MIN_AMPLITUDE_CLOSED, PASS, strText);
				}

				sprintf_s(strText, MAXSTRING, "Min amplitude filter closed:      MIN %.2f < %.2f < MAX %.2f:", MinClosedFilterPeakAmplitude, wavelengthAmplitude, MaxClosedFilterPeakAmplitude);
				if (tempResult == PASS) strcat_s(strText, MAXSTRING, "    PASS\r\n");
				else strcat_s(strText, MAXSTRING, "    FAIL\r\n");	
				DisplayLog(strText);

				// FWHM_CLOSED:
				sprintf_s(strText, MAXSTRING, "%.2f", FWHM);
				if (FWHM < MinClosedFilterFWHM || FWHM > MaxClosedFilterFWHM) {
					tempResult = testResult = FAIL;
					storeTestResult(FWHM_CLOSED, FAIL, strText);
				}
				else {
					tempResult = PASS;
					storeTestResult(FWHM_CLOSED, PASS, strText);
				}

				sprintf_s(strText, MAXSTRING, "FWHM filter closed:      MIN %.2f < %.2f < MAX %.2f:", MinClosedFilterFWHM, FWHM, MaxClosedFilterFWHM);
				if (tempResult == PASS) strcat_s(strText, MAXSTRING, "    PASS\r\n");
				else strcat_s(strText, MAXSTRING, "    FAIL\r\n");	
				DisplayLog(strText);
			}
			// SPECTROMETER TEST FOR FILTER OPEN:
			else if (scanType == FILTER_OPEN) 
			{
				convertScopeDataToIrradiance(openFilterData, irradianceIntensity, m_enableLinearCorrection, m_enableStrayLightCorrection);
				copyToIrradianceSpectrum (irradianceIntensity);
				getColorTemp(&openColorTemp);
				getIrradiance ((float)175.0, (float)1099.0, &openIrradianceIntegral);


				// IRRADIANCE OPEN:
				sprintf_s(strText, MAXSTRING, "%.2f", openIrradianceIntegral);
				if (openIrradianceIntegral < MinOpenIrradiance || openIrradianceIntegral > MaxOpenIrradiance) {
					tempResult = testResult = FAIL;
					storeTestResult(IRRADIANCE_OPEN, FAIL, strText);
				}
				else {
					tempResult = PASS;
					storeTestResult(IRRADIANCE_OPEN, PASS, strText);
				}

				sprintf_s(strText, MAXSTRING, "Irradiance filter open:      MIN %.2f < %.2f < MAX %.2f:", MinOpenIrradiance, openIrradianceIntegral, MaxOpenIrradiance);
				if (tempResult == PASS) strcat_s(strText, MAXSTRING, "    PASS\r\n");
				else strcat_s(strText, MAXSTRING, "    FAIL\r\n");	
				DisplayLog(strText);

				// COLOR TEMPERATURE OPEN:
				sprintf_s(strText, MAXSTRING, "%.2f", openColorTemp);
				if (openColorTemp < MinOpenColorTemp || openColorTemp > MaxOpenColorTemp ) {
					tempResult = testResult = FAIL;
					storeTestResult(COLOR_TEMP_OPEN, FAIL, strText);
				}
				else {
					tempResult = PASS;
					storeTestResult(COLOR_TEMP_OPEN, PASS, strText);
				}

				sprintf_s(strText, MAXSTRING, "Color Temp open:      MIN %.2f < %.2f < MAX %.2f:", MinOpenColorTemp, openColorTemp, MaxOpenColorTemp);
				if (tempResult == PASS) strcat_s(strText, MAXSTRING, "    PASS\r\n");
				else strcat_s(strText, MAXSTRING, "    FAIL\r\n");	
				DisplayLog(strText);				
			}
		}
		else {
			elapsedScanTime = elapsedScanTime + 3;
			sprintf_s(strText, MAXSTRING, "Scan #%d time: %d seconds", numCompletedScans+1, elapsedScanTime);
			writeInstructionLine(2, MAXSTRING, strText);
			testResult = NOT_DONE_YET;
		}
		if (numCompletedScans < m_NumberOfScans && scanType != DARK_SCAN) testResult = NOT_DONE_YET;
		subStepNumber++;
		return testResult;
	}

void TestApp::resetSubStepNumber(){
	subStepNumber = 0;
}


/*
BOOL TestApp::writeSpectrometerDataToFile(const char *filename, double *ptrLambda, float *ptrScopeData) 
{
	ofstream myfile;
	uint16 i;
	char strSpectrometerData[MAXSTRING];
	double lambdaValue;
	float scopeDataValue;
		
	if (m_StopPixel >= PIXEL_ARRAY_SIZE) return FALSE;

	myfile.open(filename);
	for (i = 0; i < m_StopPixel; i++) {		
		scopeDataValue = (float) ptrScopeData[i];
		lambdaValue = (float) ptrLambda[i];
		// sprintf_s (strSpectrometerData, MAXSTRING, "%f\t%f\n", lambdaValue, scopeDataValue);
		sprintf_s (strSpectrometerData, MAXSTRING, "%f, %f\n", lambdaValue, scopeDataValue);
		myfile << strSpectrometerData;
	}
	myfile.close();
	return TRUE;
}
*/

// For displaying system errors onscreen:
void TestApp::displaySpectrometerCOMerror()
{
	writeInstructionLine(3, MAXLINE, "Spectrometer not communicating");
	writeInstructionLine(4, MAXLINE, "Check spectrometer USB cable.");
	writeInstructionLine(5, MAXLINE, "If problem occurs again, contact system administrator.");
	DisplayLog("ERROR: Spectrometer not communicating\r\n");
}

void TestApp::displayInterfaceCOMerrorr()
{
	writeInstructionLine(3, MAXLINE, "INTERFACE BOX isn't communicating");
	writeInstructionLine(4, MAXLINE, "Check serial connections to interface box.");
	writeInstructionLine(5, MAXLINE, "Make sure power is on");
	DisplayLog("ERROR: Spectrometer not communicating\r\n");
}

/*
void TestApp::insertBogusData()
{
	int i;
	int testResult = PASS;
	char strMeasuredVoltage[MAXSTRING];
	float testVoltage = (float) 1.12;

	for (i = VREF_TEST; i <= COLOR_TEMP_OPEN; i++)
	{
		if (i != FAULT_TEST)
		{
			sprintf_s(strMeasuredVoltage, MAXSTRING, "%.2f", testVoltage);	
			storeTestResult(i, testResult, strMeasuredVoltage);
			testVoltage = testVoltage * (float) 2.0;
			if (testResult == PASS) testResult = FAIL;
			else testResult = PASS;
		}
	}
	storeTestResult(0, PASS, "One");  
	storeTestResult(1, PASS, "Two");
	storeTestResult(2, PASS, "Three");
}
*/
