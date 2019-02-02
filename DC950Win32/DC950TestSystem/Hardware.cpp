// COMMENTED
/* Hardware.cpp - mid level routines for operating Interface box, multimeter, and power supply.
 * All three devices use serial ports for communication, so SerialCom file is required
 * for low level routines.
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
//#include "BasicExcel.hpp"
//#include "ExcelFormat.h"
#include <time.h>
#include "ReadOnlyEdit.h"
#include "PasswordEntry_Dlg.h"	

HANDLE gDoneEvent;
HANDLE hTimer = NULL;
HANDLE hTimerQueue = NULL;

extern HANDLE handleInterfaceBoard, handleHPmultiMeter, handleACpowerSupply;

	// Closes all three serial ports
	void TestApp::closeAllSerialPorts(){
		closeSerialPort(handleInterfaceBoard); 
		closeSerialPort(handleHPmultiMeter);
		closeSerialPort(handleACpowerSupply);
	}

	// Returns absolute value of a float variable
	float TestApp::getAbs(float floatValue) {
		float absoluteValue;
		if (floatValue < (float) 0.0) absoluteValue = (float) 0.0 - floatValue;
		else absoluteValue = floatValue;
		return absoluteValue;
	}

	// These two routines implement a millisecond timer used for short delays
	VOID CALLBACK TimerRoutine(PVOID lpParam, BOOLEAN TimerOrWaitFired)
	{
		SetEvent(gDoneEvent);
	}

	void TestApp::msDelay(int milliseconds) {		
		gDoneEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
		hTimerQueue = CreateTimerQueue();
		int arg = 123;
		CreateTimerQueueTimer(&hTimer, hTimerQueue, (WAITORTIMERCALLBACK)TimerRoutine, &arg, milliseconds, 0, 0);
		WaitForSingleObject(gDoneEvent, INFINITE);
		CloseHandle(gDoneEvent);
	}

			
	// This routine commands the multimeter to read the selected voltage source.
	// Inputs:  multiplexerSelect - (should really be called relaySelect) commands
	//				the Interface box to set the relays to route the selected voltage source
	//				to the HP multimeter. 
	// Outputs:	ptrMeasuredVoltage - references the floating point variable
	//				for the multimeter output voltage.
	// Returns: TRUE if communication with the Interface box and multimeter works,
	//				or FALSE if there is a SYSTEM ERROR with either serial port.
	BOOL TestApp::ReadVoltage(int multiplexerSelect, float *ptrMeasuredVoltage)
	{
#ifdef NOSERIAL
		return TRUE;
#endif
		char strResponse[BUFFERSIZE] = "";
		char strReadVoltage[BUFFERSIZE] = ":MEAS?\r\n";		
		float floatVoltage;

		// Set relays on interface box for desired voltage source:
		// Return FALSE if Interface box isn't responding to serial commands:
		if (!SetInterfaceBoardMultiplexer(multiplexerSelect)) return FALSE;
		msDelay(500);

		// Send READ VOLTAGE command to multimeter and get response:
		// Return FALSE if multimeter isn't responding to serial commands:
		if (!sendReceiveSerial(HP_METER, strReadVoltage, strResponse, TRUE)) return FALSE;		
		// If communication with multimeter was successful, convert response string to float value:
		else {
			floatVoltage = (float)atof(strResponse);
			*ptrMeasuredVoltage = getAbs(floatVoltage);
		}
		return TRUE;
	}
		
	// Sends a command to the Interface box to set the relays for the mutlimeter input.
	// Input:  multiplexerSelect - (should really be called relaySelect) 
	//				indicates desired multimeter voltage source:
	//				LAMP, VREF, or remote CONTROL_VOLTAGE
	// Returns: TRUE if communication with the Interface box works,
	//				or FALSE if there is a SYSTEM ERROR.
	BOOL TestApp::SetInterfaceBoardMultiplexer(int multiplexerSelect)
	{
#ifdef NOSERIAL
		return TRUE;
#endif

		char strCommand[BUFFERSIZE];
		char strResponse[BUFFERSIZE];

		// Create command string to send to the Interface box for desired voltage source:
		switch (multiplexerSelect) {
		default:	
		case LAMP:
			strcpy_s(strCommand, BUFFERSIZE, "$LAMP");
			break;
		case VREF:
			strcpy_s(strCommand, BUFFERSIZE, "$VREF");
			break;
		case CONTROL_VOLTAGE:		
			strcpy_s(strCommand, BUFFERSIZE, "$CTRL");
			break;
		}

		// Send command to Interface box. If there is a serial communication error, return FALSE:
		if (!sendReceiveSerial(INTERFACE_BOARD, strCommand, strResponse, TRUE)) 
			return FALSE;		
		// If the Interface box doesn't acknowledge the command, return FALSE:
		else if (!strstr(strResponse, "OK")) 
			return FALSE;		
		// If Interface box responds with "OK" then command was accepted:
		return TRUE;
	}

	// This command sends a command to the INterface box to set the 
	// remote control voltage output to the DC950.
	// Input: PWMvalue - is an integer value between 0 and MAX PWM
	// Returns: TRUE if Interface box acknowledges command.
	BOOL TestApp::SetInterfaceBoardPWM(int PWMvalue)
	{
#ifdef NOSERIAL
		return TRUE;
#endif
		char strCommand[BUFFERSIZE] = "$PWM>";
		char strValue[BUFFERSIZE], strResponse[BUFFERSIZE];
		sprintf_s(strValue, BUFFERSIZE, "%d", PWMvalue);
		strcat_s(strCommand, BUFFERSIZE, strValue);

		// Send PWM command to Interface box. If there is a serial communication error, return FALSE:
		if (!sendReceiveSerial(INTERFACE_BOARD, strCommand, strResponse, TRUE)) 
			return FALSE;		
		// If the Interface box doesn't acknowledge the command, return FALSE:
		else if (!strstr(strResponse, "OK")) 
			return FALSE;
		// If Interface box responds with "OK" then command was accepted:
		else return TRUE;
	}
	
	// This routine is used when testing the remote inhibit feature on the DC950
	// Input: flagON - is a flag for the inhibit relay state - TRUE turns relay ON.
	// Returns 
	BOOL TestApp::SetInhibitRelay(BOOL flagON)
	{
#ifdef NOSERIAL
		return TRUE;
#endif
		char strCommand[BUFFERSIZE];
		char strResponse[BUFFERSIZE];

		// Select command for relay ON or OFF:
		if (flagON) strcpy_s(strCommand, BUFFERSIZE, "$INHIBIT>ON");
		else strcpy_s(strCommand, BUFFERSIZE, "$INHIBIT>OFF");

		// Send command to Interface box to set inhibit relay. 
		// If there is a serial communication error, return FALSE:
		if (!sendReceiveSerial(INTERFACE_BOARD, strCommand, strResponse, TRUE)) 
			return FALSE;		
		// If the Interface box doesn't acknowledge the command, return FALSE:
		else if (!strstr(strResponse, "OK")) 
			return FALSE;
		// If Interface box responds with "OK" then command was accepted:
		else return TRUE;
	}
	

	// This routine sets the AC voltage and frequency of the programmable power supply
	// Inputs:	commandVoltage is an integer value from 0 - MAX_AC_VOLTAGE
	//			frequency - must be 50 or 60 fro AC frequency in Hz
	// Returns: TRUE if serial communication is successful
	BOOL TestApp::SetPowerSupplyVoltage(int commandVoltage, int frequency) {
#ifdef NOSERIAL
		return TRUE;
#endif
		char strCommand[BUFFERSIZE];
		char strResponse[BUFFERSIZE] = "";
		char strReadVoltage[BUFFERSIZE] = "MEAS:VOLT?\r\n";		

		#ifdef NOPOWERSUPPLY 
			return TRUE; 
		#endif

		if (frequency != 50 && frequency != 60) frequency = 60;

    	  // Send SET FREQUENCY command:
		sprintf_s(strCommand, "FREQ %d\r\n", frequency);

		// Return FALSE if there is a  serial port error:
		// Return FALSE if power supply serial port isn't working:
		if (!sendReceiveSerial(AC_POWER_SUPPLY, strCommand, NULL, FALSE)) return FALSE;		

		msDelay(500);
    	 // Send SET VOLTAGE command:
		sprintf_s(strCommand, "VOLT %d\r\n", commandVoltage);

		// Return FALSE if power supply serial port isn't working:
		if (!sendReceiveSerial(AC_POWER_SUPPLY, strCommand, NULL, FALSE)) return FALSE;	

		msDelay(500);
		// Now send a MEASURE command to power supply.
		// This hopefully indicates that the supply is receiving commands and responding to them.		 
		// The TRUE input tells sendReceiveSerial() to expect a response
		// If there is no repsonse then sendReceiveSerial() returns FALSE:
		strcpy_s(strCommand, BUFFERSIZE, "MEAS:VOLT?\r\n");
		
		// If programmable supply doesn't respond to serial command, return FALSE
		// indicating a serial communication error:
		if (!sendReceiveSerial(AC_POWER_SUPPLY, strCommand, NULL, TRUE)) 
			return FALSE;
			
		// Otherwise, return TRUE to indicate power supply is communicating and functioning:
		return TRUE;
	}

	// This routine sends a command to the Interface box to set the remote TTL input on the DC950.
	// This input should be set to HIGH to turn actuator filter ON, or LOW to open actuator.
	// The DC950 has a TTL output which is some sort of fault indicator.
	// A HIGH TTL output indicates actuator is working.
	//
	// Inputs: scanType - is FILTER_ON, FILTER_OFF, or DARK.
	//				For a FILTER_ON scan, the actuator is turned ON, otherwise it's OFF
	//	Output:	testResult - is set to FAIL if actuator TTL signal indicates a fault.
	//	Returns TRUE if serial communication with Interface box is working.
	BOOL TestApp::SetInterfaceBoardActuatorOutput(int scanType, int *testResult)
	{
#ifdef NOSERIAL
		return TRUE;
#endif
		char strCommand[BUFFERSIZE] = "$TTL_HIGH";
		char strInputCommand[BUFFERSIZE] = "$TTL_IN";
		char strResponse[BUFFERSIZE];

		SetInterfaceBoardPWM(MAX_PWM);
		msDelay(100);

		// 0V signal to TTL Input to close filter
		if (scanType == FILTER_CLOSED) strcpy_s(strCommand, BUFFERSIZE, "$TTL_LOW");
		else strcpy_s(strCommand, BUFFERSIZE, "$TTL_HIGH"); 

		// Send SET TTL command to interface box:
		// Return FALSE if Interface box isn't responding to serial commands:
		if (!sendReceiveSerial(INTERFACE_BOARD, strCommand, strResponse, TRUE)) 			
			return FALSE;
		
		// Now get response from interface box:
		else if (!strstr(strResponse, "OK")) 
			return FALSE;

		// Check TTL input for actuator fault detect:
		// Return FALSE if Interface box isn't responding to serial commands:
		if (!sendReceiveSerial(INTERFACE_BOARD, strInputCommand, strResponse, TRUE)) 
			return FALSE;		
		else if (!strstr(strResponse, "OK")) 
			*testResult = FAIL;
				
		return TRUE; // Interface box is repsonding to commands
	}
	
	// Sends a command to Interface box to check remote fault signal on the DC950
	// Inputs: none
	// Output: testResult set to FAIL if DC950 fault output is HIGH.
	// Returns: TRUE if serial communication with Interface box is working.
	BOOL TestApp::CheckFaultSignal(int *testResult)
	{
#ifdef NOSERIAL
		return TRUE;
#endif
		char strCommand[BUFFERSIZE] = "$FAULT_IN"; // Command for Interface box to check fault 
		char strResponse[BUFFERSIZE];

		msDelay(100);

		// Send FAULT INPUT command to interface box to check state of FAULT input.
		// Return FALSE if Interface box isn't responding to serial commands:
		if (!sendReceiveSerial(INTERFACE_BOARD, strCommand, strResponse, TRUE)) 			
			return FALSE;

		// Get response from interface box. 
		// The DC950 LAMP FAULT signal goes HIGH when there is a problem with the lamp.
		else if (strstr(strResponse, "LOW")) 
			*testResult = PASS;
		else *testResult = FAIL;

		return TRUE; // Interface box is working
	}

