// COMMENTED
/* SerialCom.cpp - low level routines for sending and receiving data on serial ports, as well as opening and closing.
 * Written in Visual C++ for DC950
 *
 */

// NOTE: INCUDES MUST BE IN THIS ORDER!!!
#include "stdafx.h"
#include "DC950Test.h"
#include "DC950TestDialog.h"
#include <string.h>
#include "TestApp.h"
#include "Definitions.h"

extern CTestDialog *ptrDialog;
extern UINT16  CRCcalculate(char *ptrPacket, BOOL addCRCtoPacket);
extern BOOL CRCcheck(char *ptrPacket);
extern HANDLE handleInterfaceBoard, handleHPmultiMeter, handleACpowerSupply;

// Opens serial port with name string specified by ptrPortName, i.e. "COM1"
// If successful, it ptrPortHandle points to the handle for the activated port
// Returns TRUE if port is properly opened and configured, otherwise FALSE
BOOL TestApp::openSerialPort (const char *ptrPortName, HANDLE *ptrPortHandle) {
			DCB serialPortConfig;
			BOOL tryAgain = FALSE;		

			if (ptrPortName == NULL) return FALSE;
			do {
					*ptrPortHandle = CreateFile((LPCSTR)ptrPortName,  // Specify port device such as "COM1"		
					GENERIC_READ | GENERIC_WRITE,       // Specify mode that open device.
					0,                                  // the devide isn't shared.
					NULL,                               // the object gets a default security.
					OPEN_EXISTING,                      // Specify which action to take on file. 
					0,                                  // default.
					NULL);                              // default.

				// Try to open port here:
				if (GetCommState(*ptrPortHandle, &serialPortConfig) == 0) {
					// DisplayMessageBox("ERROR: CANNOT OPEN SERIAL PORT", "Close this program, check USB connections, then restart program", 1);
					return FALSE;
				}
				else break;   // SUCCESS! Port opened OK, no need to try opening it again
				msDelay(100);
			} while (tryAgain);

			// Now set baud rate and other parameters:
			DCB dcb;
			dcb.BaudRate = CBR_9600;			// Fixed baud rate is 9600, 1 stop bit, no parity, 8 data bits
			dcb.StopBits = ONESTOPBIT;
			dcb.Parity = NOPARITY;
			dcb.ByteSize = DATABITS_8;

			// Assign user parameter.
			serialPortConfig.BaudRate = dcb.BaudRate;    // Specify buad rate of communicaiton.
			serialPortConfig.StopBits = dcb.StopBits;    // Specify stopbit of communication.
			serialPortConfig.Parity = dcb.Parity;        // Specify parity of communication.
			serialPortConfig.ByteSize = dcb.ByteSize;    // Specify size of communication.

			// Now set current configuration of serial communication port.																 
			if (SetCommState(*ptrPortHandle, &serialPortConfig) == 0)
			{
				AfxMessageBox((LPCTSTR)"PROGRAM ERROR: Set configuration port has problem.");
				return FALSE;
			}

			// Instance an object of COMMTIMEOUTS.
			COMMTIMEOUTS comTimeOut;
			comTimeOut.ReadIntervalTimeout = 3;
			comTimeOut.ReadTotalTimeoutMultiplier = 3;
			comTimeOut.ReadTotalTimeoutConstant = 2;
			comTimeOut.WriteTotalTimeoutMultiplier = 3;
			comTimeOut.WriteTotalTimeoutConstant = 2;

			SetCommTimeouts(*ptrPortHandle, &comTimeOut);		// set the time-out parameter into device control.												
			return TRUE;  // Return TRUE to indicate port is successfully opened and configured
	}

	// Closes serial port pointed to by input handle
	BOOL TestApp::closeSerialPort(HANDLE ptrPortHandle) {
		if (ptrPortHandle == NULL) 
			return TRUE;

		if (!CloseHandle(ptrPortHandle))
		{
			// AfxMessageBox((LPCTSTR)"Port Closing failed.");
			return FALSE;
		}
		ptrPortHandle = NULL;
		return(TRUE);
	}


	
	// Writes character string to serial port.
	// Inputs: ptrPortHandle points to desired serial port,
	// ptrPacket is the ASCII character string to send
	// 
	// Returns TRUE if successful, FALSE if unable to send string after several retries
	BOOL TestApp::WriteSerialPort (HANDLE ptrPortHandle, char *ptrPacket) {
		int length;
		int trial = 0;
		int numBytesWritten = 0;		

		// Make sure inputs are real:
		if (ptrPacket == NULL || ptrPortHandle == NULL) return (FALSE);
		
		// Make sure character string is within valid length:
		length = (int) strlen(ptrPacket);					
		if (length >= BUFFERSIZE) {			
			return (FALSE);
		}		

		// Now send string to serial port. Allow up to MAXTRIES retries,
		// then return FALSE if not succesful:
		do {
			trial++;
			if (WriteFile(ptrPortHandle, ptrPacket, length, (LPDWORD) &numBytesWritten, NULL)) break;
			msDelay(100);
		} while (trial < MAXTRIES && numBytesWritten == 0);

		// Return TRUE if serial port write was successful:
		if (trial >= MAXTRIES) return FALSE;
		else return TRUE;
	}
	
	// Reads character string from serial port. String must be terminated by 
	// carriage return '\r' or newline '\n' or both.
	// Inputs: ptrPortHandle points to serial port to read from
	// Outputs: the array referenced by ptrPacket gets filled with incoming null characters, terminated by '\0'
	//	 
	// Returns TRUE if successful, FALSE if unable no string received after several retries		
	BOOL TestApp::ReadSerialPort(HANDLE ptrPortHandle, char *ptrPacket) {
		int totalBytesRead = 0;
		DWORD numBytesRead;
		char inBytes[BUFFERSIZE];
		ptrPacket[0] = '\0';
		int trial = 0;

		// Make sure ptrPacket references an actual array:
		if (ptrPacket == NULL) {			
			return (FALSE);
		}

		trial = 0;
		// Initialize receive string to contain no characters to start with:
		ptrPacket[0] = '\0';
		// Check serial port for received characters
		// and concatenate them to ptrPacket as they come in.
		// Quit when a carriage return or newline character is received.
		// Allow up to MAXTRIES attmpts to receive string, then quit if necessary:
		do {
			trial++;		
			msDelay(100);
			if (ReadFile(ptrPortHandle, inBytes, BUFFERSIZE, &numBytesRead, NULL)) {
				if (numBytesRead > 0 && numBytesRead < BUFFERSIZE) {
					inBytes[numBytesRead] = '\0';
					strcat_s(ptrPacket, BUFFERSIZE, inBytes);
					if (strchr(inBytes, '\r')) break;	
					if (strchr(inBytes, '\n')) break;
				}
			}
			
		} while (trial < MAXTRIES);

		// return TRUE if complete string is received, otherwise FALSE:
		if (trial >= MAXTRIES) return (FALSE);		
		else return (TRUE);		
	}
	
	// This routine sends a command string to a serial device
	// and if a response string is expected, receive it.
	// Inputs: COMdevice indicates which serial port to use
	//			outPacket is the outgoing command string
	//			expectReply should be set to TRUE if a response is expected
	//	Outputs: If a response is expected, inPacket is filled with received ASCII character string
	//	Returns: TRUE if successful
	BOOL TestApp::sendReceiveSerial(int COMdevice,  char *outPacket, char *inPacket, BOOL expectReply)
	{
#ifdef NOSERIAL
		return TRUE;
#endif

		HANDLE ptrPortHandle = NULL;

		// Use handle for selected serial port:
		switch (COMdevice) {
		case HP_METER:
			ptrPortHandle = handleHPmultiMeter;
			break;
		case AC_POWER_SUPPLY:
			ptrPortHandle = handleACpowerSupply;
			break;
		case INTERFACE_BOARD:
			ptrPortHandle = handleInterfaceBoard;
		default:
			break;
		}		

		if (ptrPortHandle == NULL || outPacket == NULL) {
			return (FALSE);
		}
			
		// If communicating with the Interface box, calculate the CRC 
		// for the outgoing command string and add it to end of string:
		if (COMdevice == INTERFACE_BOARD) 
			CRCcalculate(outPacket, TRUE);

		if (outPacket == NULL) return (FALSE);
				
		// Send command string to serial port. If unsuccessful,
		// display error message in user instruction window
		if (!WriteSerialPort(ptrPortHandle, outPacket))
		{
			if (COMdevice == INTERFACE_BOARD)
			{
				writeInstructionLine(1, MAXLINE, "INTERFACE BOX COM ERROR");
				writeInstructionLine(2, MAXLINE, "Check serial cable to interface box.");
				writeInstructionLine(3, MAXLINE, "Make sure interface box is turned on.");
				writeInstructionLine(4, MAXLINE, "Then click ENTER to restart test");
				writeInstructionLine(5, MAXLINE, "");
			}
			else if (COMdevice == AC_POWER_SUPPLY)
			{
				writeInstructionLine(1, MAXLINE, "AC POWER SUPPLY COM ERROR");
				writeInstructionLine(2, MAXLINE, "Make sure AC power supply is ON");
				writeInstructionLine(3, MAXLINE, "Check RS232 cable");		
				writeInstructionLine(4, MAXLINE, "Then click ENTER to restart test");
				writeInstructionLine(5, MAXLINE, "");
			}
			else
			{
				writeInstructionLine(1, MAXLINE, "MULTIMETER COM ERROR");
				writeInstructionLine(2, MAXLINE, "Check serial cable to meter.");
				writeInstructionLine(3, MAXLINE, "Turn meter power off, then turn it back on.");
				writeInstructionLine(4, MAXLINE, "Then click ENTER to restart test");
				writeInstructionLine(5, MAXLINE, "");
			}
			return FALSE;
		}

		// If string was sent and no reply was expected, return TRUE and exit here:
		if (!expectReply) return TRUE;

		// If a reply atring is expected but no array has been provided to place it in,
		// use a temp dummy array instead. This is done 
		// with the programmable power supply, when a command is sent just to check
		// communication but the reply string doesn't need to be processed.
		if (inPacket == NULL) {
			char temp[BUFFERSIZE];
			inPacket = temp;
		}

		// Now receive reply string from serial port.
		// If unsuccessful after several retries,
		// display system error in user instruction window
		// then return FALSE:
		if (!ReadSerialPort(ptrPortHandle, inPacket)) 
		{		
			if (COMdevice == INTERFACE_BOARD)
			{
				writeInstructionLine(1, MAXLINE, "INTERFACE BOX COM READ ERROR");
				writeInstructionLine(2, MAXLINE, "Check serial cable to interface box.");
				writeInstructionLine(3, MAXLINE, "Make sure interface box is turned on.");
				writeInstructionLine(4, MAXLINE, "Then click ENTER to restart test");
				writeInstructionLine(5, MAXLINE, "");
			}
			else if (COMdevice == AC_POWER_SUPPLY)
			{
				writeInstructionLine(1, MAXLINE, "AC POWER SUPPLY COM READ ERROR");
				writeInstructionLine(2, MAXLINE, "Make sure AC power supply is ON");
				writeInstructionLine(3, MAXLINE, "Check RS232 cable");		
				writeInstructionLine(4, MAXLINE, "Then click ENTER to restart test");
				writeInstructionLine(5, MAXLINE, "");
			}
			else
			{
				writeInstructionLine(1, MAXLINE, "MULTIMETER COM READ ERROR");
				writeInstructionLine(2, MAXLINE, "Check serial cable to meter.");
				writeInstructionLine(3, MAXLINE, "Turn meter power off, then on.");
				writeInstructionLine(4, MAXLINE, "Then click ENTER to restart test");
				writeInstructionLine(5, MAXLINE, "");
			}
			return FALSE;
		}

		// If serial device is the Interface box, check the CRC on the end of incoming string.
		// If string isn't valid return TRUE:
		if (COMdevice == INTERFACE_BOARD) {
			if (!CRCcheck(inPacket)) 
			{
				writeInstructionLine(1, MAXLINE, "INTERFACE BOX CRC ERROR");
				writeInstructionLine(2, MAXLINE, "Check serial cable to interface box.");
				writeInstructionLine(3, MAXLINE, "Make sure interface box is turned on.");
				writeInstructionLine(4, MAXLINE, "Then click ENTER to restart test");
				writeInstructionLine(5, MAXLINE, "");
				return FALSE;
			}
		}		
		// Return TRUE of both write and read are OK:
		return (TRUE);		
	}


	