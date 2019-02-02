#ifndef _TEST_ROUTINES_H
#define _TEST_ROUTINES_H
#include "stdafx.h"
#include "DC950TestDialog.h"
#include "Definitions.h"
#include "avaspec.h"
#include "irrad.h"
#include <iostream>
#include <sstream>

class TestApp : public CDialog
{
public:			
	TestStep testStep[TOTAL_STEPS];	
	BOOL SaveTestDataToSpreadsheet();
	void CreateSpreadsheetErrorText();
	BOOL getCurrentRow(int *ptrCurrentRow);
	BOOL validateTempPath();
	void ShutDownExcel();
	void makeRangeString(int row, int column, char *strCell);
	void Errorf( LPCTSTR pszFormat, ... );
	BOOL InitializeSpreadsheets();	
	BOOL createSecondarySpreadsheetName(char *ptrPrimary, char *ptrSecondary);
	int LoadExcelFilenames(char *ptrPrimary, char *ptrSecondary);
	void writeInstructionLine (int lineNumber, int lineLength, char *ptrText);
	BOOL sendReceiveSerial (int COMdevice, char *outPacket, char *inPacket, BOOL useCRC);
	void msDelay (int milliseconds);	
	BOOL openSerialPort(const char *ptrPortName, HANDLE *ptrPortHandle);
	BOOL closeSerialPort(HANDLE ptrPortHandle);
	void closeAllSerialPorts();		
	BOOL ReadSerialPort(HANDLE ptrPortHandle, char *ptrPacket);
	BOOL WriteSerialPort(HANDLE ptrPortHandle, char *ptrPacket);	
	BOOL InitializeTestEditBoxes();	
	void resetTestData();	
	int  Execute(int stepNumber);		
	void enableBarcodeScan();
	void disableBarcodeScan();	
	void enablePartNumberScan();
	void disablePartNumberScan();		
	BOOL SetPowerSupplyVoltage(int commandVoltage, int frequency);
	BOOL SetInterfaceBoardMultiplexer(int multiplexerSelect);
	BOOL ReadVoltage(int multiplexerSelect, float *ptrLampVoltage);
	float getAbs(float floatValue);
	BOOL SetInterfaceBoardPWM(int PWMvalue);
	int  RunRemoteTests();
	int  RunPowerSupplyTest();
	BOOL SetInhibitRelay(BOOL flagON);	
	BOOL SetInterfaceBoardActuatorOutput(int scanType, int *testResult);
	void ShutDown(BOOL copySpreadsheet);
	BOOL createNewSpreadsheet();	
	BOOL storeSerialNumberAndPartNumber(char *ptrSerialNumber, char *ptrPartNumber);
	BOOL storeTestResult(int col, int testResult, char *ptrTestData);
	BOOL SaveSpreadsheet();		
	BOOL ActivateSpectrometer();
	BOOL getSpectrometerConfiguration();
	BOOL startMeasurement();
	BOOL stopMeasurement();
	BOOL isSpectrometerDataReady();
	BOOL CloseSpectrometer();				
	int  RunSpectrometerScan(BOOL filterState);
	BOOL InitializeSystem (bool UseSpectrometer);	
	BOOL InitializeHP34401();
	BOOL InitializeInterfaceBoard();
	BOOL InitializePowerSupply();
	BOOL InitializeSpectrometer();
	void resetSubStepNumber();
	double applyPolynomial(double ADcounts);
	void ClearLog();
	int  TTL_inputTest();	
	void ConfigureFont(CFont &ptrFont, int fontHeight, int fontWidth, BOOL flgBold);	
	BOOL InitializeDisplayInstructions();
	void DisplayLog(char *newString);
	BOOL DisplayTestEditBox(int boxNumber, int passFailStatus);	
	void DisplayStatusBarText(int panel, LPCTSTR strText);	
	void DisplayIntructions(int stepNumber, int status);
	void resetDisplays(BOOL enableTextBoxes);	
	void ClearInstructBox();
	void clearDataArray(double *ptrData);	
	BOOL getScanData(float *ptrScanData);
	BOOL fetchSpectrometerData(float *ptrScopeData);	
	BOOL LoadScanData(char *filename, double *ptrLambda, float *ptrScanData);
	BOOL writeSpectrometerDataToFile(const char *filename, double *ptrLambda, float *ptrScopeData);			
	BOOL copyToIrradianceSpectrum (float *ptrIntensity);
	BOOL convertScopeDataToIrradiance(float *scopeData, float *irradianceIntensity, BOOL enableLinearization, BOOL enableSLC);
	BOOL linearizeScanData(float *ptrScanData, float *ptrLinearizedData);	
	BOOL getColorTemp(float *colorTemperature);
	BOOL getIrradiance (float startWavelength, float endWavelength, float *irradianceIntegral);
	BOOL getCenterWavelengthAndFWHM (float startWavelength, float endWavelength, float *centerWavelength, float wavelengthAmplitude, float *FWHM, int splineFactor);
	BOOL getCenterWavelengthAndFWHM (float startWavelength, float endWavelength, int splineFactor, float *centerWavelength, float *wavelengthAmplitude, float *FWHM);
	BOOL LoadIRFpartNumbers(char *filename, int *numberOfPartNumbers);	
	BOOL CheckAndCreateDirectory (char *path);
	BOOL InitializeFonts();	
	BOOL loadINIfileBinary();
	void storeINIfileBinary();		
	BOOL storeExcelFilenames(char *ptrPrimaryFilename,  char *ptrSecondaryFilename);
	BOOL validateSpreedsheetFilename(char *ptrFilename);
	BOOL loadINIpartNumbers();
	BOOL storeINIpartNumbers();	
	void writeBinaryFile(std::ostream& out, float value);
	void readBinaryFile(std::istream& in, float& value);			
	void readBinaryConfigFile(std::istream& in, DeviceConfigType* l_pDeviceData);
	void writeBinaryConfigFile(std::ostream& out, DeviceConfigType* l_pDeviceData);
	BOOL loadConfigFile();
	BOOL SpreadsheetIsOpen(char *filename);
	int	 stringLength(char *ptrString);
	void clearString(char *ptrString);
	BOOL LoadAdminPassword();
	BOOL isValidPassword(char *ptrPassword);
	BOOL storeAdminPassword();
	void displaySpectrometerCOMerror();
	void displayInterfaceCOMerrorr();
	BOOL CheckFaultSignal(int *testResult);		
	TestApp(CWnd* pParent = NULL);
	~TestApp();
};

#endif

