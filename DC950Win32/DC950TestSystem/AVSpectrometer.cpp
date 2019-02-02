// AVSPectrometer - routines for communicating with Avantes spectrometer
// Inlcudes function for receiving and sending configuration information
// as well as initiating scans amd fetching readings.

#include "stdafx.h"
#include "DC950Test.h"
#include "DC950TestDialog.h"
#include <string.h>
#include "TestApp.h"
#include "Definitions.h"
#include "avaspec.h"
#include <fstream>
#include <iostream>
#include <sstream>

using namespace std;

const char *CONFIGfilename = "CONFIGfile.bin";
// SPECTROMETER GLOBAL VARIABLES
DeviceConfigType* l_pDeviceData;
int m_NrDevices = 0;
long handleSpectrometer = INVALID_AVS_HANDLE_VALUE;
unsigned int numberOfPixels = 0;

double Lambda[PIXEL_ARRAY_SIZE];
float openFilterData[PIXEL_ARRAY_SIZE];
float darkData[PIXEL_ARRAY_SIZE];
float closedFilterData[PIXEL_ARRAY_SIZE];
float irradianceIntensity[PIXEL_ARRAY_SIZE];
float scanDataCorrected[PIXEL_ARRAY_SIZE];
float linearizedData[PIXEL_ARRAY_SIZE];
DataType irradianceSpectrum;

float m_currentIntegrationTime = 165; // was 100
uint16	m_StartPixel = 0;
uint16	m_StopPixel = 1061;
uint16	m_NumberOfScans = 4;
BOOL m_enableLinearCorrection = FALSE;
BOOL m_enableStrayLightCorrection = FALSE;
unsigned long m_numberOfAverages = 300; //  was 100
char* l_pData;
AvsIdentityType*    spectrometerID = NULL;

void TestApp::clearDataArray(double *ptrData) {
	for (int i = 0; i < PIXEL_ARRAY_SIZE; i++) ptrData[i] = 0.0;	
}

BOOL TestApp::ActivateSpectrometer() {

#ifdef NOSPECTROMETER 
	return TRUE;
#endif
	
	// If the global handleSpectrometer is valid,
	// then spectrometer must have been activated already,
	// so quit here if it's ready to go:
	if (handleSpectrometer != INVALID_AVS_HANDLE_VALUE) return TRUE;

	// Otherwise, activate and check result:
	long newHandle = AVS_Activate(spectrometerID);

	if (newHandle == INVALID_AVS_HANDLE_VALUE) return FALSE;
	else handleSpectrometer = newHandle;

	// If activated successfully, read all stored configuration data:
	if (getSpectrometerConfiguration()) return TRUE;
	else return FALSE;
}


BOOL TestApp::getSpectrometerConfiguration() {
	// DeviceConfigType* l_pDeviceData = NULL;
	unsigned short intNumberOfPixels = 0;
	unsigned int l_Size;

#ifdef NOSPECTROMETER
	return TRUE;
#endif
	if (handleSpectrometer == INVALID_AVS_HANDLE_VALUE) return FALSE;

	char a_Fpga[16];
	char a_As5216[16];
	char a_Dll[16];

	int l_Res = AVS_GetVersionInfo(handleSpectrometer, a_Fpga, a_As5216, a_Dll);

	if (ERR_SUCCESS == l_Res)
	{		
		//DisplayInstructions(1, a_Fpga);
		//DisplayInstructions(1, a_As5216);
		//DisplayInstructions(1, a_Dll);
	}

	if (ERR_SUCCESS == AVS_GetNumPixels(handleSpectrometer, &intNumberOfPixels))
			numberOfPixels = (unsigned int) intNumberOfPixels;

	l_Res = AVS_GetParameter(handleSpectrometer, 0, &l_Size, l_pDeviceData);

	if (l_Res == ERR_INVALID_SIZE)
	{
		l_pDeviceData = (DeviceConfigType*)new char[l_Size];
	}

	l_Res = AVS_GetParameter(handleSpectrometer, l_Size, &l_Size, l_pDeviceData);
	
	if (l_Res != ERR_SUCCESS) return FALSE;

	m_StartPixel = l_pDeviceData->m_StandAlone.m_Meas.m_StartPixel;
	m_StopPixel = l_pDeviceData->m_StandAlone.m_Meas.m_StopPixel;
	if (m_StopPixel > PIXEL_ARRAY_SIZE) m_StopPixel = PIXEL_ARRAY_SIZE;
	
	if (ERR_SUCCESS != AVS_GetLambda(handleSpectrometer, Lambda)) return FALSE;
	msDelay(100);
	return TRUE;
}


BOOL TestApp::startMeasurement()
{
	double l_NanoSec = 0.0;
	long l_Ticks = 0;
	long l_Dif = 0;
	unsigned int l_Time = 0;
	uint32 l_FPGAClkCycles = 0;
	long l_Res = 0;
	unsigned char l_Status = 0;
	MeasConfigType l_PrepareMeasData;

#ifdef NOSPECTROMETER
	return TRUE;
#endif		
		if (spectrometerID->Status == USB_IN_USE_BY_APPLICATION)
		{
			// StartPixel				
			l_PrepareMeasData.m_StartPixel = m_StartPixel;
			// StopPixel
			l_PrepareMeasData.m_StopPixel = m_StopPixel;
			// IntegrationTime			
			l_PrepareMeasData.m_IntegrationTime = m_currentIntegrationTime; // fltIntegrationTime;
			// IntegrationDelay
			l_NanoSec = -20.83;
			l_FPGAClkCycles = (uint32)(6.0 * (l_NanoSec + 20.84) / 125.0);
			l_PrepareMeasData.m_IntegrationDelay = l_FPGAClkCycles;
			// Number of Averages
			l_PrepareMeasData.m_NrAverages = m_numberOfAverages;
			// DarkCorrectionType
			l_PrepareMeasData.m_CorDynDark.m_Enable = 1;
			l_PrepareMeasData.m_CorDynDark.m_ForgetPercentage = 100;
			// SmoothingType
			l_PrepareMeasData.m_Smoothing.m_SmoothPix = 1;
			l_PrepareMeasData.m_Smoothing.m_SmoothModel = 0;
			// SaturationDetection
			l_PrepareMeasData.m_SaturationDetection = 1;
			// TriggerType
			l_PrepareMeasData.m_Trigger.m_Mode = 0;
			l_PrepareMeasData.m_Trigger.m_Source = 0;
			l_PrepareMeasData.m_Trigger.m_Source = 1;
			l_PrepareMeasData.m_Trigger.m_SourceType = 0;
			l_PrepareMeasData.m_Trigger.m_SourceType = 1;

			// ControlSettingsType
			l_PrepareMeasData.m_Control.m_StrobeControl = 0;
			l_NanoSec = 0;
			l_FPGAClkCycles = (uint32)(6.0 * l_NanoSec / 125.0);
			l_PrepareMeasData.m_Control.m_LaserDelay = l_FPGAClkCycles;
			l_NanoSec = 0;
			l_FPGAClkCycles = (uint32)(6.0 * l_NanoSec / 125.0);
			l_PrepareMeasData.m_Control.m_LaserWidth = l_FPGAClkCycles;
			l_PrepareMeasData.m_Control.m_LaserWaveLength = 0;
			l_PrepareMeasData.m_Control.m_StoreToRam = 0;

			if (ERR_SUCCESS != AVS_PrepareMeasure(handleSpectrometer, &l_PrepareMeasData))
				return FALSE;			

			// Start Measurement
			m_StartPixel = l_PrepareMeasData.m_StartPixel;
			m_StopPixel = l_PrepareMeasData.m_StopPixel;
						
			if (ERR_SUCCESS != AVS_GetLambda(handleSpectrometer, Lambda))
				return FALSE;
			if (ERR_SUCCESS != AVS_Measure(handleSpectrometer, NULL, m_NumberOfScans))
				return FALSE;
			else return TRUE;
		}
		else return FALSE;
}


BOOL TestApp::stopMeasurement() {
#ifdef NOSPECTROMETER
	return TRUE;
#endif
	if (spectrometerID == NULL) return TRUE;
	if (spectrometerID->Status == USB_IN_USE_BY_APPLICATION)
	{
		if (ERR_SUCCESS != AVS_StopMeasure(handleSpectrometer))
			return FALSE;
		else return TRUE;		
	}
	else return FALSE;
}

BOOL TestApp::isSpectrometerDataReady() {
#ifdef NOSPECTROMETER
	return TRUE;
#endif

	if (AVS_PollScan(handleSpectrometer))
		return TRUE;
	else return FALSE;
}

BOOL TestApp::CloseSpectrometer()
{
#ifdef NOSPECTROMETER
	return TRUE;
#endif

	if (handleSpectrometer) AVS_Done();
	return TRUE;
}

BOOL TestApp::InitializeSpectrometer() {	
	long l_Port = 0;	
	unsigned int        l_Size = 0;
	unsigned int        l_RequiredSize = 0;

#ifdef NOSPECTROMETER
	return TRUE;
#endif
	if (m_NrDevices == 1) return TRUE;

	msDelay(100);
	l_Port = AVS_Init(0);  	
	
	if (l_Port > 0)
	{
		// UpdateList();
		msDelay(100);
		m_NrDevices = AVS_UpdateUSBDevices();		
		if (m_NrDevices != 1) return FALSE;
		
		msDelay(100);

		m_NrDevices = AVS_GetNrOfDevices();
		l_RequiredSize = m_NrDevices * sizeof(AvsIdentityType);

		if (l_RequiredSize > 0)
		{
			delete[] l_pData;
			l_pData = new char[l_RequiredSize];
			l_Size = l_RequiredSize;
			spectrometerID = (AvsIdentityType*)l_pData;
			m_NrDevices = AVS_GetList(l_Size, &l_RequiredSize, spectrometerID);
		}
		if (m_NrDevices != 1) return FALSE;
		else return TRUE;		
	}
	else AVS_Done();		
	
	return FALSE;
}


// Read binary float value
void TestApp::readBinaryConfigFile(std::istream& in, DeviceConfigType* l_pDeviceData){
	in.read(reinterpret_cast<char *>(l_pDeviceData), sizeof(DeviceConfigType));
}
// Write binary float value
void TestApp::writeBinaryConfigFile(std::ostream& out, DeviceConfigType* l_pDeviceData){
	out.write(reinterpret_cast<char *>(l_pDeviceData), sizeof(DeviceConfigType));
}


BOOL TestApp::loadConfigFile(){
	std::ifstream inFile;

	l_pDeviceData = (DeviceConfigType*)new char[sizeof(DeviceConfigType)];
	
	inFile.open(CONFIGfilename, ios::in|ios::binary); // Open INI file		
	if (inFile.is_open()) readBinaryConfigFile(inFile, l_pDeviceData);
	else
	{		
		AfxMessageBox((LPCTSTR)"SYSTEM ERROR - No CONFIG file found");
		return FALSE;
	}
	m_StopPixel = l_pDeviceData->m_StandAlone.m_Meas.m_StopPixel;
	if (m_StopPixel > PIXEL_ARRAY_SIZE) m_StopPixel = PIXEL_ARRAY_SIZE;
	return TRUE;
}


BOOL TestApp::copyToIrradianceSpectrum (float *ptrIntensity)
{
int i;

	if (ptrIntensity == NULL || m_StopPixel >= PIXEL_ARRAY_SIZE) return FALSE;	

		for (i = 0; i < m_StopPixel; i++){
			irradianceSpectrum.intensity[i] = ptrIntensity[i];
			irradianceSpectrum.wl[i] = Lambda[i];
		}
		for (i = m_StopPixel; i < MAX_PIXELS; i++){
			irradianceSpectrum.intensity[i] = 0.0;
			irradianceSpectrum.wl[i] = Lambda[i];
		}

	return TRUE;
}


BOOL TestApp::getScanData(float *ptrScanData)
{
uint32 l_Time = 0;
double scopeData[PIXEL_ARRAY_SIZE];
double doubleData = 0.0;
int i;

#ifdef NOSPECTROMETER
	return TRUE;
#endif	
	if (m_StopPixel >= PIXEL_ARRAY_SIZE || ptrScanData == NULL) return FALSE;

	if (ERR_SUCCESS == AVS_GetScopeData(handleSpectrometer, &l_Time, scopeData) ){
		for (i = 0; i < m_StopPixel; i++) ptrScanData[i] = (float) scopeData[i];		
		for (i = m_StopPixel; i < PIXEL_ARRAY_SIZE; i++) ptrScanData[i] = (float) 0.0;			
	}
	return TRUE;	
}

// Start static Excel Automation demo
BOOL TestApp::LoadScanData(char *filename, double *ptrLambda, float *ptrScanData) 
{
float scanData;
double lambda;
char *ptrLine;
int i;
std::string strLine;

	if (filename == NULL || ptrLambda == NULL || ptrScanData == NULL) return FALSE;

	std::ifstream inFile(filename);

	i = 0;
	while (std::getline(inFile, strLine)){
		ptrLine = strdup(strLine.c_str());
		sscanf(ptrLine, "%lf %f", &lambda, &scanData);
		free (ptrLine);
		ptrLambda[i] = lambda;
		ptrScanData[i] = scanData;
		i++;
		if (i >= PIXEL_ARRAY_SIZE) return FALSE;
	}	
	return TRUE;
}

