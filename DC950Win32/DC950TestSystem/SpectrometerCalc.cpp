// SpectrometerCalc.cpp  - math and calibration routines for Avantes spectrometer
#include "stdafx.h"
#include "DC950Test.h"
#include "DC950TestDialog.h"
#include <string.h>
#include "TestApp.h"
#include "Definitions.h"
#include "avaspec.h"
#include "irrad.h"
#include <fstream>
#include <iostream>

using namespace std;

unsigned char resolution = 0;
unsigned char observer = 0;

// SPECTROMETER GLOBAL VARIABLES
extern DeviceConfigType* l_pDeviceData;
extern int m_NrDevices;
extern long handleSpectrometer;
extern unsigned int numberOfPixels;

extern DataType irradianceSpectrum;
extern double Lambda[];
extern float openFilterData[];
extern float darkData[];
extern float closedFilterData[];
extern float irradianceIntensity[];
extern float scanDataCorrected[];
extern float linearizedData[];

extern float	m_currentIntegrationTime;
extern uint16	m_StartPixel;
extern uint16	m_StopPixel;
extern uint16	m_NumberOfScans;
extern BOOL		m_enableLinearCorrection;
extern BOOL		m_enableStrayLightCorrection;
double			tempIntensity[PIXEL_ARRAY_SIZE];
double			dblIrradianceIntensity[PIXEL_ARRAY_SIZE];


BOOL TestApp::convertScopeDataToIrradiance(float *scopeData, float *irradianceIntensity, BOOL enableLinearization, BOOL enableSLC)
{
int i;	

	float intTimeFactor = calIntTime / m_currentIntegrationTime;
	float *ptrIntensityCal;

	if (scopeData == NULL || irradianceIntensity == NULL) return FALSE;
	if (l_pDeviceData->m_Irradiance.m_IntensityCalib.m_aCalibConvers ==  NULL) return FALSE;
	if (m_StopPixel > PIXEL_ARRAY_SIZE) m_StopPixel = PIXEL_ARRAY_SIZE;
		
	for (i = 0; i < m_StopPixel; i++) scanDataCorrected[i] = scopeData[i] - darkData[i];
	ptrIntensityCal = l_pDeviceData->m_Irradiance.m_IntensityCalib.m_aCalibConvers;

	if (enableLinearization){
		linearizeScanData (scanDataCorrected, linearizedData);
		for (i = 0; i < m_StopPixel; i++) irradianceIntensity[i] = intTimeFactor * (linearizedData[i] / ptrIntensityCal[i]);		
	}
	else {
		for (i = 0; i < m_StopPixel; i++) irradianceIntensity[i] = intTimeFactor * (scanDataCorrected[i] / ptrIntensityCal[i]);
	}

	/*
	if (enableSLC){
		for (i = 0; i < m_StopPixel; i++) tempIntensity[i] = (double) irradianceIntensity[i];
		AVS_SuppressStrayLight(handleSpectrometer, (float) 1.0, tempIntensity, dblIrradianceIntensity);
		for (i = 0; i < m_StopPixel; i++) irradianceIntensity[i] = (float) dblIrradianceIntensity[i];
	}
	*/
	
	return TRUE;
}

BOOL TestApp::linearizeScanData (float *ptrScanData, float *ptrLinearizedData){
float lowCorrection, highCorrection;
int i;
				
	double lowCounts = l_pDeviceData->m_Detector.m_aLowNLCounts;
	double highCounts = l_pDeviceData->m_Detector.m_aHighNLCounts;
		
	lowCorrection = (float) applyPolynomial(lowCounts);
	highCorrection = (float) applyPolynomial(highCounts);
		
	for (i = 0; i < m_StopPixel; i++){
		double ADcounts = (double) ptrScanData[i];
		if (ADcounts < lowCounts) ptrLinearizedData[i] = (float) (ADcounts / lowCorrection);
		else if (ADcounts > highCounts)	ptrLinearizedData[i] = (float) (ADcounts / highCorrection);	
		else ptrLinearizedData[i] = (float) (ADcounts / applyPolynomial(ADcounts));
	}	
	return TRUE;
}

double TestApp::applyPolynomial(double ADcounts){
double normalizedCountsPerSecond;
	normalizedCountsPerSecond = 
	l_pDeviceData->m_Detector.m_aNLCorrect[0] +
	l_pDeviceData->m_Detector.m_aNLCorrect[1] * pow (ADcounts, 1) +
	l_pDeviceData->m_Detector.m_aNLCorrect[2] * pow (ADcounts, 2) +
	l_pDeviceData->m_Detector.m_aNLCorrect[3] * pow (ADcounts, 3) +
	l_pDeviceData->m_Detector.m_aNLCorrect[4] * pow (ADcounts, 4) +
	l_pDeviceData->m_Detector.m_aNLCorrect[5] * pow (ADcounts, 5) +
	l_pDeviceData->m_Detector.m_aNLCorrect[6] * pow (ADcounts, 6) +
	l_pDeviceData->m_Detector.m_aNLCorrect[7] * pow (ADcounts, 7);
	if (normalizedCountsPerSecond == 0.0) return (1.0);
	else return normalizedCountsPerSecond;
}


BOOL TestApp::getColorTemp(float *colorTemperature){	
double smallx,smally,smallz,bigX,bigY,bigZ,u,v,colortemp;

	int dataSize = sizeof(irradianceSpectrum);
	int err = Color_GetColorOfLightParam (
			&irradianceSpectrum,
	        resolution,
            observer,
	        &smallx,
	        &smally,
	        &smallz,
	        &bigX,
	        &bigY,
	        &bigZ,
	        &u,
	        &v,
	        &colortemp);
	*colorTemperature = (float) colortemp;
	return TRUE;								  
}

BOOL TestApp::getIrradiance (float startWavelength, float endWavelength, float *irradianceIntegral)
{
double dbIrradianceIntegral;

	Radio_GetIrradiance(&irradianceSpectrum, (double) startWavelength, (double) endWavelength, &dbIrradianceIntegral);
	*irradianceIntegral = (float) dbIrradianceIntegral;
	return FALSE;
}

BOOL TestApp::getCenterWavelengthAndFWHM (float startWavelength, float endWavelength, int splineFactor, float *centerWavelength, float *wavelengthAmplitude, float *FWHM)
{
double a_peaknm;
double a_peakamp;
double a_fwhm;
double a_cwlnm;
double a_cwlamp;
double a_centroidnm;
double a_centroidamp;

	int err = Radio_GetPeak(&irradianceSpectrum, (double) startWavelength, (double) endWavelength, (short int) splineFactor, &a_peaknm, &a_peakamp, &a_fwhm, &a_cwlnm, &a_cwlamp, &a_centroidnm, &a_centroidamp);
	*centerWavelength = (float) a_cwlnm;
	*wavelengthAmplitude = (float) a_cwlamp;
	*FWHM = (float) a_fwhm;
	return TRUE;
}