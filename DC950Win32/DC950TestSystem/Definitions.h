#pragma once

#ifndef DEFINITIONS_H
#define DEFINITIONS_H

#define MAXROWS 65536 // 1048576 

#define NUM_REMOTE_TEST_VALUES 8

#define FILE_LOAD_ERROR 0
#define USING_PRIMARY_SPREADSHEET 1
#define USING_SECONDARY_SPREADSHEET 2

#define MAX_CONTROL_VOLTAGE_ERROR 1.0

#define MAXPASSWORD MAXLINE
#define MINPASSWORD 5

#define NR_DEFECTIVE_PIXELS     30
#define	PIXEL_ARRAY_SIZE 2048 // MAX_NR_PIXELS // was 2048

#define calIntTime l_pDeviceData->m_Irradiance.m_IntensityCalib.m_CalInttime
#define IntensityCal l_pDeviceData->m_Irradiance.m_IntensityCalib.m_aCalibConvers
#define LinCorrect l_pDeviceData->m_Detector.m_aNLCorrect
#define NonLinLowCounts l_pDeviceData->m_Detector.m_aLowNLCounts
#define NonLinHighCounts l_pDeviceData->m_Detector.m_aHighNLCounts

#define MAXINIVALUES 39
#define MAX_PORTNUMBER 50

#define BLACK RGB(0,0,0)
#define YELLOW RGB(255,255,0)
#define RED RGB(255,0,0)
#define GREEN RGB(0,255,0)
#define CYAN RGB(128,255,255)
#define WHITE RGB(255,255,255)

#define MAXLOG 65535
#define MAXLINE 512
//#define NOPOWERSUPPLY 
#define USE_SPREADSHEET
//#define NOSERIAL
//#define NOSPECTROMETER 
#define MAX_AC_VOLTAGE 254

enum SCAN_TYPE {
	FILTER_OPEN = 0,
	FILTER_CLOSED,
	DARK_SCAN
};


#define VALID_PART_NUMBER_LENGTH 12
#define ON TRUE
#define OFF FALSE
#define ENABLED TRUE
#define DISABLED FALSE

// #include "Definitions.h"
#define MAXFILENAME 256 // was 64
#define CHARWIDTH 340
#define MAXSTRING 128
#define MAX_PART_NUMBERS 64
#define MAX_PWM 6215
#define MULTIMETER 1
#define BUFFERSIZE 128
#define FONTHEIGHT 26
#define FONTWIDTH 10
#define MAXTRIES 5
#define MULTIMETER 1

#define STARTUP 0
#define LEFTPANEL 0
#define RIGHTPANEL 1

#define DATAIN 0
#define DATAOUT 1

typedef struct DATAstring {
	char String[MAXSTRING];
} var;

enum DATA_COLUMNS {
	TEST_DATE = 0,
	SERIAL_NUMBER,	
	PART_NUMBER,
	HI_POT,
	GROUND_BOND,
	LAMP_OFF,
	LAMP_ON,
	VREF_TEST,
	REM_VOLT_0,
	REM_VOLT_1,
	REM_VOLT_2,
	REM_VOLT_3,
	REM_VOLT_4,
	REM_VOLT_5,
	INHIBIT,		
	SWEEP_0,
	SWEEP_1,
	SWEEP_2,
	SWEEP_3,
	SWEEP_4,
	SWEEP_5,		
	FAULT_TEST,	
	IRRADIANCE_CLOSED,
	WAVELENGTH_CLOSED,
	MIN_AMPLITUDE_CLOSED,
	FWHM_CLOSED,	
	IRRADIANCE_OPEN,
	COLOR_TEMP_OPEN,	
	FINAL_RESULT	
};


#define LASTCOLUMN FINAL_RESULT
#define TOTAL_COLUMNS LASTCOLUMN + 1

#define SERIAL_NUMBER_COLUMN SERIAL_NUMBER
#define PART_NUMBER_COLUMN PART_NUMBER
#define TEST_DATE_COLUMN TEST_DATE

#define FIRST_TEST_COLUMN SCAN_BARCODE
#define FIRST_FILTER_TEST_COLUMN ACTUATOR_OPEN

enum MULTIMETER_INPUTS {
	LAMP = 0,
	VREF,
	CONTROL_VOLTAGE
};

enum COM_PORTS {
	INTERFACE_BOARD = 0,
	HP_METER,
	AC_POWER_SUPPLY
};

enum TIMER_STATES {
	TIMER_PAUSED = 0,
	TIMER_RUN
};

enum STATUS {
	NOT_DONE_YET,	
	PASS,
	FAIL,
	SYSTEM_ERROR
};

enum STEP_TYPE {			
	MANUAL = 0,
	AUTO
};

#define REMOTE_TEST 7
#define BARCODE_SCAN 1

struct TestStep
{
public:
	char *lineOne;
	char *lineTwo;
	char *lineThree;
	char *lineFour;
	char *lineFive;		
	INT  editBoxNumber;	
	int  stepType;
	int  testID;
	BOOL enableENTER;
	BOOL enablePREVIOUS;
	BOOL enableRETRY;
	BOOL enablePASS;
	BOOL enableFAIL;
	BOOL enableADMIN;
};

enum { HI_POT_EDIT = 1,   GROUND_BOND_EDIT,  POT_LOW_EDIT,  POT_HIGH_EDIT,	REMOTE_EDIT,    AC_SWEEP_EDIT,   ACTUATOR_EDIT,  FILTER_ON_EDIT,  FILTER_OFF_EDIT,  FINAL_ASSEMBLY_EDIT};

#define HI_POT_TEST 3
#define GROUND_BOND_TEST 4
#define END_STANDARD_UNIT_TESTS 10
#define REMOTE_TEST_SETUP 7
#define DARK_SCAN_SETUP 10
#define FINAL_ASSEMBLY 16
#define FINAL_PASS 17
#define FINAL_FAIL 18
#define RETRY_TEST 19
#define SYSTEM_FAILED 20
#define TOTAL_STEPS 21

struct testData
{
	CString result;
	COLORREF color;
};

#endif

