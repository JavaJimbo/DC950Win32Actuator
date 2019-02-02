// SpecDiagnostics.h
#pragma once

#ifndef __AFXWIN_H__
#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols
#include "NumSpinCtrl.h"
#include "NumEdit.h"

class SpecDiagnostics : public CDialog
{
public:
	SpecDiagnostics(CWnd* pParent = NULL);   // standard constructor
	~SpecDiagnostics();

	enum { IDD = IDD_DIALOG_SPECTROMETER };	
	
	afx_msg void OnNMThemeChangedScrollbar1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMThemeChangedScrollComInterface(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnDeltaposSpin1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	void StartTimer();
	void StopTimer();
	void timerHandler();	
	BOOL startSpectrometerTest(BOOL filterState, int intPWM);
	CEdit   m_EditBox_Spectrometer;
	CEdit*  pEdit;
	CButton m_static_ButtonTestSpectrometer;
	CButton m_static_ButtonActuator;
	CButton m_static_ButtonScan;
	CButton m_static_ButtonDarkScan;
	CButton m_static_checkStrayLightCorrection;
	void ClearLog();
	void DisplayLog(char *newString);	
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	afx_msg void OnClose();
	afx_msg void OnPaint();
	virtual BOOL OnInitDialog();
	DECLARE_MESSAGE_MAP()

public:
	afx_msg void OnBnClickedButtonRunTest();
	afx_msg void OnBnClickedButtonHaltTest();
	afx_msg void OnBnClickedButtonScan();
	afx_msg void OnBnClickedButtonDarkScan();
	afx_msg void OnBnClickedButtonActuator();	
	afx_msg void OnCheckedBoxSLC();
	afx_msg void OnBnClickedButtonLoadEeprom();
};