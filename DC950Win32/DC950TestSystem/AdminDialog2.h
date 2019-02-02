// Admin2.h
#pragma once

#ifndef __AFXWIN_H__
#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols
#include "NumSpinCtrl.h"
#include "NumEdit.h"



class Admin2 : public CDialog
{
public:
	Admin2(CWnd* pParent = NULL);   // standard constructor
	~Admin2();

	enum { IDD = IDD_DIALOG_ADMIN2 };	
	
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	void StartTimer();
	void StopTimer();
	void timerHandler();	
	int RunSpectrometerTest(int scanType);
	CEdit*  pEdit, ptrPartNumber;
	

	CButton m_checkEnableStrayLightCorrection, m_checkEnableLinearCorrection, m_static_ButtonPassword;
	CNumSpinCtrl m_SpinIntegrationTime, m_SpinNumAverages, m_SpinNumScans;
	CNumEdit m_EditIntegrationTime, m_EditNumAverages, m_EditNumScans;
	CEdit m_EditBox_PartNumber, m_EditBox_Password;
	CEdit m_EditBox_Min1, m_EditBox_Min2, m_EditBox_Min3, m_EditBox_Min4, m_EditBox_Min5, m_EditBox_Min6;
	CEdit m_EditBox_Max1, m_EditBox_Max2, m_EditBox_Max3, m_EditBox_Max4, m_EditBox_Max5, m_EditBox_Max6
;	
	void ClearLog();
	void DisplayLog(char *newString);	
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	afx_msg void OnClose();
	afx_msg void OnPaint();
	virtual BOOL OnInitDialog();
	BOOL PreTranslateMessage(MSG* pMsg);	
	BOOL CopyTextFromPartNumberEditBox();
	DECLARE_MESSAGE_MAP()

public:
	afx_msg void OnCheckedBoxSLC();
	afx_msg void OnBnClickedCheckLinear();
	afx_msg void OnBnClickedButtonPassword();

public:
	afx_msg void OnNMThemeChangedScrollbar1(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMThemeChangedScrollComInterface(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnDeltaposSpin1(NMHDR *pNMHDR, LRESULT *pResult);
};