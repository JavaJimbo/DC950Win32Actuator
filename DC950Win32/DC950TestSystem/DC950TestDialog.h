// DC950TestDialog.h : header file
//

#pragma once

#ifndef __AFXWIN_H__
#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols
#include "ReadOnlyEdit.h"
#include "PasswordEntryDlg.h"
#include "AdminDialog1.h"
#include "AdminDialog2.h"


class CTestDialog : public CDialog
{
// Construction
public:
	CTestDialog(CWnd* pParent = NULL);	// standard constructor
	~CTestDialog();						// destructor

// Dialog Data
	enum { IDD = IDD_SERIALCTRLDEMO_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
	// Generated message map functions
	virtual BOOL OnInitDialog();
	// afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();	
	afx_msg void OnAdd();
	afx_msg void OnBack();
	afx_msg void OnText();
	
	HICON m_hIcon;	
	DECLARE_MESSAGE_MAP()
public:
	CStatic m_static_BarcodeLabel;
	CStatic m_static_PartNumber;		
	CStatic m_static_SpreadsheetFilename;
	CButton m_static_ButtonEnter;		
	CButton m_static_ButtonRetry;
	CButton m_static_ButtonPass;
	CButton m_static_ButtonFail;
	// CButton m_static_ButtonHalt;
	CMFCButton m_static_ButtonHalt;
	CButton m_static_ButtonPrevious;	
	CButton m_static_ButtonAdmin1;
	CButton m_static_ButtonAdmin2;
	CEdit   m_BarcodeEditBox;
	CEdit   m_PartNumberEditBox;
	CEdit   m_EditBox_Log;
	
	
	PasswordEntry m_PED;
	Admin1 m_Admin1;
	Admin2 m_SpecDiag;
	CEdit*  pEdit;
	CReadOnlyEdit*	pInstruct;

	CReadOnlyEdit   m_EditBox_Instruct;
	CReadOnlyEdit	m_EditBox_Test1;
	CReadOnlyEdit	m_EditBox_Test2;
	CReadOnlyEdit	m_EditBox_Test3;
	CReadOnlyEdit	m_EditBox_Test4;
	CReadOnlyEdit	m_EditBox_Test5;
	CReadOnlyEdit	m_EditBox_Test6;
	CReadOnlyEdit	m_EditBox_Test7;
	CReadOnlyEdit	m_EditBox_Test8;
	CReadOnlyEdit	m_EditBox_Test9;
	CReadOnlyEdit	m_EditBox_Test10;
	CReadOnlyEdit	m_EditBox_OutputFileName;
	#define LAST_TEST_EDIT_BOX 10
		
	CStatusBar m_StatusBar;		
	BOOL enableAdmin;		

	void timerHandler();
	void StartTimer();
	void StopTimer();	
	void resetSystem();
	void ConfigureFont(CFont &ptrFont, int fontHeight, int fontWidth, BOOL flgBold);
	void CreateStatusBars();	
	void TestHandler(int caller);		
	void SetDialogColor(COLORREF rgb);
	BOOL PreTranslateMessage(MSG* pMsg);	
public:			
	afx_msg void OnClickedButtonHalt();
	afx_msg void OnClickedButtonPrevious();	
	afx_msg void OnClickedButtonPass();
	afx_msg void OnClickedButtonFail();
	afx_msg void OnClickedButtonEnter();
	afx_msg void OnClickedButtonAdmin1();
	afx_msg void OnClickedButtonAdmin2();		
	afx_msg void OnClickedButtonRetry();
	afx_msg void OnClickActuatorOPEN();
	afx_msg void OnClickActuatorCLOSED();
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnClickFileExit();
protected:
public:

};
