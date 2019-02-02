#pragma once

class CMSExcel
{
protected:
	HRESULT m_hr;
	IDispatch*	m_pEApp;
	IDispatch*  m_pBooks;
	IDispatch*	m_pActiveBook;
private:
	HRESULT Initialize(bool bVisible=true);
public:
	CMSExcel(void);
	~CMSExcel(void);
	HRESULT SetCellValueAndColor(LPCTSTR szRange, LPCTSTR szValue, LPCSTR szFormat, COLORREF crColor, BOOL boldFlag);
	HRESULT TurnOffDisplayAlerts();
	HRESULT Save(LPCTSTR szFilename); 
	HRESULT SetVisible(bool bVisible=true);
	HRESULT OpenExcelBook(LPCTSTR szFilename, bool bVisible=true);
	HRESULT GetExcelValue(LPCTSTR szCell, CString &sValue);
	HRESULT Quit();
};
