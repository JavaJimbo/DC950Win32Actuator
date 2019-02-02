/*	MSExcel.cpp - modified Excel routines adapted from code writtne by Naren Neelamegam - see below
*
*	See Code Project article: "MS Office OLE Automation Using C++" by Naren Neelamegam, 5 Apr 2009
*	Link: https://www.codeproject.com/Articles/34998/MS-Office-OLE-Automation-Using-C
*
*	Also helpful - Microsoft's article "How to automate Excel from C++ without using MFC or #import"
*	Link: https://support.microsoft.com/en-us/help/216686/how-to-automate-excel-from-c-without-using-mfc-or-import
*	
*	3-28-18 JBS:	
*/
#include "StdAfx.h"
#include "MSExcel.h"
#include "OLEMethod.h"
#include <string.h>
#include <stdio.h>
#include <stdlib.h>
#include <comdef.h>

CMSExcel::CMSExcel(void)
{
	m_hr=S_OK;
	m_pEApp=NULL;
	m_pBooks=NULL;
	m_pActiveBook=NULL;
}

HRESULT CMSExcel::Initialize(bool bVisible)
{
	CoInitialize(NULL);
	CLSID clsid;
	m_hr = CLSIDFromProgID(L"Excel.Application", &clsid);
	if(SUCCEEDED(m_hr))
	{
		m_hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void **)&m_pEApp);
		if(FAILED(m_hr)) m_pEApp=NULL;
	}
	{
		m_hr=SetVisible(bVisible);
	}
	return m_hr;
}


HRESULT CMSExcel::SetVisible(bool bVisible)
{
	if(!m_pEApp) return E_FAIL;
	VARIANT x;
	x.vt = VT_I4;
	x.lVal = bVisible;
	m_hr=OLEMethod(DISPATCH_PROPERTYPUT, NULL, m_pEApp, L"Visible", 1, x);

	return m_hr;
}

HRESULT CMSExcel::OpenExcelBook(LPCTSTR szFilename, bool bVisible)
{
	if(m_pEApp==NULL) 
	{
		if(FAILED(m_hr=Initialize(bVisible)))
			return m_hr;
	}

	{
		VARIANT result;
		VariantInit(&result);
		m_hr=OLEMethod(DISPATCH_PROPERTYGET, &result, m_pEApp, L"Workbooks", 0);
		m_pBooks = result.pdispVal;
	}	

	{
		COleVariant sFname(szFilename);
		VARIANT result;
		VariantInit(&result);
		m_hr=OLEMethod(DISPATCH_PROPERTYGET, &result, m_pBooks, L"Open", 1,sFname.Detach());
		m_pActiveBook = result.pdispVal;
	}
	return m_hr;
}

// HRESULT CMSExcel::SaveAs(LPCTSTR szFilename, int nSaveAs)
HRESULT CMSExcel::Save(LPCTSTR szFilename)
{
	COleVariant varFilename(szFilename);
	VARIANT varAs;
	varAs.vt=VT_I4;
	// varAs.intVal=nSaveAs;
	varAs.intVal = -4143; // This apparently is the "SaveAs" value needed for this operation - not sure what it does
	m_hr=OLEMethod(DISPATCH_METHOD,NULL,m_pActiveBook,L"SaveAs",2,varAs,varFilename.Detach());
	return m_hr;
}

HRESULT CMSExcel::Quit()
{
	if(m_pEApp==NULL) return E_FAIL;
	DISPID dispID;
	m_hr = OLEMethod(DISPATCH_METHOD, NULL, m_pEApp, L"Quit", 0);  
		

	LPOLESTR ptName = L"Quit"; 	
	 m_hr = m_pEApp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);

	if(SUCCEEDED(m_hr))
	{
		DISPPARAMS dp = { NULL, NULL, 0, 0 };
		m_hr = m_pEApp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, 
									&dp, NULL, NULL, NULL);
	}
	
	m_pEApp->Release();
	m_pEApp=NULL;
	return m_hr;
}


CMSExcel::~CMSExcel(void)
{
	Quit();
	CoUninitialize();
}

HRESULT CMSExcel::TurnOffDisplayAlerts()
{
	if(m_pEApp==NULL) return E_FAIL;
	VARIANT displayAlerts;
	VariantInit ( & displayAlerts);
	displayAlerts.vt = VT_BOOL;
	displayAlerts.boolVal = false;
	m_hr = OLEMethod (DISPATCH_PROPERTYPUT, NULL, m_pEApp, L"DisplayAlerts", 1, displayAlerts);
	return m_hr;
}


HRESULT CMSExcel::SetCellValueAndColor(LPCTSTR szRange, LPCTSTR szValue, LPCSTR szFormat, COLORREF crColor, BOOL boldFlag)
{
	static BOOL testFlag = TRUE;
	if(!m_pEApp) return E_FAIL;
	if(!m_pActiveBook) return E_FAIL;
	IDispatch *pSheets;
	{
		VARIANT result;
		VariantInit(&result);

		m_hr=OLEMethod(DISPATCH_PROPERTYGET, &result, m_pActiveBook, L"Sheets", 0);
		pSheets = result.pdispVal;
	}
	IDispatch *pSheet;
	{
		VARIANT result;
		VariantInit(&result);
		VARIANT itemn;
		itemn.vt = VT_I4;
		itemn.lVal = 1;
		m_hr=OLEMethod(DISPATCH_PROPERTYGET, &result, pSheets, L"Item", 1, itemn);
		pSheet = result.pdispVal;
	}

	IDispatch* pRange;
	{
		COleVariant oleRange(szRange);
		VARIANT result;
		VariantInit(&result);
		m_hr=OLEMethod(DISPATCH_PROPERTYGET, &result, pSheet, L"Range", 1, oleRange.Detach());
		pRange = result.pdispVal;
	}

	{
		COleVariant oleValue(szValue);
		m_hr=OLEMethod(DISPATCH_PROPERTYPUT, NULL, pRange, L"Value", 1, oleValue.Detach());
	}

	VARIANT x;
	x.vt = VT_I4;
	x.lVal = 2;
	m_hr=OLEMethod(DISPATCH_PROPERTYPUT, NULL, pRange, L"HorizontalAlignment", 1, x);

	COleVariant oleParam(szFormat);
	m_hr=OLEMethod(DISPATCH_PROPERTYPUT, NULL, pRange, L"NumberFormat", 1, oleParam.Detach());

	IDispatch *pFont;
	{
		VARIANT result;
		VariantInit(&result);
		m_hr=OLEMethod(DISPATCH_PROPERTYGET, &result, pRange, L"Font",0);
		pFont=result.pdispVal;
	}
	{
		COleVariant oleName("Arial");
		m_hr=OLEMethod(DISPATCH_PROPERTYPUT, NULL, pFont, L"Name", 1, oleName.Detach());
		VARIANT x;
		x.vt = VT_I4;
		x.lVal = 10;
		m_hr=OLEMethod(DISPATCH_PROPERTYPUT, NULL, pFont, L"Size", 1, x);
		x.lVal = crColor; 
		m_hr=OLEMethod(DISPATCH_PROPERTYPUT, NULL, pFont, L"Color", 1, x);
		x.lVal = boldFlag;
		m_hr=OLEMethod(DISPATCH_PROPERTYPUT, NULL, pFont, L"Bold", 1, x);
	}

	pFont->Release();
	pRange->Release();
	pSheet->Release();
	pSheets->Release();
	return m_hr;
}

HRESULT CMSExcel::GetExcelValue(LPCTSTR szCell, CString &sValue)
{
	if(!m_pEApp) return E_FAIL;
	if(!m_pActiveBook) return E_FAIL;

	IDispatch *pSheets;
	{
		VARIANT result;
		VariantInit(&result);

		m_hr=OLEMethod(DISPATCH_PROPERTYGET, &result, m_pActiveBook, L"Sheets", 0);
		pSheets = result.pdispVal;
	}
	IDispatch *pSheet;
	{
		VARIANT result;
		VariantInit(&result);
		VARIANT itemn;
		itemn.vt = VT_I4;
		itemn.lVal = 1;
		m_hr=OLEMethod(DISPATCH_PROPERTYGET, &result, pSheets, L"Item", 1, itemn);
		pSheet = result.pdispVal;
	}

	IDispatch* pRange;
	{
		COleVariant oleRange(szCell);
		VARIANT result;
		VariantInit(&result);
		m_hr=OLEMethod(DISPATCH_PROPERTYGET, &result, pSheet, L"Range", 1, oleRange.Detach());
		pRange = result.pdispVal;
	}


	{
		VARIANT result;
		VariantInit(&result);
		m_hr=OLEMethod(DISPATCH_PROPERTYGET, &result, pRange, L"Value", 0);		
		if (result.vt == VT_NULL)
			sValue = CString("");
		else sValue = (char*)(_bstr_t)result;		 
	}

	pRange->Release();
	pSheet->Release();
	pSheets->Release();
	return m_hr;
}

