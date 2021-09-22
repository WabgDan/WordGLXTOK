#if !defined(AFX_WDGLDLG_H__9C951B13_F35A_44B5_B6C8_79A7F917AA13__INCLUDED_)
#define AFX_WDGLDLG_H__9C951B13_F35A_44B5_B6C8_79A7F917AA13__INCLUDED_
#include "Dwxxb.h"
#include "Zdmlb.h"
#include "Zdxxb.h"
#include "Users.h"
#include "Rizhib.h"
#include "ADOConn.h"


#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// WDgldlg.h : header file
//
extern CUsers user;

/////////////////////////////////////////////////////////////////////////////
// CWDgldlg dialog

class CWDgldlg : public CDialog
{
// Construction
public:
	void SetCurTab(UINT m_index=0);
	int wdbh;
	int lbbh;
	int dwbh;
	int str;   //µ¥Ñ¡°´Å¥
	CString strText;
	CDwxxb dwb;
	CZdmlb mlb;
	CZdxxb xxb;
	CRizhib zhi;
	CTime t;
	UINT tabindex;
	CWDgldlg(CWnd* pParent = NULL);		// standard constructor
// Dialog Data
	//{{AFX_DATA(CWDgldlg)
	enum { IDD = IDD_WDgldlg };
	CEdit	m_wdlb;
	CComboBox	m_combo4;
	CEdit	m_dwmc;
	CComboBox	m_combo1;
	CComboBox	m_combo3;
	CListCtrl	m_list;
	CEdit	m_wjlj;
	CTabCtrl	m_tab;
	CString	m_wdmc;
	CString	m_gjz;
	CString	m_memo;
	int		m_wdbh;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CWDgldlg)
	public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL
public:
	bool m_bAutoComplete;
// Implementation
protected:

	HICON m_hIcon;
	// Generated message map functions
	//{{AFX_MSG(CWDgldlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnWjljxz();
	afx_msg void OnSelchangeTab1(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnSelchangingTab1(NMHDR* pNMHDR, LRESULT* pResult);
	virtual void OnOK();
	afx_msg void OnRadio1();
	afx_msg void OnRadio2();
	afx_msg void OnDblclkList1(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnSelchangeCombo3();
	afx_msg void OnBUTTONDelWD();
	afx_msg void OnClickList1(NMHDR* pNMHDR, LRESULT* pResult);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_WDGLDLG_H__9C951B13_F35A_44B5_B6C8_79A7F917AA13__INCLUDED_)
