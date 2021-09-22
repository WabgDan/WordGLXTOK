#if !defined(AFX_WDLBIEDLG_H__C175C508_0123_4730_BBAF_82B3479AC9A2__INCLUDED_)
#define AFX_WDLBIEDLG_H__C175C508_0123_4730_BBAF_82B3479AC9A2__INCLUDED_

#include "Zdmlb.h"
#include "Zdxxb.h"
#include "Rizhib.h"
#include "Users.h"
#include "Dwxxb.h"


#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Wdlbiedlg.h : header file
//
extern CUsers user;
/////////////////////////////////////////////////////////////////////////////
// CWdlbiedlg dialog

class CWdlbiedlg : public CDialog
{
// Construction
public:
	CWdlbiedlg(CWnd* pParent = NULL);   // standard constructor
	CTime t;
	CRizhib zhi;
	
	//CDwxxb dwxxb;  //定义单位表的一个类
	//CZdmlb  dmlb; //定义类别表的一个类
// Dialog Data
	//{{AFX_DATA(CWdlbiedlg)
	enum { IDD = IDD_WDLB };
	CComboBox	m_combo5;
	int		m_dwbh;
	int		m_lbbh;
	CString	m_lbmc;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CWdlbiedlg)
	public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	HICON m_hIcon;
	// Generated message map functions
	//{{AFX_MSG(CWdlbiedlg)
	virtual void OnOK();
	virtual BOOL OnInitDialog();
	afx_msg void OnSelchangeCombo5();
	afx_msg void OnRADIOModify();
	afx_msg void OnRadioAdd();
	afx_msg void OnRADIODel();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
private:
	void LBDel();
	void LBModify();
	void LBAdd();
	int RadioFlag;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_WDLBIEDLG_H__C175C508_0123_4730_BBAF_82B3479AC9A2__INCLUDED_)
