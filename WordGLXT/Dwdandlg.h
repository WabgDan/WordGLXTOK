#if !defined(AFX_DWDANDLG_H__612710E2_1476_4F28_B74D_DD72B2169274__INCLUDED_)
#define AFX_DWDANDLG_H__612710E2_1476_4F28_B74D_DD72B2169274__INCLUDED_

#include "Dwxxb.h"
#include "Zdmlb.h"
#include "Zdxxb.h"
#include "Rizhib.h"
#include "Users.h"

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Dwdandlg.h : header file
//
extern CUsers user;

/////////////////////////////////////////////////////////////////////////////
// CDwdandlg dialog

class CDwdandlg : public CDialog
{
// Construction
public:
	CDwdandlg(CWnd* pParent = NULL);   // standard constructor
	CRizhib zhi;
	CTime t;

	//CDwxxb dwxxb;
	void SetCurTab(UINT m_index);
private:
	void DelDW();//删除单位信息
	void ModifyDW();//修改单位信息
	void AddDW();//添加单位信息
	int RadioFlag;
// Dialog Data
	//{{AFX_DATA(CDwdandlg)
	enum { IDD = IDD_DWDAN };
	CComboBox	m_combo_dwmc;
	CListCtrl	m_list;
	CTabCtrl	m_tab;
	CString	m_lxdz;
	CString	m_lxr;
	CString	m_memo;
	int		m_dwbh;
	CString	m_lxdh;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDwdandlg)
	public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	HICON m_hIcon;
	// Generated message map functions
	//{{AFX_MSG(CDwdandlg)
	virtual void OnOK();
	virtual BOOL OnInitDialog();
	afx_msg void OnSelchangeTab1(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnSelchangingTab1(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnDblclkList1(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnSelchangeCombo2();
	afx_msg void OnRADIOModify();
	afx_msg void OnRadioAdd();
	afx_msg void OnRADIODel();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DWDANDLG_H__612710E2_1476_4F28_B74D_DD72B2169274__INCLUDED_)
