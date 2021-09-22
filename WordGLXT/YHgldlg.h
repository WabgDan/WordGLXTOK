#if !defined(AFX_YHGLDLG_H__9703467D_34F0_40D2_B8B0_63329419EA6E__INCLUDED_)
#define AFX_YHGLDLG_H__9703467D_34F0_40D2_B8B0_63329419EA6E__INCLUDED_
#include "Users.h"
#include "Rizhib.h"

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// YHgldlg.h : header file
//
extern CUsers user;
/////////////////////////////////////////////////////////////////////////////
// CYHgldlg dialog

class CYHgldlg : public CDialog
{
// Construction
public:
	CString strjb;
	CYHgldlg(CWnd* pParent = NULL);   // standard constructor
	CRizhib zhi;
	CTime t;
// Dialog Data
	//{{AFX_DATA(CYHgldlg)
	enum { IDD = IDD_YHGL };
	CComboBox	m_jb;
	CString	m_name;
	CString	m_pwd;
	CString	m_pwd1;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CYHgldlg)
	public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	HICON m_hIcon;
	// Generated message map functions
	//{{AFX_MSG(CYHgldlg)
	virtual BOOL OnInitDialog();
	virtual void OnOK();
	afx_msg void OnButton1();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_YHGLDLG_H__9703467D_34F0_40D2_B8B0_63329419EA6E__INCLUDED_)
