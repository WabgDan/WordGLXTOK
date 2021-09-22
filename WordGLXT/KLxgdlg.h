#if !defined(AFX_KLXGDLG_H__0A4E7A0B_0DCA_4C10_8EB9_E738CC49135C__INCLUDED_)
#define AFX_KLXGDLG_H__0A4E7A0B_0DCA_4C10_8EB9_E738CC49135C__INCLUDED_
#include "Users.h"
#include "Rizhib.h"

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// KLxgdlg.h : header file
//
extern CUsers user;
/////////////////////////////////////////////////////////////////////////////
// CKLxgdlg dialog

class CKLxgdlg : public CDialog
{
// Construction
public:
	CKLxgdlg(CWnd* pParent = NULL);   // standard constructor
	CRizhib zhi;
	CTime t;
// Dialog Data
	//{{AFX_DATA(CKLxgdlg)
	enum { IDD = IDD_KLXG };
	CStatic	m_name;
	CString	m_opwd;
	CString	m_npwd2;
	CString	m_npwd1;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CKLxgdlg)
	public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	HICON m_hIcon;
	// Generated message map functions
	//{{AFX_MSG(CKLxgdlg)
	virtual void OnOK();
	virtual BOOL OnInitDialog();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_KLXGDLG_H__0A4E7A0B_0DCA_4C10_8EB9_E738CC49135C__INCLUDED_)
