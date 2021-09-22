#if !defined(AFX_DIALOGIN_H__0D90A8D2_2496_4AAB_B490_B4A567CB7B90__INCLUDED_)
#define AFX_DIALOGIN_H__0D90A8D2_2496_4AAB_B490_B4A567CB7B90__INCLUDED_

#include "Rizhib.h"
#include "Users.h"

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Dialogin.h : header file
//
extern CUsers user;
/////////////////////////////////////////////////////////////////////////////
// CDialogin dialog


class CDialogin : public CDialog
{
// Construction
public:
	CString jb;
	CRizhib zhi;
	CTime t;
	CDialogin(CWnd* pParent = NULL);   // standard constructor
	
// Dialog Data
	//{{AFX_DATA(CDialogin)
	enum { IDD = IDD_DIALOGIN };
	CButton	m_Close;
	CButton	m_OK;
	CButton	m_Cancel;
	CString	m_name;
	CString	m_pwd;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDialogin)
	public:
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	CBrush m_bkBrush;

	HICON m_hIcon;
	// Generated message map functions
	//{{AFX_MSG(CDialogin)
	virtual BOOL OnInitDialog();
	virtual void OnOK();
	afx_msg void OnClose();
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
private:
//	CBitmap m_Cross;
//	CBitmap m_Title;
	HBITMAP m_BitmapCancel;
	HBITMAP m_BitmapOK;
	HBITMAP m_BitmapClose;
public:
    afx_msg void OnBnClickedOk();
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DIALOGIN_H__0D90A8D2_2496_4AAB_B490_B4A567CB7B90__INCLUDED_)
