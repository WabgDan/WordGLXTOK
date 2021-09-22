#if !defined(AFX_FILEATTRI_H__CF7342D5_E753_4326_A211_9761CF9978DE__INCLUDED_)
#define AFX_FILEATTRI_H__CF7342D5_E753_4326_A211_9761CF9978DE__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// FileAttri.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// FileAttri dialog

class FileAttri : public CDialog
{
// Construction
public:
	void ShowFileAttri();
	CString pathname;
	FileAttri(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(FileAttri)
	enum { IDD = IDD_LookFileAttri };
	CButton	m_ar;
	CButton	m_dir;
	CButton	m_vol;
	CButton	m_readonly;
	CButton	m_system;
	CButton	m_normal;
	CButton	m_hide;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(FileAttri)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(FileAttri)
	afx_msg void OnBNAddFile();
	virtual BOOL OnInitDialog();
	virtual void OnCancel();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_FILEATTRI_H__CF7342D5_E753_4326_A211_9761CF9978DE__INCLUDED_)
