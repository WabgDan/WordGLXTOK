// WordGLXTDlg.h : header file
//

#if !defined(AFX_WORDGLXTDLG_H__86F2AA3C_11E2_4AF1_A7EF_095C2F9A89D1__INCLUDED_)
#define AFX_WORDGLXTDLG_H__86F2AA3C_11E2_4AF1_A7EF_095C2F9A89D1__INCLUDED_


#include "msword9.h"
#include "atlbase.h"
#include "Zdmlb.h"
#include "Zdxxb.h"
#include "Dwxxb.h"
#include "Dwdandlg.h"
#include "Wdlbiedlg.h"
//#include "Chaxundlg.h"
#include "YHgldlg.h"
#include "KLxgdlg.h"
#include "Dialogin.h"
#include "WDgldlg.h"
#include "Rizhib.h"
#include "ADOConn.h"
#include "FileAttri.h"


#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

extern CWordGLXTApp theApp;

/////////////////////////////////////////////////////////////////////////////
// CWordGLXTDlg dialog

class CWordGLXTDlg : public CDialog
{

private:
	HTREEITEM arrays[10],brrays[20],hitem[100];
	HTREEITEM m_root,temp;
	CDwxxb dwb;
	CZdmlb mlb;
	CZdxxb xxb;
	CRizhib zhi;
	CImageList m_treeImageList;
	CTime t;
	CStatusBarCtrl m_StatusBar;  
	CString strWordpath;  //记录word路径
	CString strText;     //暂存word文档的内容
    void AddtoTree(HTREEITEM m_node);
public:	
	CWordGLXTDlg(CWnd* pParent = NULL);	// standard constructor
	
	
// Dialog Data
	//{{AFX_DATA(CWordGLXTDlg)
	enum { IDD = IDD_WORDGLXT_DIALOG };
	CRichEditCtrl	m_richedit;
	CTreeCtrl	m_tree;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CWordGLXTDlg)
	public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CWordGLXTDlg)
	BOOL OnToolTipNotify(UINT id, NMHDR *pNMHDR, LRESULT *pResult);
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
    afx_msg void OnDblclkTree1(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnMenudwdan();
	afx_msg void OnMenuwdlb();
	afx_msg void OnMenuliulwd();
	afx_msg void OnMenuyhgl();
	afx_msg void OnMenumodpwd();
	afx_msg void OnMenuaddwd();
	afx_msg void OnMenudelwd();
	afx_msg void OnMenumodwd();
	afx_msg void OnMenurzgl();
	afx_msg void OnMENULookFileAttri();
	afx_msg void OnExit();
	afx_msg void OnButtonAdd();
	afx_msg void OnButtonDwDA();
	afx_msg void OnButtonWDLB();
	afx_msg void OnButtonModify();
	afx_msg void OnButtonDel();
	afx_msg void OnButtonScan();
	afx_msg void OnButtonAttri();
	afx_msg void OnButtonUserManage();
	afx_msg void OnButtonmimaModify();
	afx_msg void OnButtonLogManage();
	afx_msg void OnButtonExit();
	afx_msg void OnMenuDBBackUp();
	afx_msg void OnMenuDBRestore();
	afx_msg void OnButtonBackUp();
	afx_msg void OnButtonRestore();
	afx_msg void OnSystemExit();
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_WORDGLXTDLG_H__86F2AA3C_11E2_4AF1_A7EF_095C2F9A89D1__INCLUDED_)
