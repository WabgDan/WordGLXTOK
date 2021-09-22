// WordGLXT.h : main header file for the WORDGLXT application
//

#if !defined(AFX_WORDGLXT_H__82BF2AC1_FC9C_4BBC_9EEC_198D72B4EF96__INCLUDED_)
#define AFX_WORDGLXT_H__82BF2AC1_FC9C_4BBC_9EEC_198D72B4EF96__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CWordGLXTApp:
// See WordGLXT.cpp for the implementation of this class
//

class CWordGLXTApp : public CWinApp
{
public:
	CWordGLXTApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CWordGLXTApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CWordGLXTApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_WORDGLXT_H__82BF2AC1_FC9C_4BBC_9EEC_198D72B4EF96__INCLUDED_)
