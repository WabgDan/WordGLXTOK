// Rizhib.h: interface for the CRizhib class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_RIZHIB_H__DE5B54B1_DC28_486E_9BB5_C32D67EDB8FB__INCLUDED_)
#define AFX_RIZHIB_H__DE5B54B1_DC28_486E_9BB5_C32D67EDB8FB__INCLUDED_

#include "ADOConn.h"
#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CRizhib  
{
private:
	CString Name;
	CString DLsj;
	CString DZ;

public:
	CRizhib();
	virtual ~CRizhib();

	CString GetName();
	void SetName(CString cName);
	CString GetDLsj();
	void SetDLsj(CString cDLsj);
	CString GetDZ();
	void SetDZ(CString cDZ);
	void sql_insert();
};

#endif // !defined(AFX_RIZHIB_H__DE5B54B1_DC28_486E_9BB5_C32D67EDB8FB__INCLUDED_)
