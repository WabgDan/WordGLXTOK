// Zdmlb.h: interface for the CZdmlb class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_ZDMLB_H__CDAA4E35_3780_42F4_9426_7359972233E7__INCLUDED_)
#define AFX_ZDMLB_H__CDAA4E35_3780_42F4_9426_7359972233E7__INCLUDED_

#include "ADOConn.h"
#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CZdmlb  
{
private:
	int DWbh;
	int LBbh;
	CString LBmc;

public:
	CZdmlb();
	virtual ~CZdmlb();

	CStringArray a_DWbh;
	CStringArray a_LBbh;
	CStringArray a_LBmc;

	int GetDWbh();
	void SetDwbh(int iDWbh);
	int GetLBbh();
	void SetLBbh(int iLBbh);
	CString GetLBmc();
	void SetLBmc(CString cLBmc);
	void sql_insert();
	void sql_update(int iDWbh,int iLBbh);
	void sql_delete(int iDWbh,int iLBbh);
	void sql_deletedw(int iLBbh);
	void Load_dep();
	int HaveId(int iDWbh,int iLBbh);
};

#endif // !defined(AFX_ZDMLB_H__CDAA4E35_3780_42F4_9426_7359972233E7__INCLUDED_)
