// Zdxxb.h: interface for the CZdxxb class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_ZDXXB_H__4427C5FC_A4E3_4253_BC7D_44B5A90A0E68__INCLUDED_)
#define AFX_ZDXXB_H__4427C5FC_A4E3_4253_BC7D_44B5A90A0E68__INCLUDED_

#include "ADOConn.h"
#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CZdxxb  
{
private:
	int LBbh;
	int WDbh;
	int DWbh;
	CString GJz;
	CString WDmc;
	CString WJlj;
	CString Memo;
	CString Tjrxm;

public:
	CZdxxb();
	virtual ~CZdxxb();

	CStringArray a_WDbh; 
	CStringArray a_LBbh;
	CStringArray a_WJlj;
	CStringArray a_DWbh;
	CStringArray a_WDmc;
	CStringArray a_GJz;

	int GetDWbh();
	void SetDWbh(int iDWbh);
	int GetLBbh();
	void SetLBbh(int iLBbh);
	int GetWDbh();
	void SetWDbh(int iWDbh);
	CString GetGJz();
	void SetGJz(CString cGJz);
	CString GetWDmc();
	void SetWDmc(CString cWDmc);
	CString GetWJlj();
	void SetWJlj(CString cWJlj);
	CString GetMemo();
	void SetMemo(CString cMemo);
	CString GetTjrxm();
	void SetTjrxm(CString cTjrxm);
	void sql_insert();
	void sql_update(int iWDbh);
	void sql_deletelb(int iDWbh,int iLBbh);
	void sql_delete(int iWDbh);
	void sql_deletedw(int iDWbh);
	int sql_selectwdmc(CString cWDmc);
	void Load_dep();
	int HaveId(int iDWbh,int iLBbh,int iWDbh);
};

#endif // !defined(AFX_ZDXXB_H__4427C5FC_A4E3_4253_BC7D_44B5A90A0E68__INCLUDED_)
