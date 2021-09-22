// WDgldlg.cpp : implementation file
//

#include "stdafx.h"
#include "WordGLXT.h"
#include "WDgldlg.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CWDgldlg dialog



CWDgldlg::CWDgldlg(CWnd* pParent /*=NULL*/)
	: CDialog(CWDgldlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CWDgldlg)
	m_wdmc = _T("");
	m_gjz = _T("");
	m_memo = _T("");
	m_wdbh = 0;
	tabindex = 0;
	//}}AFX_DATA_INIT

}


void CWDgldlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CWDgldlg)
	DDX_Control(pDX, IDC_COMBO4, m_combo4);
	DDX_Control(pDX, IDC_COMBO1, m_combo1);
	DDX_Control(pDX, IDC_COMBO3, m_combo3);
	DDX_Control(pDX, IDC_LIST1, m_list);
	DDX_Control(pDX, IDC_EDIT6, m_wjlj);
	DDX_Control(pDX, IDC_TAB1, m_tab);
	DDX_Text(pDX, IDC_EDIT4, m_wdmc);
	DDX_Text(pDX, IDC_EDIT5, m_gjz);
	DDX_Text(pDX, IDC_EDIT7, m_memo);
	DDX_Text(pDX, IDC_EDIT3, m_wdbh);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CWDgldlg, CDialog)
	//{{AFX_MSG_MAP(CWDgldlg)
	ON_BN_CLICKED(IDC_WJLJXZ, OnWjljxz)
	ON_NOTIFY(TCN_SELCHANGE, IDC_TAB1, OnSelchangeTab1)
	ON_NOTIFY(TCN_SELCHANGING, IDC_TAB1, OnSelchangingTab1)
	ON_BN_CLICKED(IDC_RADIO1, OnRadio1)
	ON_BN_CLICKED(IDC_RADIO2, OnRadio2)
	ON_NOTIFY(NM_DBLCLK, IDC_LIST1, OnDblclkList1)
	ON_CBN_SELCHANGE(IDC_COMBO3, OnSelchangeCombo3)
	ON_BN_CLICKED(IDC_BUTTON_DelWD, OnBUTTONDelWD)
	ON_NOTIFY(NM_CLICK, IDC_LIST1, OnClickList1)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CWDgldlg message handlers

BOOL CWDgldlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	m_hIcon = AfxGetApp()->LoadIcon(IDI_CHILDICON4);
	// TODO: Add extra initialization here
	SetIcon(m_hIcon, TRUE);	
	TC_ITEM tci;
	tci.mask=TCIF_TEXT;
	tci.pszText="基本信息";
	m_tab.InsertItem(0,&tci);
	tci.pszText="信息删除";
	m_tab.InsertItem(1,&tci);

	dwb.Load_dep();
	mlb.Load_dep();
	xxb.Load_dep();

	t=CTime::GetCurrentTime();

	UpdateData(true);
	for(int i=0;i<xxb.a_WDbh.GetSize();i++) //根据文档编号在文档表中搜索文档名称
	
		m_combo1.AddString(xxb.a_WDmc.GetAt(i));//往标签页2下的组合框添加文档名称
	
	for(int i=0;i<dwb.a_DWbh.GetSize();i++) //根据单位编号在单位表中搜索单位名称
	
		m_combo3.AddString(dwb.a_DWmc.GetAt(i));//往标签页1下的组合框添加单位名称
	    
	
	m_list.SetExtendedStyle(LVS_EX_FLATSB|LVS_EX_FULLROWSELECT|LVS_EX_GRIDLINES);
	m_list.InsertColumn(0,"单位名称",LVCFMT_LEFT,100,0);
	m_list.InsertColumn(1,"文档类别",LVCFMT_LEFT,100,1);
	m_list.InsertColumn(2,"文档编号",LVCFMT_LEFT,100,2);
	m_list.InsertColumn(3,"文档名称",LVCFMT_LEFT,100,3);
	m_list.InsertColumn(4,"关键字",LVCFMT_LEFT,100,4);
	m_list.InsertColumn(5,"文档路径",LVCFMT_LEFT,100,5);
	m_list.InsertColumn(6,"备注",LVCFMT_LEFT,100,6);
	CString dwmc[100],wdlb[100],pp;
	//根据单位编号返回单位名称
    int i = 0;
	for(i=0;i<xxb.a_WDmc.GetSize();i++)
	{
		for(int j=0;j<dwb.a_DWbh.GetSize();j++)
		{
			if(atoi(xxb.a_DWbh.GetAt(i))==atoi(dwb.a_DWbh.GetAt(j)))
			{
				dwmc[i]=dwb.a_DWmc.GetAt(j);
			}
		}
		//根据类别编号返回类别名称
		for(int j=0;j<mlb.a_DWbh.GetSize();j++)
		{
			if(atoi(xxb.a_DWbh.GetAt(i))==atoi(mlb.a_DWbh.GetAt(j)) && atoi(xxb.a_LBbh.GetAt(i))==atoi(mlb.a_LBbh.GetAt(j)))
			{
				wdlb[i]=mlb.a_LBmc.GetAt(j);
			}
		}
		
	}
	ADOConn m_AdoConn;
	m_AdoConn.OnInitADOConn();
	CString sql;
	sql.Format("select* from Zdxxb order by wdbh desc");
	m_AdoConn.GetRecordSet((_bstr_t)sql);
	while(m_AdoConn.m_pRecordset->adoEOF==0)
	{
		m_list.InsertItem(0,"");
		m_list.SetItemText(0,0,dwmc[i-1]);
		m_list.SetItemText(0,1,wdlb[i-1]);
		m_list.SetItemText(0,2,(char*)(_bstr_t)m_AdoConn.m_pRecordset->GetCollect("wdbh"));
		m_list.SetItemText(0,3,(char*)(_bstr_t)m_AdoConn.m_pRecordset->GetCollect("wdmc"));
		m_list.SetItemText(0,4,(char*)(_bstr_t)m_AdoConn.m_pRecordset->GetCollect("gjz"));
		m_list.SetItemText(0,5,(char*)(_bstr_t)m_AdoConn.m_pRecordset->GetCollect("wjlj"));
		m_list.SetItemText(0,6,(char*)(_bstr_t)m_AdoConn.m_pRecordset->GetCollect("memo"));
		i--;
		m_AdoConn.m_pRecordset->MoveNext();
	}
	m_AdoConn.ExitConnect();
	//根据菜单选项使不同的单选按钮处于选中状态
	if(str==0)
	{
		CButton* tempbutton = (CButton*)GetDlgItem(IDC_RADIO1);
		tempbutton->SetCheck(1);
	}
	else
	{
		CButton* tempbutton = (CButton*)GetDlgItem(IDC_RADIO2);
		tempbutton->SetCheck(1);
	}
	//调用SetCurTab()
	SetCurTab(tabindex);

	m_wdbh = xxb.a_WDmc.GetSize()+1;  //把默认编号为1的给弄成自动排序的

	UpdateData(false);
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void CWDgldlg::OnWjljxz()     //文件路径选择
{
	// TODO: Add your control notification handler code here
	CFileDialog file(true,NULL,NULL,OFN_HIDEREADONLY|OFN_OVERWRITEPROMPT,"All Files(*.*)|*.*| |",AfxGetMainWnd());
		if(file.DoModal()==IDOK)
		{
			
			strText= file.GetPathName();
		    m_wjlj.SetWindowText(strText);
			
			m_wdmc = file.GetFileName();   //自动添加文档名称
			int index=m_wdmc.ReverseFind('.');
			CString tem = m_wdmc;
			if(index!=-1) tem.Delete(index,m_wdmc.GetLength()-index);
			m_gjz  = tem;
			UpdateData(false); //将变量m_wdmc的数据输出到编辑框中
		}
}



void CWDgldlg::OnSelchangeTab1(NMHDR* pNMHDR, LRESULT* pResult) 
{
	// TODO: Add your control notification handler code here
	switch(m_tab.GetCurSel())
	{
	case 0: 
		//标签0的控件显示
		//GetDlgItem(IDC_EDIT1)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_COMBO3)->ShowWindow(SW_SHOW);
		//GetDlgItem(IDC_EDIT2)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_COMBO4)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_EDIT3)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_EDIT4)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_EDIT5)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_EDIT6)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_EDIT7)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC1)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC2)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC3)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC4)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC5)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC6)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC7)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC8)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_WJLJXZ)->ShowWindow(SW_SHOW);
		GetDlgItem(IDOK)->ShowWindow(SW_SHOW);
		GetDlgItem(IDCANCEL)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_RADIO1)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_RADIO2)->ShowWindow(SW_SHOW);
		break;
	case 1: 
		//标签1的控件显示
		GetDlgItem(IDC_LIST1)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_COMBO1)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_BUTTONDEL)->ShowWindow(SW_SHOW);
		break;
	}
	*pResult = 0;
}

void CWDgldlg::OnSelchangingTab1(NMHDR* pNMHDR, LRESULT* pResult) 
{
	// TODO: Add your control notification handler code here
	switch(m_tab.GetCurSel())
	{
	case 0: 
		//标签0的控件隐藏
		//GetDlgItem(IDC_EDIT1)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_COMBO3)->ShowWindow(SW_HIDE);
		//GetDlgItem(IDC_EDIT2)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_COMBO4)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT3)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT4)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT5)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT6)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT7)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC1)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC2)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC3)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC4)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC5)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC6)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC7)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC8)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_WJLJXZ)->ShowWindow(SW_HIDE);
		GetDlgItem(IDOK)->ShowWindow(SW_HIDE);
		GetDlgItem(IDCANCEL)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_RADIO1)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_RADIO2)->ShowWindow(SW_HIDE);
		break;
	case 1: 
		//标签1的控件隐藏
		GetDlgItem(IDC_LIST1)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_COMBO1)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_BUTTONDEL)->ShowWindow(SW_HIDE);
		break;
	}
	*pResult = 0;
}

void CWDgldlg::OnOK() //保存按钮的代码
{
	// TODO: Add extra validation here
	UpdateData(true);
	CString strdwmc,strwdlb;
    
	if(m_combo3.GetCurSel()==CB_ERR) 
	{
		MessageBox("单位名称不能为空,请选择单位!","文档管理系统");
	    return;
	}
	else 
		m_combo3.GetLBText(m_combo3.GetCurSel(),strdwmc);
	
	
	if(m_combo4.GetCurSel()==CB_ERR)
	{
		MessageBox("文档类别不能为空,请选择文档类别!","文档管理系统");
		return;
	}
	else 
		m_combo4.GetLBText(m_combo4.GetCurSel(),strwdlb);

	if(m_wdmc=="")
	{
		MessageBox("文档名称不能为空","文档管理系统");
		return;
	}
	CString strwjlj;
	m_wjlj.GetWindowText(strwjlj);
	if(strwjlj=="")
	{
		MessageBox("文档路径不能为空","文档管理系统");
		return;
	}

	int dw=0,lb=0;
	for(int i=0;i<dwb.a_DWbh.GetSize();i++)//根据单位编号搜索单位名称
	{
		if(strdwmc==dwb.a_DWmc.GetAt(i))
		{
			dwbh=atoi(dwb.a_DWbh.GetAt(i));
			dw++; 
		}
	}
	if(dw==0)
	{
		MessageBox("单位名称不存在","文档管理系统");
		return;
	}
	for(int i=0;i<mlb.a_DWbh.GetSize();i++)//根据单位编号搜索单位类别
	{
		if(dwbh==atoi(mlb.a_DWbh.GetAt(i)) && strwdlb==mlb.a_LBmc.GetAt(i))
		{
			lbbh=atoi(mlb.a_LBbh.GetAt(i));//类别编号
			lb++;
		}
	}
	if(lb==0)
	{
		MessageBox("文档类别不存在","文档管理系统");
		return;
	}

	xxb.SetDWbh(dwbh);
	xxb.SetLBbh(lbbh);
	xxb.SetWDbh(m_wdbh);
	xxb.SetWDmc(m_wdmc);
	xxb.SetGJz(m_gjz);
	    
	xxb.SetWJlj(strwjlj);
	xxb.SetMemo(m_memo);
	xxb.SetTjrxm(user.GetUsername());
	switch(str)
	{
		case 0: //添加
			if(xxb.HaveId(dwbh,lbbh,m_wdbh)==1)
			{
				MessageBox("文档已存在","文档管理系统");
				return;
			}
			xxb.sql_insert();
			zhi.SetDLsj(t.Format("%y-%m-%d"));
			zhi.SetName(user.GetUsername());
			zhi.SetDZ("添加文档");
			zhi.sql_insert();
           
			break;
		case 1:  //修改
			xxb.sql_update(m_wdbh);
			zhi.SetDLsj(t.Format("%y-%m-%d"));
			zhi.SetName(user.GetUsername());
			zhi.SetDZ("修改文档");
			zhi.sql_insert();
			break;
	}
	dw=0;
	lb=0;
	CDialog::OnOK();
}

void CWDgldlg::OnRadio1() //选择添加
{
	// TODO: Add your control notification handler code here
	str=0;
	
}
void CWDgldlg::OnRadio2() //选择修改
{
	// TODO: Add your control notification handler code here
	str=1;		
}

void CWDgldlg::SetCurTab(UINT m_index)
{
	m_tab.SetCurSel(m_index);
	if(m_index==0)
	{
		//标签1的控件隐藏
		GetDlgItem(IDC_LIST1)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_COMBO1)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_BUTTONDEL)->ShowWindow(SW_HIDE);
		//标签0的控件显示
		//GetDlgItem(IDC_EDIT1)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_COMBO3)->ShowWindow(SW_SHOW);
		//GetDlgItem(IDC_EDIT2)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_COMBO4)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_EDIT3)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_EDIT4)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_EDIT5)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_EDIT6)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_EDIT7)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC1)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC2)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC3)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC4)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC5)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC6)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC7)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC8)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_WJLJXZ)->ShowWindow(SW_SHOW);
		GetDlgItem(IDOK)->ShowWindow(SW_SHOW);
		GetDlgItem(IDCANCEL)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_RADIO1)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_RADIO2)->ShowWindow(SW_SHOW);
	}
	else
	{   //标签0的控件隐藏
		//GetDlgItem(IDC_EDIT1)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_COMBO3)->ShowWindow(SW_HIDE);
		//GetDlgItem(IDC_EDIT2)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_COMBO4)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT3)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT4)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT5)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT6)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT7)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC1)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC2)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC3)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC4)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC5)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC6)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC7)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC8)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_WJLJXZ)->ShowWindow(SW_HIDE);
		GetDlgItem(IDOK)->ShowWindow(SW_HIDE);
		GetDlgItem(IDCANCEL)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_RADIO1)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_RADIO2)->ShowWindow(SW_HIDE);
        //标签1的控件显示
		GetDlgItem(IDC_LIST1)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_COMBO1)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_BUTTONDEL)->ShowWindow(SW_SHOW);
	}
}
void CWDgldlg::OnClickList1(NMHDR* pNMHDR, LRESULT* pResult) 
{
	// TODO: Add your control notification handler code here
	UpdateData(true);
	int pos;
	pos=m_list.GetSelectionMark();
	m_combo1.SelectString(-1,m_list.GetItemText(pos,3));

	*pResult = 0;
}

void CWDgldlg::OnDblclkList1(NMHDR* pNMHDR, LRESULT* pResult) 
{
	// TODO: Add your control notification handler code here
	UpdateData(true);

	int pos;
	pos=m_list.GetSelectionMark();
	SetCurTab(0); //切换到标签0下
	//将列表中内容添加到标签0的各个控件中
	CString strdwmc,strwdlb;

	strdwmc=m_list.GetItemText(pos,0);	
    m_combo3.SelectString(-1,strdwmc);

    OnSelchangeCombo3(); //先调用执行该函数,才能正确执行下面的语句

    strwdlb=m_list.GetItemText(pos,1);	
	m_combo4.SelectString(-1,strwdlb);

	m_wdbh=atoi(m_list.GetItemText(pos,2));
	m_wdmc=m_list.GetItemText(pos,3);
	m_gjz=m_list.GetItemText(pos,4);
	m_wjlj.SetWindowText(m_list.GetItemText(pos,5));
	m_memo=m_list.GetItemText(pos,6);


	*pResult = 0;
}


void CWDgldlg::OnSelchangeCombo3() //自己添加的函数
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	CString strdwmc;
	m_combo3.GetLBText(m_combo3.GetCurSel(),strdwmc);//获得当前选中的单位名称
			
	dwb.Load_dep();
	mlb.Load_dep();
	xxb.Load_dep();
	m_combo4.ResetContent();//删除数据
	for(int i=0;i<dwb.a_DWbh.GetSize();i++) //根据单位编号在单位表中搜索单位名称
	{
		if(strdwmc == dwb.a_DWmc.GetAt(i))
		{
			for(int j=0;j<mlb.a_LBbh.GetSize();j++) //根据类别编号在类别表中搜索类别名称
				if(atoi(dwb.a_DWbh.GetAt(i))==atoi(mlb.a_DWbh.GetAt(j)))
					//自动往标签页1下的类别组合框添加文档类别 		   
					m_combo4.AddString(mlb.a_LBmc.GetAt(j));
		}
	}
}
void CWDgldlg::OnBUTTONDelWD() //删除文档的按钮响应函数
{
	// TODO: Add your control notification handler code here
		
	CString wdmc;

	if(m_combo1.GetCurSel()==CB_ERR) 
	{
		MessageBox("文档名称不能为空,请选择文档!","文档管理系统");
	    return;
	}
	else 
		m_combo1.GetLBText(m_combo1.GetCurSel(),wdmc);
	
	for(int i=0;i<xxb.a_WDbh.GetSize();i++)
	{
		if(wdmc==xxb.a_WDmc.GetAt(i))
		{
			wdbh=atoi(xxb.a_WDbh.GetAt(i));
		}
	}
	xxb.sql_delete(wdbh);
	zhi.SetDLsj(t.Format("%y-%m-%d"));
	zhi.SetName(user.GetUsername());
	zhi.SetDZ("文档删除");
	zhi.sql_insert();	
	CDialog::OnOK();

}


BOOL CWDgldlg::PreTranslateMessage(MSG* pMsg) 
{
	// TODO: Add your specialized code here and/or call the base class
	if(pMsg->message==WM_KEYDOWN && pMsg->wParam==13)
	{	pMsg->wParam=9;	
	}
	
	return CDialog::PreTranslateMessage(pMsg);
}


