// Dwdandlg.cpp : implementation file
//

#include "stdafx.h"
#include "WordGLXT.h"
#include "Dwdandlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CDwdandlg dialog


CDwdandlg::CDwdandlg(CWnd* pParent /*=NULL*/)
	: CDialog(CDwdandlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CDwdandlg)
	m_lxdz = _T("");
	m_lxr = _T("");
	m_memo = _T("");
	m_dwbh = 0;
	RadioFlag = 0;
	m_lxdh = _T("");
	//}}AFX_DATA_INIT
	m_hIcon = AfxGetApp()->LoadIcon(IDI_ICONDANWEI);
}


void CDwdandlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CDwdandlg)
	DDX_Control(pDX, IDC_COMBO2, m_combo_dwmc);
	DDX_Control(pDX, IDC_LIST1, m_list);
	DDX_Control(pDX, IDC_TAB1, m_tab);
	DDX_Text(pDX, IDC_EDIT5, m_lxdz);
	DDX_Text(pDX, IDC_EDIT3, m_lxr);
	DDX_Text(pDX, IDC_EDIT6, m_memo);
	DDX_Text(pDX, IDC_EDIT1, m_dwbh);
	DDX_Text(pDX, IDC_EDIT4, m_lxdh);
	DDV_MaxChars(pDX, m_lxdh, 13);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CDwdandlg, CDialog)
	//{{AFX_MSG_MAP(CDwdandlg)
	ON_NOTIFY(TCN_SELCHANGE, IDC_TAB1, OnSelchangeTab1)
	ON_NOTIFY(TCN_SELCHANGING, IDC_TAB1, OnSelchangingTab1)
	ON_NOTIFY(NM_DBLCLK, IDC_LIST1, OnDblclkList1)
	ON_CBN_SELCHANGE(IDC_COMBO2, OnSelchangeCombo2)
	ON_BN_CLICKED(IDC_RADIO_Modify, OnRADIOModify)
	ON_BN_CLICKED(IDC_RADIO_ADD, OnRadioAdd)
	ON_BN_CLICKED(IDC_RADIO_Del, OnRADIODel)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CDwdandlg message handlers

void CDwdandlg::OnOK()  //添加按钮
{
	// TODO: Add extra validation here
	switch(RadioFlag)
	{
	case 1:
		   ModifyDW();
		   break;
	case 2:
		   AddDW();
		   break;
	case 3: 
			DelDW();
		    break;
	default:
			AfxMessageBox("请选择操作!");
		    return ;
	}

	CDialog::OnOK();
}
BOOL CDwdandlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	// TODO: Add extra initialization here
	SetIcon(m_hIcon, TRUE);	
	t=CTime::GetCurrentTime();//获取当前时间
	TC_ITEM tci; //一个标签可能含有文本/图片等属性，
	             //TC_ITEM结构的mask指定了标签的哪种属性是有效的。
	             //mask=TCIF_IMAGE，图片有效
	tci.mask=TCIF_TEXT;    //文本有效
	tci.pszText="单位档案";
	m_tab.InsertItem(0,&tci);
	tci.pszText="单位列表";
	m_tab.InsertItem(1,&tci);
	
	GetDlgItem(IDC_LIST1)->ShowWindow(SW_HIDE);//使ListContro控件不可见

	UpdateData(true);

	//为ListContro控件设置列
	m_list.SetExtendedStyle(LVS_EX_FLATSB|LVS_EX_FULLROWSELECT|LVS_EX_GRIDLINES);
	m_list.InsertColumn(0,"单位编号",LVCFMT_LEFT,100,0);
	m_list.InsertColumn(1,"单位名称",LVCFMT_LEFT,100,1);
	m_list.InsertColumn(2,"联系人",LVCFMT_LEFT,100,2);
	m_list.InsertColumn(3,"联系电话",LVCFMT_LEFT,100,3);
	m_list.InsertColumn(4,"联系地址",LVCFMT_LEFT,100,4);
	m_list.InsertColumn(6,"备注",LVCFMT_LEFT,100,5);

	ADOConn m_AdoConn;
	m_AdoConn.OnInitADOConn();//连接数据库

	CString sql;
	sql.Format("select* from Dwxxb order by dwbh desc");

	m_AdoConn.GetRecordSet((_bstr_t)sql);

	while(m_AdoConn.m_pRecordset->adoEOF==0)
	{
		m_list.InsertItem(0,"");
		m_list.SetItemText(0,0,(char*)(_bstr_t)m_AdoConn.m_pRecordset->GetCollect("dwbh"));
		m_list.SetItemText(0,1,(char*)(_bstr_t)m_AdoConn.m_pRecordset->GetCollect("dwmc"));
		m_combo_dwmc.AddString((char*)(_bstr_t)m_AdoConn.m_pRecordset->GetCollect("dwmc"));
		m_list.SetItemText(0,2,(char*)(_bstr_t)m_AdoConn.m_pRecordset->GetCollect("lxr"));
		m_list.SetItemText(0,3,(char*)(_bstr_t)m_AdoConn.m_pRecordset->GetCollect("lxdh"));
		m_list.SetItemText(0,4,(char*)(_bstr_t)m_AdoConn.m_pRecordset->GetCollect("lxdz"));
		m_list.SetItemText(0,5,(char*)(_bstr_t)m_AdoConn.m_pRecordset->GetCollect("memo"));
		m_AdoConn.m_pRecordset->MoveNext();
	}
	m_AdoConn.ExitConnect();//断开连接

	CDwxxb dwb;
	dwb.Load_dep();
    m_dwbh=	dwb.a_DWbh.GetSize()+1;;//设置单位档案中默认单位编号

	UpdateData(false);
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void CDwdandlg::OnSelchangeTab1(NMHDR* pNMHDR, LRESULT* pResult) //在用户改变标签时产生
{
	// TODO: Add your control notification handler code here
	switch(m_tab.GetCurSel())//实现标签的切换功能
	{
	case 0: 
		GetDlgItem(IDC_EDIT1)->ShowWindow(SW_SHOW);
        GetDlgItem(IDC_COMBO2)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_EDIT3)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_EDIT4)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_EDIT5)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_EDIT6)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC1)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC2)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC3)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC4)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC5)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC6)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC7)->ShowWindow(SW_SHOW);
		GetDlgItem(IDOK)->ShowWindow(SW_SHOW);
		GetDlgItem(IDCANCEL)->ShowWindow(SW_SHOW);

		break;
	case 1: 
		GetDlgItem(IDC_LIST1)->ShowWindow(SW_SHOW);
		break;
	}
	*pResult = 0;
}

void CDwdandlg::OnSelchangingTab1(NMHDR* pNMHDR, LRESULT* pResult) 
{
	// TODO: Add your control notification handler code here
	switch(m_tab.GetCurSel())
	{
	case 0: 
		GetDlgItem(IDC_EDIT1)->ShowWindow(SW_HIDE);	
		GetDlgItem(IDC_COMBO2)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT3)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT4)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT5)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT6)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC1)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC2)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC3)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC4)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC5)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC6)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC7)->ShowWindow(SW_HIDE);
		GetDlgItem(IDOK)->ShowWindow(SW_HIDE);
		GetDlgItem(IDCANCEL)->ShowWindow(SW_HIDE);

		break;
	case 1: 
		GetDlgItem(IDC_LIST1)->ShowWindow(SW_HIDE);
		
		break;

	}
	*pResult = 0;
}
void CDwdandlg::SetCurTab(UINT m_index)
{
	m_tab.SetCurSel(m_index);
	if(m_index==0)
	{
		GetDlgItem(IDC_LIST1)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT1)->ShowWindow(SW_SHOW);
		
        GetDlgItem(IDC_COMBO2)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_EDIT3)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_EDIT4)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_EDIT5)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_EDIT6)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC1)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC2)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC3)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC4)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC5)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC6)->ShowWindow(SW_SHOW);
		GetDlgItem(IDC_STATIC7)->ShowWindow(SW_SHOW);
		GetDlgItem(IDOK)->ShowWindow(SW_SHOW);
		GetDlgItem(IDCANCEL)->ShowWindow(SW_SHOW);

	}
	else 
	{
		GetDlgItem(IDC_EDIT1)->ShowWindow(SW_HIDE);
		
		GetDlgItem(IDC_COMBO2)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT3)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT4)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT5)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_EDIT6)->ShowWindow(SW_HIDE);

		GetDlgItem(IDC_STATIC1)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC2)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC3)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC4)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC5)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC6)->ShowWindow(SW_HIDE);
		GetDlgItem(IDC_STATIC7)->ShowWindow(SW_HIDE);
		GetDlgItem(IDOK)->ShowWindow(SW_HIDE);
		GetDlgItem(IDCANCEL)->ShowWindow(SW_HIDE);
		
		GetDlgItem(IDC_LIST1)->ShowWindow(SW_SHOW);
	}
	
}

void CDwdandlg::OnDblclkList1(NMHDR* pNMHDR, LRESULT* pResult) //双击列表的响应函数
{
	// TODO: Add your control notification handler code here
	UpdateData(true);
	int pos;
	pos=m_list.GetSelectionMark();
	SetCurTab(0);
	m_dwbh=atoi(m_list.GetItemText(pos,0));

	CString Dbldwmc = m_list.GetItemText(pos,1);

	m_combo_dwmc.SelectString(-1,Dbldwmc);

	m_lxr=m_list.GetItemText(pos,2);
	m_lxdh=m_list.GetItemText(pos,3);
	m_lxdz=m_list.GetItemText(pos,4);
	m_memo=m_list.GetItemText(pos,5);
	UpdateData(false);
	*pResult = 0;
}

BOOL CDwdandlg::PreTranslateMessage(MSG* pMsg) 
{
	// TODO: Add your specialized code here and/or call the base class
	if(pMsg->message==WM_KEYDOWN && pMsg->wParam==13)
		pMsg->wParam=9;	
	return CDialog::PreTranslateMessage(pMsg);
}

void CDwdandlg::OnSelchangeCombo2() 
{
	// TODO: Add your control notification handler code here
	 GetDlgItem(IDC_EDIT1)->EnableWindow(false);
	 CDwxxb dwb;	
	CString strdwmc;

	m_combo_dwmc.GetLBText(m_combo_dwmc.GetCurSel(),strdwmc);    
   
	dwb.Load_dep();
    int m =dwb.a_DWbh.GetSize();
	for(int i=0;i<m;i++)//根据单位编号搜索单位名称
	{
		
		if(strdwmc==dwb.a_DWmc.GetAt(i))
		{
			m_dwbh = atoi(dwb.a_DWbh.GetAt(i));
		}
	}
	UpdateData(false);
}

void CDwdandlg::OnRADIOModify() 
{
	// TODO: Add your control notification handler code here
	GetDlgItem(IDC_EDIT1)->EnableWindow(TRUE);
	RadioFlag = 1;
}

void CDwdandlg::OnRadioAdd() 
{
	// TODO: Add your control notification handler code here
    CDwxxb dwb;
	dwb.Load_dep();
    m_dwbh=	dwb.a_DWbh.GetSize()+1;;//重新显示默认的单位编号
    GetDlgItem(IDC_EDIT1)->EnableWindow(FALSE);
	RadioFlag = 2;
	UpdateData(false);
}

void CDwdandlg::OnRADIODel() 
{
	// TODO: Add your control notification handler code here   	
	RadioFlag = 3;
}

void CDwdandlg::AddDW()  //添加单位信息
{
	UpdateData(true);
	CString strdwmc; 	
    m_combo_dwmc.GetWindowText(strdwmc);
	CDwxxb dwb;	
	if(strdwmc=="")
		::AfxMessageBox("单位名称不能为空");
	
	else if(dwb.HaveId(m_dwbh)==1)	
		MessageBox("单位编号已存在","文档管理系统");			
	else if(m_lxdh.GetLength()>12)	
		MessageBox("电话号码不正确","文档管理系统");
	else
	{		
		dwb.SetDWbh(m_dwbh);
		dwb.SetDWmc(strdwmc);
		dwb.SetLxr(m_lxr);
		dwb.SetLxdh(m_lxdh);
		dwb.SetLxdz(m_lxdz);
		dwb.SetMemo(m_memo);
		dwb.sql_insert();
		zhi.SetDLsj(t.Format("%y-%m-%d"));
		zhi.SetName(user.GetUsername());
		zhi.SetDZ("单位添加");
		zhi.sql_insert();
	}
}

void CDwdandlg::ModifyDW() //修改单位信息
{
	UpdateData(true);
	CString strdwmc;	
	if(m_combo_dwmc.GetCurSel()!=CB_ERR)//若从下拉列表中选择 
	    	m_combo_dwmc.GetLBText(m_combo_dwmc.GetCurSel(),strdwmc);
	else	 //若没有从下拉列表中选择
	{
		m_combo_dwmc.GetWindowText(strdwmc); 
        int a =m_combo_dwmc.SelectString(-1,strdwmc);
		if(a == CB_ERR) strdwmc="";  //若数据库中没有该单位,则清空strdwmc		
	}

	if(strdwmc=="")
		MessageBox("单位名称不能为空","文档管理系统");	
	else if(m_lxdh.GetLength()>12)
		MessageBox("电话号码不正确","文档管理系统");
	else
	{
		CDwxxb dwb;
		dwb.SetDWmc(strdwmc);
		dwb.SetLxr(m_lxr);
		dwb.SetLxdh(m_lxdh);
		dwb.SetLxdz(m_lxdz);
		dwb.SetMemo(m_memo);
		dwb.sql_update(m_dwbh);
		zhi.SetDLsj(t.Format("%y-%m-%d"));
		zhi.SetName(user.GetUsername());
		zhi.SetDZ("单位修改");
		zhi.sql_insert();
	}

}

void CDwdandlg::DelDW() //删除单位信息
{
	CDwxxb dwb;
	if(dwb.HaveId(m_dwbh)==-1)
		MessageBox("单位编号不存在,无法执行删除操作!","文档管理系统");
	else
	{
		
		dwb.sql_delete(m_dwbh); //删除单位表中该单位信息
		CZdmlb mlb;
		mlb.sql_deletedw(m_dwbh); //删除类别表中该单位信息
		CZdxxb xxb;
		xxb.sql_deletedw(m_dwbh); //删除文档表中该单位信息
		zhi.SetDLsj(t.Format("%y-%m-%d"));
		zhi.SetName(user.GetUsername());
		zhi.SetDZ("单位删除");
		zhi.sql_insert();
	}	

}


