// YHgldlg.cpp : implementation file
//

#include "stdafx.h"
#include "WordGLXT.h"
#include "YHgldlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CYHgldlg dialog


CYHgldlg::CYHgldlg(CWnd* pParent /*=NULL*/)
	: CDialog(CYHgldlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CYHgldlg)
	m_name = _T("");
	m_pwd = _T("");
	m_pwd1 = _T("");
	//}}AFX_DATA_INIT
	m_hIcon = AfxGetApp()->LoadIcon(IDI_ICONYHGL);
}


void CYHgldlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CYHgldlg)
	DDX_Control(pDX, IDC_COMBO1, m_jb);
	DDX_Text(pDX, IDC_EDIT1, m_name);
	DDX_Text(pDX, IDC_EDIT2, m_pwd);
	DDX_Text(pDX, IDC_EDIT3, m_pwd1);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CYHgldlg, CDialog)
	//{{AFX_MSG_MAP(CYHgldlg)
	ON_BN_CLICKED(IDC_BUTTON1, OnButton1)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CYHgldlg message handlers

BOOL CYHgldlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	// TODO: Add extra initialization here
	SetIcon(m_hIcon, TRUE);	
	m_jb.AddString("系统管理员");
	m_jb.AddString("普通管理员");
	t=CTime::GetCurrentTime();
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void CYHgldlg::OnOK() 
{
	// TODO: Add extra validation here
	UpdateData(true);
	if(m_name=="")
	{
		MessageBox("请输入用户名","文档管理系统");
		return;
	}
	CString jb;
	m_jb.GetWindowText(jb);
	if(jb=="")
	{
		MessageBox("级别不能为空","文档管理系统");
		return;
	}
    if(m_pwd=="")
	{
		MessageBox("请输入密码","文档管理系统");
		return;
	}
	if(m_pwd1=="")
	{
		MessageBox("请确认密码","文档管理系统");
		return;
	}
	if(m_pwd!=m_pwd1)
	{
		MessageBox("两次密码不同","文档管理系统");
		return;
	}
	CUsers ser;
	if(ser.Havename(m_name)==1)
	{
		MessageBox("用户以存在，请重新输入","文档管理系统");
		return;
	}
	
	ser.SetUsername(m_name);
	ser.SetPwd(m_pwd);
	m_jb.GetLBText(m_jb.GetCurSel(),strjb);
	if(strjb=="系统管理员")
		strjb="1";
	else
		strjb="0";
	ser.SetJB(strjb);
	ser.sql_insert();
	zhi.SetDLsj(t.Format("%y-%m-%d"));
	zhi.SetName(user.GetUsername());
	zhi.SetDZ("添加用户");
	zhi.sql_insert();
	CDialog::OnOK();
}

void CYHgldlg::OnButton1() 
{
	// TODO: Add your control notification handler code here
	UpdateData(true);
	if(MessageBox("是否删除当前记录","请确认",MB_YESNO)==IDYES)
	{
		CUsers ser;
		ser.sql_delete(m_name);
		zhi.SetDLsj(t.Format("%y-%m-%d"));
		zhi.SetName(user.GetUsername());
		zhi.SetDZ("用户删除");
		zhi.sql_insert();
	}
	CDialog::OnOK();
}

BOOL CYHgldlg::PreTranslateMessage(MSG* pMsg) 
{
	// TODO: Add your specialized code here and/or call the base class
	if(pMsg->message==WM_KEYDOWN && pMsg->wParam==13)
		pMsg->wParam=9;	
	return CDialog::PreTranslateMessage(pMsg);
}
