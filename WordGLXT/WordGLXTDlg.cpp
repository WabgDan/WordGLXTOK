// WordGLXTDlg.cpp : implementation file
//

#include "stdafx.h"
#include "WordGLXT.h"
#include "WordGLXTDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About
CToolBar m_ToolBar;
CString strShow;        //提示文本显示
CImageList m_ImageList; //图片列表
char buf[512];           //显示当前目录

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	//{{AFX_MSG_MAP(CAboutDlg)
		// No message handlers
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CWordGLXTDlg dialog

CWordGLXTDlg::CWordGLXTDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CWordGLXTDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CWordGLXTDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDI_CHILDICON4);
}

void CWordGLXTDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CWordGLXTDlg)
	DDX_Control(pDX, IDC_RICHEDIT1, m_richedit);
	DDX_Control(pDX, IDC_TREE1, m_tree);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CWordGLXTDlg, CDialog)
	//{{AFX_MSG_MAP(CWordGLXTDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_NOTIFY(NM_DBLCLK, IDC_TREE1, OnDblclkTree1)
	ON_NOTIFY_EX(TTN_NEEDTEXT,0, OnToolTipNotify)
	ON_COMMAND(ID_MENUDWDAN, OnMenudwdan)
	ON_COMMAND(ID_MENUWDLB, OnMenuwdlb)	
	ON_COMMAND(ID_MENULIULWD, OnMenuliulwd)
	ON_COMMAND(ID_MENUYHGL, OnMenuyhgl)
	ON_COMMAND(ID_MENUMODPWD, OnMenumodpwd)
	ON_COMMAND(ID_MENUADDWD, OnMenuaddwd)
	ON_COMMAND(ID_MENUDELWD, OnMenudelwd)
	ON_COMMAND(ID_MENUMODWD, OnMenumodwd)
	ON_COMMAND(ID_MENURZGL, OnMenurzgl)
    ON_COMMAND(IDD_MENULookFileAttri, OnMENULookFileAttri)
    ON_COMMAND(ID_EXIT, OnExit)
	ON_COMMAND(ID_Button_Add, OnButtonAdd)
	ON_COMMAND(ID_Button_DwDA, OnButtonDwDA)
	ON_COMMAND(ID_Button_WDLB, OnButtonWDLB)
	ON_COMMAND(ID_Button_Modify, OnButtonModify)
	ON_COMMAND(ID_Button_Del, OnButtonDel)
	ON_COMMAND(ID_Button_Scan, OnButtonScan)
	ON_COMMAND(ID_Button_Attri, OnButtonAttri)
	ON_COMMAND(ID_Button_UserManage, OnButtonUserManage)
	ON_COMMAND(ID_Button_mimaModify, OnButtonmimaModify)
	ON_COMMAND(ID_Button_LogManage, OnButtonLogManage)
	ON_COMMAND(ID_Button_Exit, OnButtonExit)
	ON_COMMAND(ID_Menu_DBBackUp, OnMenuDBBackUp)
	ON_COMMAND(ID_Menu_DBRestore, OnMenuDBRestore)
	ON_COMMAND(ID_Button_BackUp, OnButtonBackUp)
	ON_COMMAND(ID_Button_Restore, OnButtonRestore)
	ON_COMMAND(ID_System_Exit, OnSystemExit)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CWordGLXTDlg message handlers

BOOL CWordGLXTDlg::OnInitDialog()
{

     CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	// TODO: Add extra initialization here
	CDialogin gin;
	
	if(gin.DoModal()!=IDOK)//启动登录对话框
		   OnOK();
	dwb.Load_dep();//批量加载单位表中的记录数据
	mlb.Load_dep();//批量加载类别表中的记录数据
	xxb.Load_dep();//批量加载文档表中的记录数据
	m_treeImageList.Create(16,16,ILC_MASK,4,1);
	m_treeImageList.Add(theApp.LoadIcon(IDI_ROOTICON)); //显示根目录图标
	m_treeImageList.Add(theApp.LoadIcon(IDI_CHILDICON1));//显示一级目录的图标
	m_treeImageList.Add(theApp.LoadIcon(IDI_CHILDICON2));//显示二级目录的图标
	m_treeImageList.Add(theApp.LoadIcon(IDI_CHILDICON4));//显示三级目录的图标，就是那个word的图标
	m_tree.SetImageList(&m_treeImageList,LVSIL_NORMAL);
	m_root=m_tree.InsertItem("基本信息管理",0,0);
	AddtoTree(m_root);
	m_tree.Expand(m_root,TVE_EXPAND);
    //状态栏显示内容的设置
	m_StatusBar.EnableAutomation();
	m_StatusBar.Create(WS_CHILD|WS_VISIBLE,CRect(0,0,0,0),this,0);
	
	int width[]={200,300,600};
	m_StatusBar.SetParts(4, &width[0]);
	m_StatusBar.SetText("索尼精密",0,0);//显示单位名称
    
	CString StatusText;
	StatusText.Format("当前用户：%s",user.GetUsername());//显示当前用户
    m_StatusBar.SetText(StatusText,0,1);
	
	t=CTime::GetCurrentTime();
	CString strdate;
	strdate.Format("当前日期:%s",t.Format("%y-%m-%d %X"));//显示当前时间
	m_StatusBar.SetText(strdate,0,2);

    //工具栏显示内容的设置
	m_ImageList.Create(32,32,ILC_COLOR|ILC_MASK,1,1);  //创建图像列表
	
    m_ImageList.Add(AfxGetApp()->LoadIcon(IDI_ICONDWDA));  //单位档案
	m_ImageList.Add(AfxGetApp()->LoadIcon(IDI_ICONWDLB));    //文档类别
	
	m_ImageList.Add(AfxGetApp()->LoadIcon(IDI_ICONAdd));   //添加文档
	m_ImageList.Add(AfxGetApp()->LoadIcon(IDI_ICONMod));  //修改文档
	m_ImageList.Add(AfxGetApp()->LoadIcon(IDI_ICONDel));   //删除文档
	m_ImageList.Add(AfxGetApp()->LoadIcon(IDI_ICONScan));   //浏览文档
	m_ImageList.Add(AfxGetApp()->LoadIcon(IDI_ICONFileAttri));//查看属性
	m_ImageList.Add(AfxGetApp()->LoadIcon(IDI_ICONUser));    //用户管理
	m_ImageList.Add(AfxGetApp()->LoadIcon(IDI_ICONMIMA));    //口令修改
	m_ImageList.Add(AfxGetApp()->LoadIcon(IDI_ICONLog));       //日志管理
	m_ImageList.Add(AfxGetApp()->LoadIcon(IDI_ICONSJKBF));      //备份
	m_ImageList.Add(AfxGetApp()->LoadIcon(IDI_ICONSJKHF));      //恢复
	m_ImageList.Add(AfxGetApp()->LoadIcon(IDI_ICONExit));    //退出系统


	UINT array[16];
	for(int i=0;i<16;i++)
	{
		if(i==2||i==8||i==12)
		{
			array[i]=ID_SEPARATOR;//第三个和第九个按钮为分隔条
		}
		else  array[i]=i+1101;
	}
	
	m_ToolBar.Create(this);
	m_ToolBar.SetButtons(array,16);
	m_ToolBar.SetButtonText(0,"单位档案");
	m_ToolBar.SetButtonText(1,"文档类别");
	m_ToolBar.SetButtonText(3,"添加文档");
	m_ToolBar.SetButtonText(4,"修改文档");
	m_ToolBar.SetButtonText(5,"删除文档");
	m_ToolBar.SetButtonText(6,"浏览文档");
	m_ToolBar.SetButtonText(7,"查看属性");
	m_ToolBar.SetButtonText(9,"用户管理");
	m_ToolBar.SetButtonText(10,"口令修改");
	m_ToolBar.SetButtonText(11,"日志管理");
    m_ToolBar.SetButtonText(13,"备份数据");
    m_ToolBar.SetButtonText(14,"恢复数据");
	m_ToolBar.SetButtonText(15,"退出系统");
    m_ToolBar.GetToolBarCtrl().SetImageList(&m_ImageList);//关联图像列表
	
    m_ToolBar.SetSizes(CSize(60,60),CSize(32,32));//设置按钮和按钮位图大小
	m_ToolBar.EnableToolTips(true);
	RepositionBars(AFX_IDW_CONTROLBAR_FIRST, AFX_IDW_CONTROLBAR_LAST, 0);//show Toolbox
    ::GetCurrentDirectory(512,buf);//Obtain current path,which is used when backing up the database  
	return TRUE; 
	// return TRUE  unless you set the focus to a control
}


void CWordGLXTDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CWordGLXTDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CWordGLXTDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

void CWordGLXTDlg::AddtoTree(HTREEITEM m_node) //Define the AddtoTree function to add in each table to the tree control according hierarchy                                              
{
	int i,j;
	for(i=0;i<dwb.a_DWbh.GetSize();i++)
	{
		arrays[i]=m_tree.InsertItem(dwb.a_DWmc.GetAt(i),1,1,m_node);
		for(j=0;j<mlb.a_DWbh.GetSize();j++)
		{
			if(atoi(dwb.a_DWbh.GetAt(i))==atoi(mlb.a_DWbh.GetAt(j)))
			{
				brrays[j]=m_tree.InsertItem(mlb.a_LBmc.GetAt(j),2,2,arrays[i]);
			}
		}
	}
	for(i=0;i<xxb.a_WDbh.GetSize();i++)
	{
		for(j=0;j<mlb.a_DWbh.GetSize();j++)
		{
			if(atoi(xxb.a_DWbh.GetAt(i))==atoi(mlb.a_DWbh.GetAt(j))&&atoi(xxb.a_LBbh.GetAt(i))==atoi(mlb.a_LBbh.GetAt(j)))
			{
				CString str = xxb.a_WDmc.GetAt(i);
				hitem[i]=m_tree.InsertItem(xxb.a_WDmc.GetAt(i),3,3,brrays[j]);
			}
		}
	}
	m_tree.SetRedraw();//重画树型控件
}

void CWordGLXTDlg::OnMenudwdan() //基本信息-〉单位档案
{
	// TODO: Add your command handler code here
	CDwdandlg dlg;
	if(dlg.DoModal()==IDOK)
	{
		m_tree.DeleteAllItems();
		dwb.Load_dep();//批量加载单位表中的记录数据
		mlb.Load_dep();//批量加载类别表中的记录数据
		xxb.Load_dep();//批量加载文档表中的记录数据
		m_root=m_tree.InsertItem("基本信息管理",0,0);
		AddtoTree(m_root);
	}
}

void CWordGLXTDlg::OnMenuwdlb() //基本信息-〉文档类别
{
	// TODO: Add your command handler code here
	CWdlbiedlg dlg;
	if(dlg.DoModal()==IDOK)
	{
		m_tree.DeleteAllItems();
		dwb.Load_dep();//批量加载单位表中的记录数据
		mlb.Load_dep();//批量加载类别表中的记录数据
		xxb.Load_dep();//批量加载文档表中的记录数据
		m_root=m_tree.InsertItem("基本信息管理",0,0);
		AddtoTree(m_root);
	}
}

void CWordGLXTDlg::OnDblclkTree1(NMHDR* pNMHDR, LRESULT* pResult) //树控件的双击事件
{
	// TODO: Add your control notification handler code here
	CString strWjian="";

	//读取当前节点
	temp = m_tree.GetSelectedItem();//获得当前选定树控件的项目
	//将当前节点子节点付给temp
	temp = m_tree.GetChildItem(temp);//获取指定树控件的子项目
	if (temp != NULL)
	{
		while (temp!= NULL)
		{
			//取出temp中的文本
			CString strTemp = m_tree.GetItemText(temp);  // 
			strWjian += strTemp + "\n";			
			m_richedit.SetWindowText(strWjian);       //RichEdit控件显示数据
			temp = m_tree.GetNextItem(temp,TVGN_NEXT);//将temp的兄弟节点付给temp
		}
	}
	else
	{
		temp = m_tree.GetSelectedItem(); //重新获得当前选定树控件的项目
		for(int i=0;i<xxb.a_WDbh.GetSize();i++)
		{
			if(temp==hitem[i])
			{
				//取出temp对应的文档路径
				strWordpath = xxb.a_WJlj.GetAt(i);
				CFileFind file;
				if(!file.FindFile(strWordpath))
				{//查找文件是否存在，不存在则清除数据库中的记录
					MessageBox("文件不存在！","文档管理系统");
					
					int	wdbh=0;
					wdbh = atoi(xxb.a_WDbh.GetAt(i));
					//删除该文档
					UpdateData(true);
					xxb.sql_delete(wdbh); //删除该文档
					MessageBox("数据库中该文件的记录已删除！","文档管理系统");
					UpdateData(false);
					return ;
				}

			}
		}
		//word应用程序的调用 
		_Application app;
		//初始化连接
		//***解决初始化连接时的进程冲突问题*****************
        LPDISPATCH  pDisp;
		LPUNKNOWN  pUnk;
		CLSID clsid;
		CLSIDFromProgID(L"word.Application",&clsid);
        if(GetActiveObject(clsid,NULL,&pUnk)==S_OK)
		{
			pUnk->QueryInterface(IID_IDispatch,(void**)&pDisp);
			app.AttachDispatch(pDisp);
			
		}
		else
		if(!app.CreateDispatch("word.Application"))   // 启动word
		{ 
			MessageBox("Word启动失败！","文档管理系统"); 
			return ;
		}
		//***解决初始化连接时的进程冲突问题*****************

		Documents doc;
		CComVariant a (_T(strWordpath)),b(false),c(0),d(true),aa(0),bb(1);
		_Document doc1;
	     
		doc.AttachDispatch( app.GetDocuments());
		doc1.AttachDispatch(doc.Add(&a,&b,&c,&d));
		Range range;
		
		//取出文档的所选区域
		range = doc1.GetContent();
		//取出文件内容
		strText = range.GetText();
		m_richedit.SetWindowText(strText);
		//关闭
		app.Quit(&b,&c,&c);	
		//释放环境
		range.ReleaseDispatch();
		doc.ReleaseDispatch();
		doc1.ReleaseDispatch();
		app.ReleaseDispatch();

	}
	*pResult = 0;
}

void CWordGLXTDlg::OnMenuliulwd() //文档管理-〉文档浏览
{
	// TODO: Add your command handler code here
	CString strd,strs;
	xxb.Load_dep();
	for(int i=0;i<xxb.a_WDbh.GetSize();i++)
	{
		strd=xxb.a_WDmc.GetAt(i);
		strs+=strd+"\n";
		m_richedit.SetWindowText(strs);
	}
}

void CWordGLXTDlg::OnMenuyhgl() //用户管理功能
{
	// TODO: Add your command handler code here
	if(user.GetJB()=="1")
	{
		CYHgldlg dlg;
		dlg.DoModal();
	}
	else
	{
		MessageBox("没有权限","文档管理系统");
		return;
	}	
}

void CWordGLXTDlg::OnMenumodpwd() //口令修改功能
{
	// TODO: Add your command handler code here
	CKLxgdlg dlg;
	dlg.DoModal();
}

void CWordGLXTDlg::OnMenuaddwd() //实现添加文档功能
{
	// TODO: Add your command handler code here
	CWDgldlg dlg;
	dlg.str = 0;
	if(dlg.DoModal()==IDOK)
	{
		m_tree.DeleteAllItems();
		dwb.Load_dep();
		mlb.Load_dep();
		xxb.Load_dep();
		m_root=m_tree.InsertItem("基本信息管理",0,0);
		AddtoTree(m_root);
	}
}

void CWordGLXTDlg::OnMenudelwd() //实现删除文档功能
{
	// TODO: Add your command handler code here
	CWDgldlg dlg;
	dlg.tabindex = 1;
	if(dlg.DoModal()==IDOK)
	{
		m_tree.DeleteAllItems();
		dwb.Load_dep();
		mlb.Load_dep();
		xxb.Load_dep();
		m_root=m_tree.InsertItem("基本信息管理",0,0);
		AddtoTree(m_root);
	}
}

void CWordGLXTDlg::OnMenumodwd() //实现修改文档功能
{
	// TODO: Add your command handler code here
	CWDgldlg dlg;
	dlg.str = 1;
	if(dlg.DoModal()==IDOK)
	{
		m_tree.DeleteAllItems();
		dwb.Load_dep();
		mlb.Load_dep();
		xxb.Load_dep();
		m_root=m_tree.InsertItem("基本信息管理",0,0);
		AddtoTree(m_root);
	}
}

void CWordGLXTDlg::OnMenurzgl() //实现日志管理功能
{
	// TODO: Add your command handler code here
	ADOConn m_AdoConn;
    m_AdoConn.OnInitADOConn();
	CString sql,sqlzd="用户名\t登录时间\t动作\n";
	sql.Format("select* from Rizhib");
	m_AdoConn.GetRecordSet((_bstr_t)sql);
	while(m_AdoConn.m_pRecordset->adoEOF==0)
	{
		sqlzd+=(char*)(_bstr_t)m_AdoConn.m_pRecordset->GetCollect("name");
		sqlzd+="  \t";
		sqlzd+=(char*)(_bstr_t)m_AdoConn.m_pRecordset->GetCollect("DLsj");
		sqlzd+="\t";
		sqlzd+=(char*)(_bstr_t)m_AdoConn.m_pRecordset->GetCollect("dz");
		sqlzd+="\n";
		m_AdoConn.m_pRecordset->MoveNext();
		m_richedit.SetWindowText(sqlzd);
	}
	m_AdoConn.ExitConnect();
}

void CWordGLXTDlg::OnExit() 
{
	// TODO: Add your command handler code here
	OnOK();
}



BOOL CWordGLXTDlg::PreTranslateMessage(MSG* pMsg) 
{
	// TODO: Add your specialized code here and/or call the base class
	if(pMsg->message==WM_KEYDOWN && pMsg->wParam==13)
		pMsg->wParam=9;		
	return CDialog::PreTranslateMessage(pMsg);
}

void CWordGLXTDlg::OnMENULookFileAttri() //察看文件属性功能
{
	// TODO: Add your command handler code here
	FileAttri dlg;
	if(strWordpath!="") 
	{
		dlg.pathname = strWordpath;	
		
		if(dlg.DoModal()==IDCANCEL) return ;
	}
	else dlg.DoModal();   
}

//工具栏按钮响应函数
void CWordGLXTDlg::OnButtonDwDA() 
{
	OnMenudwdan();
	// TODO: Add your command handler code here
	
}
void CWordGLXTDlg::OnButtonWDLB() 
{
	// TODO: Add your command handler code here
	OnMenuwdlb();
}

void CWordGLXTDlg::OnButtonAdd() 
{
	// TODO: Add your command handler code here
	OnMenuaddwd();
}

 

void CWordGLXTDlg::OnButtonModify() 
{
	// TODO: Add your command handler code here
	OnMenumodwd();
}

void CWordGLXTDlg::OnButtonDel() 
{
	// TODO: Add your command handler code here
	
	OnMenudelwd();
}

void CWordGLXTDlg::OnButtonScan() 
{
	// TODO: Add your command handler code here
	OnMenuliulwd();
}

void CWordGLXTDlg::OnButtonAttri() 
{
	// TODO: Add your command handler code here
	OnMENULookFileAttri();
}


void CWordGLXTDlg::OnButtonUserManage() 
{
	// TODO: Add your command handler code here
	OnMenuyhgl();
}


void CWordGLXTDlg::OnButtonmimaModify() 
{
	// TODO: Add your command handler code here
	OnMenumodpwd();
	
}

void CWordGLXTDlg::OnButtonLogManage() 
{
	// TODO: Add your command handler code here
	OnMenurzgl();
}



void CWordGLXTDlg::OnButtonExit() 
{
	// TODO: Add your command handler code here
  
	OnOK();
}



void CWordGLXTDlg::OnMenuDBBackUp() 
{
	// TODO: Add your command handler code here
	ADOConn m_AdoConn;

	CString sql,strpath,m_filename;
	strcat(buf,"\\DataBase\\WDGL_Data.DAT");
	strpath.Format("%s",buf);
	m_filename="WDGL_Data";
	sql = "use master exec sp_addumpdevice 'disk','"+m_filename+"','"+strpath+"' backup database WenDGL to "+m_filename+" ";
	if(AfxMessageBox("您确定要备份数据库吗?",MB_OKCANCEL)==IDCANCEL)  	return; 
	   	
	m_AdoConn.ExecuteSQL((_bstr_t)sql);
	m_AdoConn.ExitConnect();
	MessageBox("数据库备份成功！","系统提示",MB_OK|MB_ICONEXCLAMATION);
		
}

void CWordGLXTDlg::OnMenuDBRestore() 
{
	// TODO: Add your command handler code here
	ADOConn m_AdoConn;
	CString sql,m_filename;
	m_filename="WDGL_Data";
	sql = "use master  restore database WenDGL from "+m_filename+" ";
	if(AfxMessageBox("还原数据库将覆盖原来的数据库。您确定要还原吗?",MB_OKCANCEL)==IDCANCEL) 
	return; 
	m_AdoConn.ExecuteSQL((_bstr_t)sql);
	m_AdoConn.ExitConnect();
	MessageBox("数据库恢复完成！","系统提示",MB_OK|MB_ICONEXCLAMATION); 
}

void CWordGLXTDlg::OnButtonBackUp() 
{
	// TODO: Add your command handler code here
	OnMenuDBBackUp();
}

void CWordGLXTDlg::OnButtonRestore() 
{
	// TODO: Add your command handler code here
	OnMenuDBRestore();
}

void CWordGLXTDlg::OnSystemExit() 
{
	// TODO: Add your command handler code here
	OnOK();
}

BOOL CWordGLXTDlg::OnToolTipNotify(UINT id, NMHDR *pNMHDR, LRESULT *pResult)
{
	TOOLTIPTEXT *pTTT=(TOOLTIPTEXT*)pNMHDR;	 
	UINT nID=pNMHDR->idFrom; //获取工具栏按钮ID
	 
	 if(nID)
	 {
		 nID = m_ToolBar.CommandToIndex(nID);//根据ID获取按钮索引
		 if(nID!=-1)
		 {
			 
			 m_ToolBar.GetButtonText(nID,strShow); //获取工具栏文本
		     pTTT->lpszText=strShow.GetBuffer(strShow.GetLength());//设置提示信息文本
             pTTT->hinst = AfxGetResourceHandle();
			 return(TRUE);
		 }
	 }
     return FALSE;
}
