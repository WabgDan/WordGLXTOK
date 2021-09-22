// FileAttri.cpp : implementation file
//

#include "stdafx.h"
#include "WordGLXT.h"
#include "FileAttri.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// FileAttri dialog


FileAttri::FileAttri(CWnd* pParent /*=NULL*/)
	: CDialog(FileAttri::IDD, pParent)
{
	//{{AFX_DATA_INIT(FileAttri)
	//}}AFX_DATA_INIT
	    
	
}


void FileAttri::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(FileAttri)
	DDX_Control(pDX, IDC_Pigeonhole, m_ar);
	DDX_Control(pDX, IDC_Catalog, m_dir);
	DDX_Control(pDX, IDC_VolumeLabel, m_vol);
	DDX_Control(pDX, IDC_ReadOnly, m_readonly);
	DDX_Control(pDX, IDC_System, m_system);
	DDX_Control(pDX, IDC_Normal, m_normal);
	DDX_Control(pDX, IDC_Hide, m_hide);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(FileAttri, CDialog)
	//{{AFX_MSG_MAP(FileAttri)
	ON_BN_CLICKED(ID_BNAddFile, OnBNAddFile)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// FileAttri message handlers



void FileAttri::OnBNAddFile()  //添加文件响应事件
{
	// TODO: Add your control notification handler code here
	CFileDialog log(TRUE,"文件","*.*",OFN_HIDEREADONLY,"FILE(*.*)|*.*||",NULL);
	if(log.DoModal()==IDCANCEL)return;
	pathname=log.GetPathName();	
    ShowFileAttri();
}

void FileAttri::ShowFileAttri()
{
	   
		GetDlgItem(IDC_STATIC_FilePath)->SetWindowText(pathname);//显示文件路径
	    CFileStatus status;
		CFile::GetStatus(pathname,status);
		CString strsize; //
		if(status.m_size<1024)  //显示文件大小的方式
		{
		   strsize.Format("%d",status.m_size);
		   strsize.Insert(strsize.GetLength(),'B');
		}
		else if(status.m_size<1024*1024)
		{
		   strsize.Format("%d",status.m_size/1024);
		   strsize.Insert(strsize.GetLength(),'K');
		}
		else
		{
		   strsize.Format("%d",status.m_size/(1024*1024));
		   strsize.Insert(strsize.GetLength(),'M');
		}
		
		CString str;
		str.Format("文件大小:%s",strsize);
        GetDlgItem(IDC_STATIC_FSize)->SetWindowText(str);

		CTime createtime = status.m_ctime;//文件的创建时间
		CTime modifytime = status.m_mtime;//文件的修改时间
		CTime readtime = status.m_atime;  //文件的访问时间
		GetDlgItem(IDC_STATIC_CTime)->SetWindowText(createtime.Format("创建时间:%Y年%m月%d日%H:%M:%S"));
		GetDlgItem(IDC_STATIC_MTime)->SetWindowText(modifytime.Format("修改时间:%Y年%m月%d日%H:%M:%S"));
		GetDlgItem(IDC_STATIC_RTime)->SetWindowText(readtime.Format("访问时间:%Y年%m月%d日%H:%M:%S"));
		
		if(status.m_attribute&0x00)	m_ar.SetCheck(1);
		if(status.m_attribute&0x10)	m_dir.SetCheck(1);
		if(status.m_attribute&0x02)	m_hide.SetCheck(1);
		if(status.m_attribute&0x20)	m_normal.SetCheck(1);
		if(status.m_attribute&0x01)	m_readonly.SetCheck(1);
		if(status.m_attribute&0x04)	m_system.SetCheck(1);
		if(status.m_attribute&0x08)	m_vol.SetCheck(1);
		
}



BOOL FileAttri::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	// TODO: Add extra initialization here
	if(pathname=="") ::AfxMessageBox("请先选择或打开文档！",MB_OK,0);
	else
	{   
		GetDlgItem(ID_BNAddFile)->EnableWindow(false);
		ShowFileAttri();
	}
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void FileAttri::OnCancel() 
{
	// TODO: Add extra cleanup here
	
	CDialog::OnCancel();
	GetDlgItem(ID_BNAddFile)->EnableWindow(true);
}
