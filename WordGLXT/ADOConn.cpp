// ADOConn.cpp: implementation of the ADOConn class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "WordGLXT.h"
#include "ADOConn.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

ADOConn::ADOConn()
{

}

ADOConn::~ADOConn()
{

}

void ADOConn::OnInitADOConn()
{
    //��ʼ��OLE/COM�⻷��
    ::CoInitialize(NULL);
    try {
        //����connection����
        //��������Ч�ڣ�m_pConnection.CreateInstance("ADODB.Connection");
        m_pConnection.CreateInstance(__uuidof(Connection));

        // ��������ļ���·��
        TCHAR szFilename[MAX_PATH] = { 0 };
        GetModuleFileName(NULL, szFilename, _countof(szFilename));
        PathRemoveFileSpec(szFilename);
        // ��ȡ�����ļ�
        const PCTSTR szAppName = _T("��������");
        CString strFilename = szFilename;
        strFilename += _T("\\Database.ini");

        CString strInitialCatalog, strDataSource, strUserID, strPassword;
        GetPrivateProfileString(szAppName, _T("InitialCatalog"), _T("WenDGL1"), strInitialCatalog.GetBuffer(1024), 1024, strFilename);
        GetPrivateProfileString(szAppName, _T("DataSource"), _T("127.0.0.1"), strDataSource.GetBuffer(1024), 1024, strFilename);
        GetPrivateProfileString(szAppName, _T("UserID"), _T("sa"), strUserID.GetBuffer(1024), 1024, strFilename);
        GetPrivateProfileString(szAppName, _T("Password"), _T("root"), strPassword.GetBuffer(1024), 1024, strFilename);
        strInitialCatalog.ReleaseBuffer();
        strDataSource.ReleaseBuffer();
        strUserID.ReleaseBuffer();
        strPassword.ReleaseBuffer();

        //���������ַ���
        CString str;
        str.Format(_T("Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=%s;Data Source=%s;User ID=%s;Password=%s;")
                   , strInitialCatalog.GetString()
                   , strDataSource.GetString()
                   , strUserID.GetString()
                   , strPassword.GetString()
                  );
        /*str.Format(_T("driver={SQL Server};Server=%s;Database=%s;UID=%s;PWD=%s;")
                   , strDataSource.GetString()
                   , strInitialCatalog.GetString()
                   , strUserID.GetString()
                   , strPassword.GetString()
                  );
                  */
        _bstr_t strConnect = str;

        //SERVER��UID,PWD�����ø���ʵ�����������
        m_pConnection->Open(strConnect, strUserID.GetString(), strPassword.GetString(), adModeUnknown);
    }
    //��׽�쳣
    catch(_com_error e) {
        //��ʾ������Ϣ
        AfxMessageBox(_T("�������ݿ�ʧ��!"));
        AfxMessageBox(e.Description());
    }
    catch(...) {
        AfxMessageBox(_T("�������ݿ�ʧ��!"));
    }
}

_RecordsetPtr &ADOConn::GetRecordSet(_bstr_t bstrSQL)
{
    try {
        //�������ݿ⣬���connection����Ϊ�գ��������������ݿ�
        if(m_pConnection == NULL) {
            OnInitADOConn();
        }
        //������¼������
        //m_pRecordset.CreateInstance(__uuidof(Recordset));
        m_pRecordset.CreateInstance("ADODB.Recordset");
        //ȡ�ñ��еļ�¼
        m_pRecordset->Open(bstrSQL, m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
    }
    catch(_com_error e) {
        e.Description();
    }
    //���ؼ�¼��
    return m_pRecordset;
}

BOOL ADOConn::ExecuteSQL(_bstr_t bstrSQL)
{
    _variant_t RecordsAffected;
    try {
        //�Ƿ����������ݿ�
        if(m_pConnection == NULL) {
            OnInitADOConn();
        }
        /***********************************************************************
        *  connection�����Execute����˵�����£�                               *
        *  Execute(_bstr_t CommandText,VARIANT * RecordsAffected,long Options) *
        *       ����CommandText�������ַ���,ͨ����SQL����                      *
        *       ����RecordsAffected�ǲ�����ɺ���Ӱ�������                    *
        *       ����Options��ʾCommandText�����͡�                             *
        *          adCmdText-�ı�����                                          *
        *          adCmdTable-����                                             *
        *          adCmdProc-�洢����                                          *
        *          adCmdUnknown-δ֪                                           *
        ***********************************************************************/
        m_pConnection->Execute(bstrSQL, NULL, adCmdText); //ִ��SQL���
        return true;
    }
    catch(_com_error e) {
        e.Description();
        return false;
    }
}

void ADOConn::ExitConnect()
{
    //�رռ�¼��������
    if(m_pRecordset != NULL) {
        m_pRecordset->Close();
    }
    m_pConnection->Close();
    //�ͷŻ���
    ::CoUninitialize();
}