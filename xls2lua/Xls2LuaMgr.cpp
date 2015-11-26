#include "StdAfx.h"
#include "Xls2LuaMgr.h"
//#include <windows.h>

#include <afxdb.h> 
#include <odbcinst.h>
#include <string>

using namespace std;

CXls2LuaMgr::CXls2LuaMgr(void)
{
}


CXls2LuaMgr::~CXls2LuaMgr(void)
{
}

void CXls2LuaMgr::LoadExcels()
{
	ReadFromExcel();
}

void CXls2LuaMgr::ReadFromExcel() 
{
	CDatabase database;
	string sSql;
	string sItem1, sItem2;
	string sDriver;
	string sDsn;
	string sFile = "npc.xls"; // ������ȡ��Excel�ļ���
	char buf[2048];
	TCHAR tBuf[1024];

	// �����Ƿ�װ��Excel���� "Microsoft Excel Driver (*.xls)" 
	sDriver = GetExcelDriver();
	if (sDriver.size() <= 0)
	{
		// û�з���Excel����
		//MessageBox(NULL,"��ʾ","����Exel����",MB_OK);
		return;
	}

	// �������д�ȡ���ַ���
	sprintf(buf,"ODBC;DRIVER={%s};DSN='';DBQ=%s", sDriver.c_str(), sFile.c_str());
	
	
	MultiByteToWideChar(CP_ACP,0,buf,strlen(buf),tBuf,strlen(buf) + 1);
	
	TRY
	{
		// �����ݿ�(��Excel�ļ�)
		database.Open(NULL, false, false, tBuf);

		CRecordset recset(&database);

		// ���ö�ȡ�Ĳ�ѯ���.
		//sSql = "SELECT Name, Age " 
		//	"FROM demo " 
		//	"ORDER BY Name ";

		// ִ�в�ѯ���
		//recset.Open(CRecordset::forwardOnly, sSql, CRecordset::readOnly);

		// ��ȡ��ѯ���
		int nRecord = recset.GetRecordCount();
		for (int i = 0; i < nRecord; i++)
		{
			//CString strTemp;
			//recset.GetFieldValue(i,strTemp);
			//printf("%s ",strTemp.GetString());
		}
		//while (!recset.IsEOF())
		//{
		//	//��ȡExcel�ڲ���ֵ
		//	recset.GetFieldValue("Name ", sItem1);
		//	recset.GetFieldValue("Age", sItem2);

		//	// �Ƶ���һ��
		//	recset.MoveNext();
		//}

		// �ر����ݿ�
		database.Close();

	}
	CATCH(CDBException, e)
	{
		// ���ݿ���������쳣ʱ...
		//MessageBox(NULL,"��ʾ","��ʾ",MB_OK);
	}
	END_CATCH;
}


// ��ȡODBC��Excel����
string CXls2LuaMgr::GetExcelDriver()
{
	TCHAR szBuf[2001];
	WORD cbBufMax = 2000;
	WORD cbBufOut;
	char *pszBuf = szBuf;
	string sDriver;

	// ��ȡ�Ѱ�װ����������(������odbcinst.h��)
	if (!SQLGetInstalledDrivers(szBuf, cbBufMax, &cbBufOut))
		return "";

	// �����Ѱ�װ�������Ƿ���Excel...
	do
	{
		if (strstr(pszBuf, "Excel") != 0)
		{
			//���� !
			sDriver = string(pszBuf);
			break;
		}
		pszBuf = strchr(pszBuf, '\0') + 1;
	}
	while (pszBuf[1] != '\0');

	return sDriver;
}