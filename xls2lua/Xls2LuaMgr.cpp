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
	string sFile = "npc.xls"; // 将被读取的Excel文件名
	char buf[2048];
	TCHAR tBuf[1024];

	// 检索是否安装有Excel驱动 "Microsoft Excel Driver (*.xls)" 
	sDriver = GetExcelDriver();
	if (sDriver.size() <= 0)
	{
		// 没有发现Excel驱动
		//MessageBox(NULL,"提示","发现Exel驱动",MB_OK);
		return;
	}

	// 创建进行存取的字符串
	sprintf(buf,"ODBC;DRIVER={%s};DSN='';DBQ=%s", sDriver.c_str(), sFile.c_str());
	
	
	MultiByteToWideChar(CP_ACP,0,buf,strlen(buf),tBuf,strlen(buf) + 1);
	
	TRY
	{
		// 打开数据库(既Excel文件)
		database.Open(NULL, false, false, tBuf);

		CRecordset recset(&database);

		// 设置读取的查询语句.
		//sSql = "SELECT Name, Age " 
		//	"FROM demo " 
		//	"ORDER BY Name ";

		// 执行查询语句
		//recset.Open(CRecordset::forwardOnly, sSql, CRecordset::readOnly);

		// 获取查询结果
		int nRecord = recset.GetRecordCount();
		for (int i = 0; i < nRecord; i++)
		{
			//CString strTemp;
			//recset.GetFieldValue(i,strTemp);
			//printf("%s ",strTemp.GetString());
		}
		//while (!recset.IsEOF())
		//{
		//	//读取Excel内部数值
		//	recset.GetFieldValue("Name ", sItem1);
		//	recset.GetFieldValue("Age", sItem2);

		//	// 移到下一行
		//	recset.MoveNext();
		//}

		// 关闭数据库
		database.Close();

	}
	CATCH(CDBException, e)
	{
		// 数据库操作产生异常时...
		//MessageBox(NULL,"提示","提示",MB_OK);
	}
	END_CATCH;
}


// 获取ODBC中Excel驱动
string CXls2LuaMgr::GetExcelDriver()
{
	TCHAR szBuf[2001];
	WORD cbBufMax = 2000;
	WORD cbBufOut;
	char *pszBuf = szBuf;
	string sDriver;

	// 获取已安装驱动的名称(涵数在odbcinst.h里)
	if (!SQLGetInstalledDrivers(szBuf, cbBufMax, &cbBufOut))
		return "";

	// 检索已安装的驱动是否有Excel...
	do
	{
		if (strstr(pszBuf, "Excel") != 0)
		{
			//发现 !
			sDriver = string(pszBuf);
			break;
		}
		pszBuf = strchr(pszBuf, '\0') + 1;
	}
	while (pszBuf[1] != '\0');

	return sDriver;
}