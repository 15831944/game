#pragma once

#include <string>
using namespace std;

class CXls2LuaMgr
{
public:
	CXls2LuaMgr(void);
	~CXls2LuaMgr(void);

	static CXls2LuaMgr* getInstance()
	{
		static CXls2LuaMgr sInstance;
		return &sInstance;
	}
	void ReadFromExcel();
	string GetExcelDriver();

public:
	void LoadExcels();
	

private:




};

