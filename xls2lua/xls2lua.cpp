// xls2lua.cpp : 定义控制台应用程序的入口点。
//

#include "stdafx.h"
#include "Xls2LuaMgr.h"

#include <iostream>
#include <conio.h>

using namespace std;

int _tmain(int argc, _TCHAR* argv[])
{
	CXls2LuaMgr::getInstance()->LoadExcels();

	char buf[1024];

	//sprintf(buf,"---------");

	printf("---------------------exel2lua-----------------------");

	char cTemp;
	cTemp = getchar();
	
	
	if (cTemp)
	{
		printf("----------------------按任意键结束-------------\n");
		exit(0);
	}

	

	return 0;
}

