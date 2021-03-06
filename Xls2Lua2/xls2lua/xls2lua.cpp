#include "stdafx.h"
#include "xls2lua.h"
#include "CSpreadSheet.h"
#include <fstream>
#include <iostream>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif
using namespace std;
void convertXls2Lua();


int _tmain(int argc, TCHAR* argv[], TCHAR* envp[])
{

	convertXls2Lua();

	system("pause");
	return 0;
}



//具体实现
void convertXls2Lua()
{
	system("color 0E");
	printf("\t*****************************************************************************************\t\n");
	printf("\t*\t\t\t\t本工具使用注意事项                                    \t*\n");
	printf("\t*\t将Excel表格中的单元格类型设置为文本                                     \t*\n");
	printf("\t*\t请将需要转成lua配置的xls文件放到工具所在目录下的xls目录。                 \t*\n");
	printf("\t*\t每个xls文件只支持转第一个sheet，并且命名为Sheet1。                        \t*\n");
	printf("\t*\tExcel配置表第一行为字段名，第二行为字段注释（可不填)，从第三行开始为数据。 \t*\n");
	printf("\t*\t字段注释不导出，如果第一行字段名为空，则该字段\"\"导出到lua。               \t*\n");
	printf("\t*\t导出的lua文件在本目录下的lua文件夹里，如果没有该文件夹请自己建该文件夹。  \t*\n");
	printf("\t*\t表头字段名不能以大写字母开头后跟数字，如F1、F2，这是Excel的关键字。       \t*\n");
	printf("\t*\t如果提示缺少驱动程序，请完整安装Office任意版本。                          \t*\n");
	printf("\t*****************************************************************************************\t\n");
	system("pause");
	cout << "开始导出文件，请稍等！" << endl;


	CFileFind finder; 
	CString strFile, strTitle, strTemp; 
	BOOL bWorking = finder.FindFile("xls\\*.xls"); 
	while(bWorking) 
	{ 
		bWorking = finder.FindNextFile(); 
		strFile = finder.GetFilePath();
		strTitle = finder.GetFileTitle();		

		CSpreadSheet sheet(strFile, "Sheet1", false);
		sheet.BeginTransaction();
			
		short tCol = sheet.GetTotalColumns();
		long tRow = sheet.GetTotalRows();
		long m_num = tRow - 2;

		cout << "行数------------------>tRow:" << tRow << endl;
		cout << "列数------------------>tCol:" << tCol << endl;

		if (tCol <= 0 || tRow <= 3)
		{
			continue;
		}
		
		ofstream outputFile;
		CStringArray headerArray, dataArray;
		sheet.ReadRow(headerArray, 1);

		//初始行
		int m_row = 3;

		for (int k = 1; k <= m_num; k++)
		{
			char oldFile[20];
			char newFile[20];
			memset(newFile, 0, sizeof(newFile));
			memset(oldFile, 0, sizeof(oldFile));
	
			sprintf_s(oldFile, "%s%d%s", "lua\\", k, ".txt");
			outputFile.open(oldFile, std::ios::out | std::ios::trunc);
			
			//每读取一行 写一个txt文件
			while (m_row <= tRow)
			{
				outputFile << "{\n";
				sheet.ReadRow(dataArray, m_row);

				for (int j = 0; j < tCol; ++j)
				{
					outputFile << "\t";
					outputFile << headerArray[j] << " = ";

					if (dataArray[j].FindOneOf(".0") == 0)
					{
						outputFile << atoi(dataArray[j]);
					}
					else
					{
						strTemp.Format("F%d", j + 1);
						if (headerArray[j] == strTemp || dataArray[j] == "")
						{
							outputFile << "\"\"";
						}
						else
						{
							outputFile << dataArray[j];
						}
					}
					outputFile << ", ";
					outputFile << "\n";
				}
				outputFile << "},\n";

				//写新文件名
				sprintf_s(newFile, "%s%d%s", "lua\\", atoi(dataArray[0]), ".txt");
				m_row++;
				break;
			}
		
			outputFile.flush();
			outputFile.close();
			                   
			//判断文件是否存在
			ifstream isExist(newFile);
			if (!isExist)
			{
				rename(oldFile, newFile);		
			}
			else
			{
				isExist.close();
				remove(newFile);
				rename(oldFile, newFile);
			}
			
		}

		cout << strFile << "导出成功！" << endl;
	}
	cout << endl;
	cout << "导出所有xls到txt成功!" << endl;

}
