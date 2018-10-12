// Class to read and write to Excel and text delimited spreadsheet
//
// Created by Yap Chun Wei
// December 2001
// 
// Version 1.1
// Updates: Fix bug in ReadRow() which prevent reading of single column spreadsheet
// Modified by jingzhou xu

#ifndef CSPREADSHEET_H
#define CSPREADSHEET_H

#include <odbcinst.h>
#include <afxdb.h>

class CSpreadSheet
{
public:
	// 打开表格进行读和写
	//	该构造函数将打开Excel(xls)文件或其他制定工作簿的文件以供读写.创建一个CSpreadSheet对象.
	//	参数：
	//	File : 文件路径, 可以是绝对路径或相对路径, 如果文件不存在将创建一个文件.
	//	SheetOrSeparator 工作簿名.
	//  Backup 制定是否备份文件, 默认未备份文件, 如果文件存在，将创建一个名为CSpreadSheetBackup 的备份文件
	CSpreadSheet(CString File, CString SheetOrSeparator, bool Backup = true); 

	//清理
	~CSpreadSheet(); 

	//添加标题行到电子表格
	//	该函数将在打开的工作簿的首行添加一个头（字段）.对于Excel，每列字段必须唯一.对于特
	//	定特征的文本文件没有限制.对于一个打开的工作簿文件, 默认将添加一列, 如果设置replace = true
	//	将替代存在的字段值.该函数返回一个Bool类型的值.对于Excel, 该函数需
	//	在添加任意行之前调用.对于特定特征的文本文件，该函数可选.
	//	参数:
	//	FieldNames   字段名数组.
	//	Replace    如字段存在，该参数将决定是否替代原有字段.
	bool AddHeaders(CStringArray &FieldNames, bool replace = false); 

	// 删除工作簿
	bool DeleteSheet(); 

	// 删除整个Excel电子表格内容，表本身并不删除
	bool DeleteSheet(CString SheetName); 

	// 插入或替换一行电子表格，默认是添加新行。
	//该函数将追加、插入或替代一行到已经打开的文档, 默认追加到行的尾部.替代将以变量的值而定, 新的一行将插入或替代到指定的行.
	//	参数:
	//	RowValues   行值
	//	Row    行编号, 如果Row = 1 第一行.即字段行.
	//	Replace   如果该行存在，指示是否替代原有行.
	bool AddRow(CStringArray &RowValues, long row = 0, bool replace = false); 

	//使用标题行或列字母替换或添加到Excel电子表格中的单元格
	//默认是将单元格添加到新行中
	//如果要强制列用作标题名称，请将AUTO设置为FALSE
	// 参数:
	//	CellValue   填充的单元格的值。
	//	Column   列编号
	//	column   列名
	//	Row    含编号，如果Row = 1 第一行，即字段行.
	//	Auto    是否让函数自动判断自动判断字段.
	bool AddCell(CString CellValue, CString column, long row = 0, bool Auto = true);

	//使用列编号将单元格替换或添加到电子表格中
	//默认是将单元格添加到新行中
	bool AddCell(CString CellValue, short column, long row = 0); 

	//在Excel电子表格中搜索和替换行
	bool ReplaceRows(CStringArray &NewRowValues, CStringArray &OldRowValues);

	// 从电子表格中读取一行,默认读取第一行
	//参数:
	//	RowValues   用于存储读取到的值。
	//	Row    行编号.默认为第一行.
	bool ReadRow(CStringArray &RowValues, long row = 0);

	// 从Excel电子表格中使用标题行或列字母表读取列
	//如果要强制列用作标题名称，请将AUTO设置为FALSE
	//参数:
	//	Short column   列编号
	//	CString column   列名或字段名.
	//	Columnvalues   存储读取到的值.
	//	Auto     设置函数自动扫描列名或字段.
	bool ReadColumn(CStringArray &ColumnValues, CString column, bool Auto = true);

	//使用列号从电子表格读取列
	bool ReadColumn(CStringArray &ColumnValues, short column); 

	//使用标题行或列字母表从Excel电子表格中读取单元格。
	//默认读取下一行中的下一个单元格
	//如果要强制列用作标题名称，请将AUTO设置为FALSE
	//参数
	//	CellValue    存储单元格的值.
	//	Short column   列编号.
	//	Row     行编号
	//	CString column   列名
	//	Auto      设置函数自动扫描列名或字段.
	bool ReadCell (CString &CellValue, CString column, long row = 0, bool Auto = true); 
	
	//使用列号从电子表格中读取单元格。默认读取第一行中的第一个单元格
	bool ReadCell (CString &CellValue, short column, long row = 0); 

	//开始事务
	void BeginTransaction(); 

	 //将更改保存到电子表格
	bool Commit(); 

	//撤消对电子表格的更改
	bool RollBack(); 

	//准备文件
	bool Convert(CString SheetOrSeparator);
	
	//从电子表格中获取标题行
	inline void GetFieldNames (CStringArray &FieldNames) {FieldNames.RemoveAll(); FieldNames.Copy(m_aFieldNames);}

	//获取电子表格中的行总数
	inline long GetTotalRows() {return m_dTotalRows;} 

	//获取电子表格中的列总数
	inline short GetTotalColumns() {return m_dTotalColumns;} 

	//在电子表格中获取当前选定的行
	inline long GetCurrentRow() {return m_dCurrentRow;} 

	//获取备份状态。如果备份成功，则如果电子表格不是备份，则为false
	inline bool GetBackupStatus() {return m_bBackup;} 

	//获取事务状态。如果事务启动，则如果事务未启动或启动出错，则为false。
	inline bool GetTransactionStatus() {return m_bTransaction;} 

	//获取最后一条错误消息
	inline CString GetLastError() {return m_sLastError;} 


private:
	//打开用于读或书写的文本文件
	bool Open(); 

	//获取Excel ODBC驱动
	void GetExcelDriver();

	//将字母表中的Excel列转换为列号
	short CalculateColumnNumber(CString column, bool Auto); 

	//新创建的电子表格或者以前创建的电子表格
	bool m_bAppend;

	//备份状态
	bool m_bBackup;

	//文件是否是Excel电子表格或文本
	bool m_bExcel; 

	//事务状态
	bool m_bTransaction;

	//当前行索引，从1开始
	long m_dCurrentRow; 

	//电子表格中的行总数
	long m_dTotalRows;

	//Excel电子表格中的列总数
	//文本最大的列数
	short m_dTotalColumns; 

	//SQL语句打开Excel电子表格用于阅读
	CString m_sSql;
	
	//DSN字符串打开Excel电子表格用于读写
	CString m_sDsn; 
	
	//SQL语句或函数使用的临时字符串
	CString m_stempSql; 
	
	//函数使用的临时字符串
	CString m_stempString; 
	
	//Excel电子表格的表名
	CString m_sSheetName; 
	
	//Excel 驱动名
	CString m_sExcelDriver;
	
	//电子表格文件名
	CString m_sFile;

	//文本分隔的电子表格中的分隔符
	CString m_sSeparator; 
	
	//上次错误信息
	CString m_sLastError;

	//函数使用的临时数组
	CStringArray m_atempArray; 

	//电子表格标题行
	CStringArray m_aFieldNames;

	//电子表格中所有行的内容
	CStringArray m_aRows;

	//Excel电子表格的数据库变量
	CDatabase *m_Database;

	//Excel电子表格的记录集
	CRecordset *m_rSheet; 
};

// 打开表格进行读和写
CSpreadSheet::CSpreadSheet(CString File, CString SheetOrSeparator, bool Backup) :
m_Database(NULL), m_rSheet(NULL), m_sFile(File),
m_dTotalRows(0), m_dTotalColumns(0), m_dCurrentRow(1),
m_bAppend(false), m_bBackup(Backup), m_bTransaction(false)
{
	// 检测文件是否是Excel电子表格还是文本分隔文件
	m_stempString = m_sFile.Right(4);
	m_stempString.MakeLower();
	if (m_stempString == ".xls") // File is an Excel spreadsheet
	{
		m_bExcel = true;
		m_sSheetName = SheetOrSeparator;
		m_sSeparator = ",;.?";
	}
	else // File is a text delimited file
	{
		m_bExcel = false;
		m_sSeparator = SheetOrSeparator;
	}

	if (m_bExcel) // 如果文件是一个Excel表格
	{
		m_Database = new CDatabase;
		GetExcelDriver();
		m_sDsn.Format("DRIVER={%s};DSN='';FIRSTROWHASNAMES=1;READONLY=FALSE;CREATE_DB=\"%s\";DBQ=%s", m_sExcelDriver, m_sFile, m_sFile);

		if (Open())
		{
			if (m_bBackup)
			{
				if ((m_bBackup) && (m_bAppend))
				{
					CString tempSheetName = m_sSheetName;
					m_sSheetName = "CSpreadSheetBackup";
					m_bAppend = false;
					if (!Commit())
					{
						m_bBackup = false;
					}
					m_bAppend = true;
					m_sSheetName = tempSheetName;
					m_dCurrentRow = 1;
				}
			}
		}
	}
	else // 如果文件是文本分隔文件
	{
		if (Open())
		{
			if ((m_bBackup) && (m_bAppend))
			{
				m_stempString = m_sFile;
				m_stempSql.Format("%s.bak", m_sFile);
				m_sFile = m_stempSql;
				if (!Commit())
				{
					m_bBackup = false;
				}
				m_sFile = m_stempString;
			}
		}
	}
}

// Perform some cleanup functions
CSpreadSheet::~CSpreadSheet()
{
	if (m_Database != NULL)
	{
		m_Database->Close();
		delete m_Database;
	}
}

// Add header row to spreadsheet
bool CSpreadSheet::AddHeaders(CStringArray &FieldNames, bool replace)
{
	if (m_bAppend) // Append to old Sheet
	{
		if (replace) // Replacing header row rather than adding new columns
		{
			if (!AddRow(FieldNames, 1, true))
			{
				return false;
			}
			else
			{
				return true;
			}
		}

		if (ReadRow(m_atempArray, 1)) // Add new columns
		{
			if (m_bExcel)
			{
				// Check for duplicate header row field
				for (int i = 0; i < FieldNames.GetSize(); i++)
				{
					for (int j = 0; j < m_atempArray.GetSize(); j++)
					{
						if (FieldNames.GetAt(i) == m_atempArray.GetAt(j))
						{
							m_sLastError.Format("Duplicate header row field:%s\n", FieldNames.GetAt(i));
							return false;
						}
					}
				}	
			}

			m_atempArray.Append(FieldNames);
			if (!AddRow(m_atempArray, 1, true))
			{
				m_sLastError = "Problems with adding headers\n";
				return false;
			}

			// Update largest number of columns if necessary
			if (m_atempArray.GetSize() > m_dTotalColumns)
			{
				m_dTotalColumns = m_atempArray.GetSize();
			}
			return true;
		}
		return false;				
	}
	else // New Sheet
	{
		m_dTotalColumns = FieldNames.GetSize();
		if (!AddRow(FieldNames, 1, true))
		{
			return false;
		}
		else
		{
			m_dTotalRows = 1;
			return true;
		}
	}
}

// Clear text delimited file content
bool CSpreadSheet::DeleteSheet()
{
	if (m_bExcel)
	{
		if (DeleteSheet(m_sSheetName))
		{
			return true;
		}
		else
		{
			m_sLastError = "Error deleting sheet\n";
			return false;
		}
	}
	else
	{
		m_aRows.RemoveAll();
		m_aFieldNames.RemoveAll();
		m_dTotalColumns = 0;
		m_dTotalRows = 0;
		if (!m_bTransaction)
		{
			Commit();			
		}
		m_bAppend = false; // Set flag to new sheet
		return true;		
	}
}

// Clear entire Excel spreadsheet content. The sheet itself is not deleted
bool CSpreadSheet::DeleteSheet(CString SheetName)
{
	if (m_bExcel) // If file is an Excel spreadsheet
	{
		// Delete sheet
		m_Database->OpenEx(m_sDsn, CDatabase::noOdbcDialog);
		SheetName = "[" + SheetName + "$A1:IV65536]";
		m_stempSql.Format ("DROP TABLE %s", SheetName);
		try
		{
			m_Database->ExecuteSQL(m_stempSql);
			m_Database->Close();
			m_aRows.RemoveAll();
			m_aFieldNames.RemoveAll();
			m_dTotalColumns = 0;
			m_dTotalRows = 0;
		}
		catch(CDBException *e)
		{
			m_sLastError = e->m_strError;
			m_Database->Close();
			return false;
		}
		return true;
	}
	else // if file is a text delimited file
	{
		return DeleteSheet();
	}
}

// Insert or replace a row into spreadsheet. 
// Default is add new row.
bool CSpreadSheet::AddRow(CStringArray &RowValues, long row, bool replace)
{
	long tempRow;
	
	if (row == 1)
	{
		if (m_bExcel) 
		{
			// Check for duplicate header row field for Excel spreadsheet
			for (int i = 0; i < RowValues.GetSize(); i++)
			{
				for (int j = 0; j < RowValues.GetSize(); j++)
				{
					if ((i != j) && (RowValues.GetAt(i) == RowValues.GetAt(j)))
					{
						m_sLastError.Format("Duplicate header row field:%s\n", RowValues.GetAt(i));
						return false;
					}
				}
			}
			
			// Check for reduced header row columns
			if (RowValues.GetSize() < m_dTotalColumns)
			{
				m_sLastError = "Number of columns in new header row cannot be less than the number of columns in previous header row";
				return false;
			}
			m_dTotalColumns = RowValues.GetSize();
		}

		// Update header row
		m_aFieldNames.RemoveAll();
		m_aFieldNames.Copy(RowValues);
	}
	else
	{
		if (m_bExcel)
		{
			if (m_dTotalColumns == 0)
			{
				m_sLastError = "No header row. Add header row first\n";
				return false;
			}
		}
	}

	if (m_bExcel) // For Excel spreadsheet
	{
		if (RowValues.GetSize() > m_aFieldNames.GetSize())
		{
			m_sLastError = "Number of columns to be added cannot be greater than the number of fields\n";
			return false;
		}
	}
	else // For text delimited spreadsheet
	{
		// Update largest number of columns if necessary
		if (RowValues.GetSize() > m_dTotalColumns)
		{
			m_dTotalColumns = RowValues.GetSize();
		}
	}

	// Convert row values
	m_stempString.Empty();
	for (int i = 0; i < RowValues.GetSize(); i++)
	{
		if (i != RowValues.GetSize()-1) // Not last column
		{
			m_stempSql.Format("\"%s\"%s", RowValues.GetAt(i), m_sSeparator);
			m_stempString += m_stempSql;
		}
		else // Last column
		{
			m_stempSql.Format("\"%s\"", RowValues.GetAt(i));
			m_stempString += m_stempSql;
		}
	}
	
	if (row)
	{
		if (row <= m_dTotalRows) // Not adding new rows
		{
			if (replace) // Replacing row
			{
				m_aRows.SetAt(row-1, m_stempString);
			}
			else // Inserting row
			{
				m_aRows.InsertAt(row-1, m_stempString);
				m_dTotalRows++;
			}

			if (!m_bTransaction)
			{
				Commit();
			}
			return true;
		}
		else // Adding new rows
		{
			// Insert null rows until specified row
			m_dCurrentRow = m_dTotalRows;
			m_stempSql.Empty();
			CString nullString;
			for (int i = 1; i <= m_dTotalColumns; i++)
			{
				if (i != m_dTotalColumns)
				{
					if (m_bExcel)
					{
						nullString.Format("\" \"%s", m_sSeparator);
					}
					else
					{
						nullString.Format("\"\"%s", m_sSeparator);
					}
					m_stempSql += nullString;
				}
				else
				{
					if (m_bExcel)
					{
						m_stempSql += "\" \"";
					}
					else
					{
						m_stempSql += "\"\"";
					}
				}
			}
			for (int j = m_dTotalRows + 1; j < row; j++)
			{
				m_dCurrentRow++;
				m_aRows.Add(m_stempSql);
			}
		}
	}
	else
	{
		tempRow = m_dCurrentRow;
		m_dCurrentRow = m_dTotalRows;
	}

	// Insert new row
	m_dCurrentRow++;
	m_aRows.Add(m_stempString);
	
	if (row > m_dTotalRows)
	{
		m_dTotalRows = row;
	}
	else if (!row)
	{
		m_dTotalRows = m_dCurrentRow;
		m_dCurrentRow = tempRow;
	}
	if (!m_bTransaction)
	{
		Commit();
	}
	return true;
}

	// Replace or add a cell into Excel spreadsheet using header row or column alphabet. 
	// Default is add cell into new row.
	// Set Auto to false if want to force column to be used as header name
bool CSpreadSheet::AddCell(CString CellValue, CString column, long row, bool Auto)
{
	short columnIndex = CalculateColumnNumber(column, Auto);
	if (columnIndex == 0)
	{
		return false;
	}

	if (AddCell(CellValue, columnIndex, row))
	{
		return true;
	}
	return false;
}

	// Replace or add a cell into spreadsheet using column number
	// Default is add cell into new row.
bool CSpreadSheet::AddCell(CString CellValue, short column, long row)
{
	if (column == 0)
	{
		m_sLastError = "Column cannot be zero\n";
		return false;
	}

	long tempRow;

	if (m_bExcel) // For Excel spreadsheet
	{
		if (column > m_aFieldNames.GetSize() + 1)
		{
			m_sLastError = "Cell column to be added cannot be greater than the number of fields\n";
			return false;
		}
	}
	else // For text delimited spreadsheet
	{
		// Update largest number of columns if necessary
		if (column > m_dTotalColumns)
		{
			m_dTotalColumns = column;
		}
	}

	if (row)
	{
		if (row <= m_dTotalRows)
		{
			ReadRow(m_atempArray, row);
	
			// Change desired row
			m_atempArray.SetAtGrow(column-1, CellValue);

			if (row == 1)
			{
				if (m_bExcel) // Check for duplicate header row field
				{										
					for (int i = 0; i < m_atempArray.GetSize(); i++)
					{
						for (int j = 0; j < m_atempArray.GetSize(); j++)
						{
							if ((i != j) && (m_atempArray.GetAt(i) == m_atempArray.GetAt(j)))
							{
								m_sLastError.Format("Duplicate header row field:%s\n", m_atempArray.GetAt(i));
								return false;
							}
						}
					}
				}

				// Update header row
				m_aFieldNames.RemoveAll();
				m_aFieldNames.Copy(m_atempArray);
			}	

			if (!AddRow(m_atempArray, row, true))
			{
				return false;
			}

			if (!m_bTransaction)
			{
				Commit();
			}
			return true;
		}
		else
		{
			// Insert null rows until specified row
			m_dCurrentRow = m_dTotalRows;
			m_stempSql.Empty();
			CString nullString;
			for (int i = 1; i <= m_dTotalColumns; i++)
			{
				if (i != m_dTotalColumns)
				{
					if (m_bExcel)
					{
						nullString.Format("\" \"%s", m_sSeparator);
					}
					else
					{
						nullString.Format("\"\"%s", m_sSeparator);
					}
					m_stempSql += nullString;
				}
				else
				{
					if (m_bExcel)
					{
						m_stempSql += "\" \"";
					}
					else
					{
						m_stempSql += "\"\"";
					}
				}
			}
			for (int j = m_dTotalRows + 1; j < row; j++)
			{
				m_dCurrentRow++;
				m_aRows.Add(m_stempSql);
			}
		}
	}
	else
	{
		tempRow = m_dCurrentRow;
		m_dCurrentRow = m_dTotalRows;
	}

	// Insert cell
	m_dCurrentRow++;
	m_stempString.Empty();
	for (int j = 1; j <= m_dTotalColumns; j++)
	{
		if (j != m_dTotalColumns) // Not last column
		{
			if (j != column)
			{
				if (m_bExcel)
				{
					m_stempSql.Format("\" \"%s", m_sSeparator);
				}
				else
				{
					m_stempSql.Format("\"\"%s", m_sSeparator);
				}
				m_stempString += m_stempSql;
			}
			else
			{
				m_stempSql.Format("\"%s\"%s", CellValue, m_sSeparator);
				m_stempString += m_stempSql;
			}
		}
		else // Last column
		{
			if (j != column)
			{
				if (m_bExcel)
				{
					m_stempString += "\" \"";
				}
				else
				{
					m_stempString += "\"\"";
				}
			}
			else
			{
				m_stempSql.Format("\"%s\"", CellValue);
				m_stempString += m_stempSql;
			}
		}
	}	

	m_aRows.Add(m_stempString);
	
	if (row > m_dTotalRows)
	{
		m_dTotalRows = row;
	}
	else if (!row)
	{
		m_dTotalRows = m_dCurrentRow;
		m_dCurrentRow = tempRow;
	}
	if (!m_bTransaction)
	{
		Commit();
	}
	return true;
}

	// Search and replace rows in Excel spreadsheet
bool CSpreadSheet::ReplaceRows(CStringArray &NewRowValues, CStringArray &OldRowValues)
{
	if (m_bExcel) // If file is an Excel spreadsheet
	{
		m_Database->OpenEx(m_sDsn, CDatabase::noOdbcDialog);
		m_stempSql.Format("UPDATE [%s] SET ", m_sSheetName);
		for (int i = 0; i < NewRowValues.GetSize(); i++)
		{
			m_stempString.Format("[%s]='%s', ", m_aFieldNames.GetAt(i), NewRowValues.GetAt(i));
			m_stempSql = m_stempSql + m_stempString;
		}
		m_stempSql.Delete(m_stempSql.GetLength()-2, 2);
		m_stempSql = m_stempSql + " WHERE (";
		for (int j = 0; j < OldRowValues.GetSize()-1; j++)
		{
			m_stempString.Format("[%s]='%s' AND ", m_aFieldNames.GetAt(j), OldRowValues.GetAt(j));
			m_stempSql = m_stempSql + m_stempString;
		}
		m_stempSql.Delete(m_stempSql.GetLength()-4, 5);
		m_stempSql += ")";

		try
		{
			m_Database->ExecuteSQL(m_stempSql);
			m_Database->Close();
			Open();
			return true;
		}
		catch(CDBException *e)
		{
			m_sLastError = e->m_strError;
			m_Database->Close();
			return false;
		}
	}
	else // if file is a text delimited file
	{
		m_sLastError = "Function not available for text delimited file\n";
		return false;
	}
}

	// Read a row from spreadsheet. 
	// Default is read the next row
bool CSpreadSheet::ReadRow(CStringArray &RowValues, long row)
{
	// 检查输入的行是否多于表中的行数
	if (row <= m_aRows.GetSize())
	{
		if (row != 0)
		{
			m_dCurrentRow = row;
		}
		else if (m_dCurrentRow > m_aRows.GetSize())
		{
			return false;
		}
		// 读取所需行
		RowValues.RemoveAll();
		m_stempString = m_aRows.GetAt(m_dCurrentRow-1);
	//	CString str;
		//str.Format("%s", m_stempString);

		//std::cout <<" 12313231" << m_stempString;
		m_dCurrentRow++;

		// Search for separator to split row
		int separatorPosition;
		m_stempSql.Format("\"%s\"", m_sSeparator);
		//m_stempSql.Format("%d", m_sSeparator);
		//str.Format("%f", m_stempString);

		separatorPosition = m_stempString.Find(m_stempSql); // If separator is "?"
		if (separatorPosition != -1)
		{
			// Save columns
			int nCount = 0;
			int stringStartingPosition = 0;
			while (separatorPosition != -1)
			{
				nCount = separatorPosition - stringStartingPosition;
				RowValues.Add(m_stempString.Mid(stringStartingPosition, nCount));
				stringStartingPosition = separatorPosition + m_stempSql.GetLength();
				separatorPosition = m_stempString.Find(m_stempSql, stringStartingPosition);
			}
			nCount = m_stempString.GetLength() - stringStartingPosition;
			RowValues.Add(m_stempString.Mid(stringStartingPosition, nCount));

			// 删除第一列的引号
			m_stempString = RowValues.GetAt(0);
			m_stempString.Delete(0, 1);
			RowValues.SetAt(0, m_stempString);
			
			// 删除最后一列的引号
			m_stempString = RowValues.GetAt(RowValues.GetSize()-1);
			m_stempString.Delete(m_stempString.GetLength()-1, 1);
			RowValues.SetAt(RowValues.GetSize()-1, m_stempString);

			return true;
		}
		else
		{
			// Save columns
			separatorPosition = m_stempString.Find(m_sSeparator); // if separator is ?
			if (separatorPosition != -1)
			{
				int nCount = 0;
				int stringStartingPosition = 0;
				while (separatorPosition != -1)
				{
					nCount = separatorPosition - stringStartingPosition;
					RowValues.Add(m_stempString.Mid(stringStartingPosition, nCount));
					stringStartingPosition = separatorPosition + m_sSeparator.GetLength();
					separatorPosition = m_stempString.Find(m_sSeparator, stringStartingPosition);
				}
				nCount = m_stempString.GetLength() - stringStartingPosition;
				RowValues.Add(m_stempString.Mid(stringStartingPosition, nCount));
				return true;
			}
			else	// Treat spreadsheet as having one column
			{
				// Remove opening and ending quotes if any
				int quoteBegPos = m_stempString.Find('\"');
				int quoteEndPos = m_stempString.ReverseFind('\"');
				if ((quoteBegPos == 0) && (quoteEndPos == m_stempString.GetLength()-1))
				{
					m_stempString.Delete(0, 1);
					m_stempString.Delete(m_stempString.GetLength()-1, 1);
				}

				RowValues.Add(m_stempString);
			}
		}
	}
	m_sLastError = "Desired row is greater than total number of rows in spreadsheet\n";
	return false;
}

	// Read a column from Excel spreadsheet using header row or column alphabet. 
	// Set Auto to false if want to force column to be used as header name
bool CSpreadSheet::ReadColumn(CStringArray &ColumnValues, CString column, bool Auto)
{
	short columnIndex = CalculateColumnNumber(column, Auto);
	if (columnIndex == 0)
	{
		return false;
	}

	if (ReadColumn(ColumnValues, columnIndex))
	{
		return true;
	}
	return false;
}

	// Read a column from spreadsheet using column number
bool CSpreadSheet::ReadColumn(CStringArray &ColumnValues, short column)
{
	if (column == 0)
	{
		m_sLastError = "Column cannot be zero\n";
		return false;
	}

	int tempRow = m_dCurrentRow;
	m_dCurrentRow = 1;
	ColumnValues.RemoveAll();
	for (int i = 1; i <= m_aRows.GetSize(); i++)
	{
		// 读每一行
		if (ReadRow(m_atempArray, i))
		{
			// Get value of cell in desired column
			if (column <= m_atempArray.GetSize())
			{
				ColumnValues.Add(m_atempArray.GetAt(column-1));
			}
			else
			{
				ColumnValues.Add("");
			}
		}
		else
		{
			m_dCurrentRow = tempRow;
			m_sLastError = "Error reading row\n";
			return false;
		}
	}
	m_dCurrentRow = tempRow;
	return true;
}

	// Read a cell from Excel spreadsheet using header row or column alphabet. 
	// Default is read the next cell in next row. 
	// Set Auto to false if want to force column to be used as header name
bool CSpreadSheet::ReadCell (CString &CellValue, CString column, long row, bool Auto)
{
	short columnIndex = CalculateColumnNumber(column, Auto);
	if (columnIndex == 0)
	{
		return false;
	}

	if (ReadCell(CellValue, columnIndex, row))
	{
		return true;
	}
	return false;
}

	// Read a cell from spreadsheet using column number. 
	// Default is read the next cell in next row.
bool CSpreadSheet::ReadCell (CString &CellValue, short column, long row)
{
	if (column == 0)
	{
		m_sLastError = "Column cannot be zero\n";
		return false;
	}

	int tempRow = m_dCurrentRow;
	if (row)
	{
		m_dCurrentRow = row;
	}
	if (ReadRow(m_atempArray, m_dCurrentRow))
	{
		// Get value of cell in desired column
		if (column <= m_atempArray.GetSize())
		{
			CellValue = m_atempArray.GetAt(column-1);
		}
		else
		{
			CellValue.Empty();
			m_dCurrentRow = tempRow;
			return false;
		}
		m_dCurrentRow = tempRow;
		return true;
	}
	m_dCurrentRow = tempRow;
	m_sLastError = "Error reading row\n";
	return false;
}

	// Begin transaction
void CSpreadSheet::BeginTransaction()
{
	m_bTransaction = true;
}

	// Save changes to spreadsheet
bool CSpreadSheet::Commit()
{
	if (m_bExcel) // If file is an Excel spreadsheet
	{
		m_Database->OpenEx(m_sDsn, CDatabase::noOdbcDialog);

		if (m_bAppend)
		{
			// Delete old sheet if it exists
			m_stempString= "[" + m_sSheetName + "$A1:IV65536]";
			m_stempSql.Format ("DROP TABLE %s", m_stempString);
			try
			{
				m_Database->ExecuteSQL(m_stempSql);
			}
			catch(CDBException *e)
			{
				m_sLastError = e->m_strError;
				m_Database->Close();
				return false;
			}
			
			// Create new sheet
			m_stempSql.Format("CREATE TABLE [%s$A1:IV65536] (", m_sSheetName);
			for (int j = 0; j < m_aFieldNames.GetSize(); j++)
			{
				m_stempSql = m_stempSql + "[" + m_aFieldNames.GetAt(j) +"]" + " char(255), ";
			}
			m_stempSql.Delete(m_stempSql.GetLength()-2, 2);
			m_stempSql += ")";
		}
		else
		{
			// Create new sheet
			m_stempSql.Format("CREATE TABLE [%s] (", m_sSheetName);
			for (int i = 0; i < m_aFieldNames.GetSize(); i++)
			{
				m_stempSql = m_stempSql + "[" + m_aFieldNames.GetAt(i) +"]" + " char(255), ";
			}
			m_stempSql.Delete(m_stempSql.GetLength()-2, 2);
			m_stempSql += ")";
		}

		try
		{
			m_Database->ExecuteSQL(m_stempSql);
			if (!m_bAppend)
			{
				m_dTotalColumns = m_aFieldNames.GetSize();
				m_bAppend = true;
			}
		}
		catch(CDBException *e)
		{
			m_sLastError = e->m_strError;
			m_Database->Close();
			return false;
		}

		// Save changed data
		for (int k = 1; k < m_dTotalRows; k++)
		{
			ReadRow(m_atempArray, k+1);

			// Create Insert SQL
			m_stempSql.Format("INSERT INTO [%s$A1:IV%d] (", m_sSheetName, k);
			for (int i = 0; i < m_atempArray.GetSize(); i++)
			{
				m_stempString.Format("[%s], ", m_aFieldNames.GetAt(i));
				m_stempSql = m_stempSql + m_stempString;
			}
			m_stempSql.Delete(m_stempSql.GetLength()-2, 2);
			m_stempSql += ") VALUES (";
			for (int j = 0; j < m_atempArray.GetSize(); j++)
			{
				m_stempString.Format("'%s', ", m_atempArray.GetAt(j));
				m_stempSql = m_stempSql + m_stempString;
			}
			m_stempSql.Delete(m_stempSql.GetLength()-2, 2);
			m_stempSql += ")";

			// Add row
			try
			{
				m_Database->ExecuteSQL(m_stempSql);
			}
			catch(CDBException *e)
			{
				m_sLastError = e->m_strError;
				m_Database->Close();
				return false;
			}
		}
		m_Database->Close();
		m_bTransaction = false;
		return true;
	}
	else // if file is a text delimited file
	{
		try
		{
			CFile *File = NULL;
			File = new CFile(m_sFile, CFile::modeCreate | CFile::modeWrite  | CFile::shareDenyNone);
			if (File != NULL)
			{
				CArchive *Archive = NULL;
				Archive = new CArchive(File, CArchive::store);
				if (Archive != NULL)
				{
					for (int i = 0; i < m_aRows.GetSize(); i++)
					{
						Archive->WriteString(m_aRows.GetAt(i));
						Archive->WriteString("\r\n");
					}
					delete Archive;
					delete File;
					m_bTransaction = false;
					return true;
				}
				delete File;
			}
		}
		catch(...)
		{
		}
		m_sLastError = "Error writing file\n";
		return false;
	}
}

// Undo changes to spreadsheet
bool CSpreadSheet::RollBack()
{
	if (Open())
	{
		m_bTransaction = false;
		return true;
	}
	m_sLastError = "Error in returning to previous state\n";
	return false;
}

bool CSpreadSheet::Convert(CString SheetOrSeparator)
{
	// Prepare file
	m_stempString = m_sFile;
	m_stempString.Delete(m_stempString.GetLength()-4, 4);
	if (m_bExcel) // If file is an Excel spreadsheet
	{
		m_stempString += ".csv";
		CSpreadSheet tempSheet(m_stempString, SheetOrSeparator, false);
		
		// Stop convert if text delimited file exists
		if (tempSheet.GetTotalColumns() != 0)
		{
			return false;
		}

		tempSheet.BeginTransaction();

		for (int i = 1; i <= m_dTotalRows; i++)
		{
			if (!ReadRow(m_atempArray, i))
			{
				return false;
			}
			if (!tempSheet.AddRow(m_atempArray, i))
			{
				return false;
			}
		}
		if (!tempSheet.Commit())
		{
			return false;
		}
		return true;
	}
	else // if file is a text delimited file
	{
		m_stempString += ".xls";
		CSpreadSheet tempSheet(m_stempString, SheetOrSeparator, false);

		// Stop convert if Excel file exists
		if (tempSheet.GetTotalColumns() != 0)
		{
			return false;
		}

		GetFieldNames(m_atempArray);

		// Check for duplicate header row field
		bool duplicate = false;
		for (int i = 0; i < m_atempArray.GetSize(); i++)
		{
			for (int j = 0; j < m_atempArray.GetSize(); j++)
			{
				if ((i != j) && (m_atempArray.GetAt(i) == m_atempArray.GetAt(j)))
				{
					m_sLastError.Format("Duplicate header row field:%s\n", m_atempArray.GetAt(i));
					duplicate = true;
				}
			}
		}

		if (duplicate) // Create dummy header row
		{
			m_atempArray.RemoveAll();
			for (int k = 1; k <= m_dTotalColumns; k++)
			{
				m_stempString.Format("%d", k);
				m_atempArray.Add(m_stempString);
			}

			if (!tempSheet.AddHeaders(m_atempArray))
			{
				return false;
			}

			for (int l = 1; l <= m_dTotalRows; l++)
			{
				if (!ReadRow(m_atempArray, l))
				{
					return false;
				}
				if (!tempSheet.AddRow(m_atempArray, l+1))
				{
					return false;
				}
			}
			return true;
		}
		else
		{
			if (!tempSheet.AddHeaders(m_atempArray))
			{
				return false;
			}

			for (int l = 2; l <= m_dTotalRows; l++)
			{
				if (!ReadRow(m_atempArray, l))
				{
					return false;
				}
				if (!tempSheet.AddRow(m_atempArray, l))
				{
					return false;
				}
			}
			return true;
		}
	}
}

// Open a text delimited file for reading or writing
bool CSpreadSheet::Open()
{
	if (m_bExcel) // If file is an Excel spreadsheet
	{
		m_Database->OpenEx(m_sDsn, CDatabase::noOdbcDialog);

		// Open Sheet
		m_rSheet = new CRecordset( m_Database );
		m_sSql.Format("SELECT * FROM [%s$A1:IV65536]", m_sSheetName);
		try
		{
			m_rSheet->Open(CRecordset::forwardOnly, m_sSql, CRecordset::readOnly);
		}
		catch(...)
		{
			delete m_rSheet;
			m_rSheet = NULL;
			m_Database->Close();
			return false;
		}

		// Get number of columns
		m_dTotalColumns = m_rSheet->m_nResultCols;

		if (m_dTotalColumns != 0)
		{
			m_aRows.RemoveAll();
			m_stempString.Empty();
			m_bAppend = true;
			m_dTotalRows++; // Keep count of total number of rows
			
			// Get field names i.e header row
			for (int i = 0; i < m_dTotalColumns; i++)
			{
				m_stempSql = m_rSheet->m_rgODBCFieldInfos[i].m_strName;
				m_aFieldNames.Add(m_stempSql);

				// Join up all the columns into a string
				if (i != m_dTotalColumns-1) // Not last column
				{
					m_stempString = m_stempString + "\"" + m_stempSql + "\"" + m_sSeparator;
				}
				else // Last column
				{	
					m_stempString = m_stempString + "\"" + m_stempSql + "\"";
				}				
			}
			
			// Store the header row as the first row in memory
			m_aRows.Add(m_stempString);

			// Read and store the rest of the rows in memory
			while (!m_rSheet->IsEOF())
			{
				m_dTotalRows++; // Keep count of total number of rows
				try
				{
					// Get all the columns in a row
					m_stempString.Empty();
					for (short column = 0; column < m_dTotalColumns; column++)
					{
						m_rSheet->GetFieldValue(column, m_stempSql);

						// Join up all the columns into a string
						if (column != m_dTotalColumns-1) // Not last column
						{
							m_stempString = m_stempString + "\"" + m_stempSql + "\"" + m_sSeparator;
						}
						else // Last column
						{	
							m_stempString = m_stempString + "\"" + m_stempSql + "\"";
						}
					}

					// Store the obtained row in memory
					m_aRows.Add(m_stempString);
					m_rSheet->MoveNext();
				}
				catch (...)
				{
					m_sLastError = "Error reading row\n";
					delete m_rSheet;
					m_rSheet = NULL;
					m_Database->Close();
					return false;
				}
			}		
		}
		
		m_rSheet->Close();
		delete m_rSheet;
		m_rSheet = NULL;
		m_Database->Close();
		m_dCurrentRow = 1;
		return true;
	}
	else // if file is a text delimited file
	{
		try
		{
			CFile *File = NULL;
			File = new CFile(m_sFile, CFile::modeRead | CFile::shareDenyNone);
			if (File != NULL)
			{
				CArchive *Archive = NULL;
				Archive = new CArchive(File, CArchive::load);
				if (Archive != NULL)
				{
					m_aRows.RemoveAll();
					// Read and store all rows in memory
					while(Archive->ReadString(m_stempString))
					{
						m_aRows.Add(m_stempString);
					}
					ReadRow(m_aFieldNames, 1); // Get field names i.e header row
					delete Archive;
					delete File;

					// Get total number of rows
					m_dTotalRows = m_aRows.GetSize();

					// Get the largest number of columns
					for (int i = 0; i < m_aRows.GetSize(); i++)
					{
						ReadRow(m_atempArray, i);
						if (m_atempArray.GetSize() > m_dTotalColumns)
						{
							m_dTotalColumns = m_atempArray.GetSize();
						}
					}

					if (m_dTotalColumns != 0)
					{
						m_bAppend = true;
					}
					return true;
				}
				delete File;
			}
		}
		catch(...)
		{
		}
		m_sLastError = "Error in opening file\n";
		return false;
	}
}

// Convert Excel column in alphabet into column number
short CSpreadSheet::CalculateColumnNumber(CString column, bool Auto)
{
	if (Auto)
	{
		int firstLetter, secondLetter;
		column.MakeUpper();

		if (column.GetLength() == 1)
		{
			firstLetter = column.GetAt(0);
			return (firstLetter - 65 + 1); // 65 is A in ascii
		}
		else if (column.GetLength() == 2)
		{
			firstLetter = column.GetAt(0);
			secondLetter = column.GetAt(1);
			return ((firstLetter - 65 + 1)*26 + (secondLetter - 65 + 1)); // 65 is A in ascii
		}
	}

	// Check if it is a valid field name
	for (int i = 0; i < m_aFieldNames.GetSize(); i++)
	{
		if (!column.Compare(m_aFieldNames.GetAt(i)))
		{
			return (i + 1);
		}
	}
	m_sLastError = "Invalid field name or column alphabet\n";
	return 0;	
}

// Get the name of the Excel-ODBC driver
void CSpreadSheet::GetExcelDriver()
{
	char szBuf[2001];
	WORD cbBufMax = 2000;
	WORD cbBufOut;
	char *pszBuf = szBuf;

	// Get the names of the installed drivers ("odbcinst.h" has to be included )
	if(!SQLGetInstalledDrivers(szBuf,cbBufMax,& cbBufOut))
	{
		m_sExcelDriver = "";
	}
	
	// Search for the driver...
	do
	{
		if( strstr( pszBuf, "Excel" ) != 0 )
		{
			// Found !
			m_sExcelDriver = CString( pszBuf );
			break;
		}
		pszBuf = strchr( pszBuf, '\0' ) + 1;
	}
	while( pszBuf[1] != '\0' );
}

#endif