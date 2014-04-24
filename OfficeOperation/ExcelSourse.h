/********************************************************************
	created:	2014/03/24
	created:	24:3:2014   22:46
	filename: 	c:\Users\ya\Documents\Visual Studio 2008\Projects\OfficeOperation\OfficeOperation\ExcelSourse.h
	file path:	c:\Users\ya\Documents\Visual Studio 2008\Projects\OfficeOperation\OfficeOperation
	file base:	ExcelSourse
	file ext:	h
	author:		zwy
	
	purpose:	excelSourse
*********************************************************************/
#ifndef ExcelSourse_h__
#define ExcelSourse_h__
#include "excel8.h"
enum ExcelColor
{
  black = 1,
  white = 2,
  red   = 3,
  green = 4,
  blue  = 5,
  yellow = 6,
  NullColor = -4142
};

class _OFFICE_PORT ExcelSourse
{
  private:
	  _Application *App;
	  Workbooks *workbooks;
	  _Workbook *workbook;
	  _Worksheet *sheet;
	  Worksheets *sheets;
	  Range * range;
	  Range * Allrage;
	  Range *Usedrage;
	 int m_RowNum;
	 int m_ColNum;
	 TCHAR m_strCol[256];
	 TCHAR m_strRow[256];
	 TCHAR m_str[256];

   public: 
	   ExcelSourse();
	   ~ExcelSourse();

   public:
	   virtual int SetVisible(bool visible);
	   virtual int Open(LPCTSTR lpszFileName);
	   virtual int SaveAs(LPCTSTR lpszFileName);
	   virtual int Save(LPCTSTR lpszFileName);
	   virtual int Save();
	   virtual int Close();
	   virtual int SetCell(int row,int col,LPCTSTR strValue);
	   virtual int SetCell(int row,int col,long lValue);
	   virtual int SetCell(int row,int col,int nValue);
	   virtual int SetCell(int row,int col,double dValue,int n);
	   //virtual int SetSheetName(LPCTSTR lpszName);
	   virtual int SelectActiveSheet(LPCTSTR lpszSheetName);
	   virtual int SelectCurSheet(LPCTSTR lpszSheetName);
	   virtual int SelectCurSheet(int nSheetIndex);
	   virtual int DeleteSheet(LPCTSTR lpszName);
	   virtual int GetCell(int row,int col,LPTSTR strValue);
	   virtual int GetCell(int row,int col,COleVariant *strValue);
	   virtual int GetRow(int row,int nStartCol,int nEndCol,CStringArray *strArr);
	   virtual int GetCol(int col,int nStartRow,int nEndRow,CStringArray *strArr);
	   virtual int GetColCell(int col,CStringArray *strArr);
	   virtual int GetRowCell(int row,CStringArray *strArr);
	   virtual int SetColWidth(int col,int nWidth);
	   virtual int SetRowHeight(int row,int nHeight);
	   virtual int GetCellWidth(int row,int col,int &nWeigth);
	   virtual int GetCellHeight(int row,int col,int &nHeight);
       virtual int SetCellPostill(int row,int col,LPCTSTR lpszstr);
	   virtual int SetCellPostill(int row,int col,LPCTSTR lpszstr,LPCTSTR lpszauthor);
	   virtual int ClearCellPostill(int row,int col);
	   virtual int GetCellPostill(int row,int col,LPTSTR lpszstr);
	   virtual int SetCellColor(int row,int col,ExcelColor color);
	   virtual int GetCellColor(int row,int col,ExcelColor &color);
	   virtual int SetRowColor(int row,ExcelColor color);
	  
	   virtual int GetAllRow();
	   virtual int GetAllCol();
	   virtual int GetSheetNum();
	   virtual int GetSheetName(int num,LPTSTR lpszSheetName);
	   virtual int GetAllSheetName(CStringArray *strarr);
       virtual VARIANT GetSlectRange(int row,int col);
	   virtual int ClearComment(int nStartRow,int nStartCol );
	   virtual int ClearAllComment();
	   virtual int ClearAllColor();
	   virtual int ClearClolar(int nStartRow,int nStartCol );
	   virtual int GetRow(int row,int nStartCol,int nEndCol,VARIANT &varvalue);
	   virtual int GetCol(int col,int nStartRow,int nEndRow,VARIANT &varvalue);
	   virtual int IsExcelAlradyOpen(LPCTSTR lpszExcelName);
       static int ShowExcel(LPCTSTR lpszExcelName);
	   static int ShowExcel(LPCTSTR lpszExcelName,LPCTSTR lpszSheetName);
	   static int ShowExcel(LPCTSTR lpszExcelName,LPCTSTR lpszSheetName,int row,int col);
protected:
	   int Show(LPCTSTR lpszFileName);
	   int Show(LPCTSTR lpszFileName,LPCTSTR lpszSheetName);
	   int Show(LPCTSTR lpszFileName,LPCTSTR lpszSheetName,int row,int col);
	   int Transition(int row,int col,LPTSTR RCstr);
};
#endif // ExcelSourse_h__