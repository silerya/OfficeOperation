#include "stdafx.h"
#include "ExcelSourse.h"
#include <comutil.h>

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#define new DEBUG_NEW
#endif


ExcelSourse::ExcelSourse()
{
	App = new _Application;
	workbooks = new Workbooks;
	workbook = new _Workbook;
	sheet = new _Worksheet;
	sheets = new Worksheets;
    range = new Range;
	Allrage = new Range;
	Allrage = new Range;
	CoInitialize(NULL);
	if (!App->CreateDispatch(_T("Excel.Application"),NULL))
	{
		AfxMessageBox(_T("创建服务失败"));

        exit(1);
	}
    App->SetScreenUpdating(FALSE);
	workbooks->AttachDispatch(App->GetWorkbooks(),TRUE);
}

ExcelSourse::~ExcelSourse()
{

	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR);
	App->SetScreenUpdating(TRUE);
	Allrage->ReleaseDispatch();
	Usedrage->ReleaseDispatch();
	range->ReleaseDispatch();
	sheet->ReleaseDispatch();
	sheets->ReleaseDispatch();
	workbook->ReleaseDispatch();
	workbooks->ReleaseDispatch();

	if (NULL!=App)
	{
		delete App;
		App = NULL;
	}
	if (NULL!=Allrage)
	{
		delete Allrage;
		Allrage = NULL;
	}
	if (NULL!=Usedrage)
	{
		delete Usedrage;
		Usedrage = NULL;
	}
	if (NULL!=range)
	{
		delete range;
		range = NULL;
	}
	if (NULL!=sheet)
	{
		delete sheet;
		sheet = NULL;
	}
	if (NULL!=sheets)
	{
		delete sheets;
		sheets = NULL;
	}

	if (NULL!=workbook)
	{
		delete workbook;
		workbook = NULL;
	}
	if (NULL!=workbooks)
	{
		delete workbooks;
		workbooks = NULL;
	}
}

int ExcelSourse::SetVisible(bool visible)
{
	App->SetVisible(visible);
	return 0;
}

CString VariantToString(COleVariant *vaData)
{

	CString s;
	switch(vaData->vt)
	{
	case VT_BSTR:
		return CString (vaData->bstrVal);
	case VT_BSTR | VT_BYREF:
		return CString(*vaData->pbstrVal);
	case  VT_I4:
		s.Format(_T("%ld"),vaData->lVal);
	case VT_I4 | VT_BYREF:
		s.Format(_T("%ld"),*vaData->plVal);
	case  VT_R8:
		s.Format(_T("%lf"),vaData->dblVal);
	case  VT_DATE:
		{
			COleDateTime dt(vaData->date);
			s = dt.Format(_T("%Y-%m-%d"));
			return s;
		}

	case  VT_EMPTY:
		return _T("");
	default:
		return CString(*vaData->pbstrVal);


	}
}


int ExcelSourse::Open( LPCTSTR lpszFileName )
{
	if (NULL == lpszFileName)
	{
		return 1;
	}
	else
	{
      workbook->AttachDispatch(workbooks->Add(_variant_t(lpszFileName)));
	}
	sheets->AttachDispatch(workbook->GetWorksheets(),true);

	return 0;
}

int ExcelSourse::SaveAs( LPCTSTR lpszFileName )
{
   CFileFind finder;
   if (finder.FindFile(lpszFileName))
   {
	   if (!DeleteFile(lpszFileName))
	   {
		   return 0;
	   }
   }
   COleVariant covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR);
   workbook->SaveAs(COleVariant(lpszFileName),covOptional,covOptional,covOptional,covOptional,covOptional,1,covOptional,covOptional,covOptional,covOptional);
   workbook->SetSaved(TRUE);
   return 0;
}

int ExcelSourse::Save( LPCTSTR lpszFileName )
{
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR);
	workbook->SaveCopyAs(COleVariant(lpszFileName));
	workbook->SetSaved((TRUE));
	return 0;
}

int ExcelSourse::Save()
{

	workbook->Save();
	workbook->SetSaved(TRUE);
	return 0;
}

int ExcelSourse::Close()
{

	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR);
	Allrage->ReleaseDispatch();
	Usedrage->ReleaseDispatch();
	range->ReleaseDispatch();
	sheet->ReleaseDispatch();
	sheets->ReleaseDispatch();
	workbook->ReleaseDispatch();
	workbooks->ReleaseDispatch();
	workbooks->Close();
	return 0;
}

int ExcelSourse::SetCell( int row,int col,LPCTSTR strValue )
{
	if ((row<=0)||(col<=0))
	{
		return 1;
	}
 
	range->AttachDispatch(sheet->GetCells(),TRUE);
	range->SetItem(_variant_t((long)row),_variant_t((long)col),_variant_t(strValue));
	return 0;
}

int ExcelSourse::SetCell( int row,int col,long lValue )
{
	if ((row<=0)||(col<=0))
	{
		return 1;
	}
	CString t;
	t.Format(_T("%ld"),lValue);
	SetCell(row,col,t);
	return 0;
}

int ExcelSourse::SetCell( int row,int col,int nValue )
{

	if ((row<=0)||(col<=0))
	{
		return 1;
	}
	CString t;
	t.Format(_T("%d"),nValue);
	SetCell(row,col,t);
	return 0;
}

int ExcelSourse::SetCell( int row,int col,double dValue,int n )
{
	if ((row<=0)||(col<=0))
	{
		return 1;
	}
    TCHAR szStr[32];
	_stprintf_s(szStr,30,_T("%.*lf"),n,dValue);
	SetCell(row,col,szStr);
	return 0;
}

int ExcelSourse::SelectActiveSheet( LPCTSTR lpszSheetName )
{
	m_RowNum = 0;
	m_ColNum = 0;
	int i;
	CStringArray *pArray = new CStringArray;
	GetAllSheetName(pArray);
	int sds = pArray->GetSize();

	for (i=0;i<sds;i++)
	{
		CString str = pArray->GetAt(i);
		while (lpszSheetName == str)
		{
			sheet->AttachDispatch(sheets->GetItem(_variant_t(lpszSheetName/*->AllocSysString()*/)),TRUE);
			Allrage->AttachDispatch(sheet->GetCells(),true);

			Usedrage->AttachDispatch(sheet->GetUsedRange(),true);
			delete pArray;
			pArray = NULL;
			m_RowNum = GetAllRow();
			m_ColNum = GetAllCol();
			return 0;
		}
	}
	
	delete pArray;
	pArray = NULL;

	return 1;

}

int ExcelSourse::SelectCurSheet( LPCTSTR lpszSheetName )
{
	return SelectActiveSheet(lpszSheetName);

}

//index start at 1;
int ExcelSourse::SelectCurSheet( int nSheetIndex )
{
	m_RowNum = 0;
	m_ColNum = 0;
 
	if (nSheetIndex <0 || nSheetIndex >GetSheetNum())
	{
		return 1;
	}

	sheet->AttachDispatch(sheets->GetItem(COleVariant(long(nSheetIndex))),true);
	Allrage->AttachDispatch(sheet->GetCells(),true);
	Usedrage->AttachDispatch(sheet->GetUsedRange(),true);
	m_ColNum = GetAllCol();
	m_RowNum = GetAllRow();

	return 0;
}

int ExcelSourse::DeleteSheet( LPCTSTR lpszName )
{
	sheet->Delete();
	return 0;

}

int ExcelSourse::GetCell( int row,int col,LPTSTR strValue )
{

	if ((row <=0)||(col<=0))
	{
		return 1;
	}
	range->AttachDispatch(GetSlectRange(row,col).pdispVal);
	COleVariant varvalue = range->GetValue();
	CString str = VariantToString(&varvalue);
	lstrcpy(strValue,str);
	return 0;
}

int ExcelSourse::GetCell( int row,int col,COleVariant *strValue )
{	
if ((row <=0)||(col<=0))
{
	return 1;
}
range->AttachDispatch(GetSlectRange(row,col).pdispVal);
COleVariant varvalue = range->GetValue();
*strValue = varvalue;
return 0;
}

int ExcelSourse::GetRow( int row,int nStartCol,int nEndCol,CStringArray *strArr )
{

	if ((row <= 0)|| (nStartCol <= 0)||(nEndCol <= 0))
	{
		return 1;
	}
	int numCol = m_ColNum;
	int j;
	if (nEndCol == -1)
	{
		nEndCol = numCol;

	}
	if (nStartCol > nEndCol)
	{
		AfxMessageBox(_T("起始lie不能小于终止lie"));
	}
	for (j = nStartCol;j <= nEndCol;j++)
	{
		GetCell(row,j,m_strRow);
		strArr->Add(m_strRow);

	}

	return 0;
}

int ExcelSourse::GetRow( int row,int nStartCol,int nEndCol,VARIANT &varvalue )
{
	TCHAR str1[32];
	TCHAR str2[32];

	Transition(row,nStartCol,str1);
	Transition(row,nEndCol,str2);

	*range = sheet->GetRange(COleVariant(str1),COleVariant(str2));
	varvalue = range->GetValue();

	return 0;

}

int ExcelSourse::GetCol( int col,int nStartRow,int nEndRow,CStringArray *strArr )
{
		if ((col <= 0)|| (nStartRow <= 0)||(nEndRow <= 0))
		{
		return 1;
		}
		int numrow = m_RowNum;
		int i;
		if (nEndRow == -1)
		{
		   nEndRow = numrow;

		}
		if (nStartRow > nEndRow)
		{
		AfxMessageBox(_T("起始行不能小于终止行"));
		}
		for (i = nStartRow;i <= nEndRow;i++)
		{
		  GetCell(i,col,m_strCol);
		  strArr->Add(m_strCol);

		}

   return 0;

}

int ExcelSourse::GetCol( int col,int nStartRow,int nEndRow,VARIANT &varvalue )
{

	if ((col <= 0)|| (nStartRow <= 0)||(nEndRow <= 0))
	{
		return 1;
	}

	TCHAR szRow1[32];
	TCHAR szRow2[32];
	Transition(nStartRow,col,szRow1);
	Transition(nEndRow,col,szRow2);
	*range = sheet->GetRange(COleVariant(szRow1),COleVariant(szRow2));
	varvalue = range->GetValue();
	return 0;

}

int ExcelSourse::GetColCell( int col,CStringArray *strArr )
{
	if (col<=0)
	{
		return 1;
	}
	int RowNum  = m_RowNum;
	int i;
	for (i = 1;i<RowNum;i++)
	{
		GetCell(i,col,m_strCol);
		strArr->Add(m_strCol);
	}
	return 0;

}

int ExcelSourse::GetRowCell( int row,CStringArray *strArr )
{

	if (row<=0)
	{
		return 1;
	}
	int ColNum  = m_ColNum;
	int i;
	for (i = 1;i<ColNum;i++)
	{
		GetCell(i,row,m_strRow);
		strArr->Add(m_strRow);
	}
	return 0;
}

int ExcelSourse::SetColWidth( int col,int nWidth )
{
	if (col <= 0)
	{
		return 1;
	}
	range->AttachDispatch(GetSlectRange(1,col).pdispVal);
	range->SetRowHeight(_variant_t((long)nWidth));
	return 0;

}

int ExcelSourse::SetRowHeight( int row,int nHeight )
{
	if (row <= 0)
	{
		return 1;
	}
    range->AttachDispatch(GetSlectRange(row,1).pdispVal);
	range->SetRowHeight(_variant_t((long)nHeight));
	return 0;
}

int ExcelSourse::GetCellWidth( int row,int col,int &nWeigth )
{
	if ((row <= 0)||(col < 0))
	{
		return 1;
	}
	range->AttachDispatch(GetSlectRange(row,col).pdispVal);
	COleVariant var = range->GetWidth();

	CString str = VariantToString(&var);
	nWeigth = _ttoi(str);
	return 0;

}

int ExcelSourse::GetCellHeight( int row,int col,int &nHeight )
{

	if ((row <= 0)||(col < 0))
	{
		return 1;
	}
	range->AttachDispatch(GetSlectRange(row,col).pdispVal);
	COleVariant var = range->GetHeight();

	CString str = VariantToString(&var);
	nHeight = _ttoi(str);
	return 0;
}

int ExcelSourse::SetCellPostill( int row,int col,LPCTSTR lpszstr )
{

	if ((row <= 0)||(col < 0))
	{
		return 1;
	}
	range->AttachDispatch(GetSlectRange(row,col).pdispVal);
	Comment com;
	com = range->GetComment();

	if (com.Text(vtMissing,vtMissing,vtMissing)!= _T(""))
	{
		range->ClearComments();
	}

	range->AddComment(_variant_t(lpszstr));

	return 0;
}

int ExcelSourse::SetCellPostill( int row,int col,LPCTSTR lpszstr,LPCTSTR lpszauthor )
{
	if ((row <= 0)||(col < 0))
	{
		return 1;
	}
	range->AttachDispatch(GetSlectRange(row,col).pdispVal);
	Comment com;
	com = range->GetComment();

	if (com.Text(vtMissing,vtMissing,vtMissing)!= _T(""))
	{
		range->ClearComments();
	}
	_stprintf_s(m_str,510,_T("%s :\n%s"),lpszauthor);
	range->AddComment(_variant_t(m_str));
	return 0;
}

int ExcelSourse::ClearCellPostill( int row,int col )
{
	if ((row <= 0)||(col < 0))
	{
		return 1;
	}
	range->AttachDispatch(GetSlectRange(row,col).pdispVal);
	Comment com;
	com = range->GetComment();

	if (com.Text(vtMissing,vtMissing,vtMissing)!= _T(""))
	{
		range->ClearComments();
	}
	return 0;

}

int ExcelSourse::GetCellPostill( int row,int col,LPTSTR lpszstr )
{

	if ((row <= 0)||(col < 0))
	{
		return 1;
	}
	range->AttachDispatch(GetSlectRange(row,col).pdispVal);
	Comment com;
	com = range->GetComment();
	CString strTemp;
	strTemp = com.Text(vtMissing,vtMissing,vtMissing);
	lstrcpy(lpszstr,strTemp);
	return 0;
}

int ExcelSourse::SetCellColor( int row,int col,ExcelColor color )
{

	if ((row <= 0)||(col < 0))
	{
		return 1;
	}
	Interior it;
	range->AttachDispatch(GetSlectRange(row,col).pdispVal);
	it.AttachDispatch(range->GetInterior());
	it.SetColorIndex(_variant_t((long)color));
	return 0;
}

int ExcelSourse::GetCellColor( int row,int col,ExcelColor &color )
{
	if ((row <= 0)||(col < 0))
	{
		return 1;
	}
	Interior it;
	range->AttachDispatch(GetSlectRange(row,col).pdispVal);
	it.AttachDispatch(range->GetInterior());
	COleVariant var = it.GetColorIndex();

	LONG colorindex = _ttol(VariantToString(&var));
	switch(colorindex)
	{
	case 1:
		color = black;
		break;
	case 2:
		color = white;
		break;
	case 3:
		color = red;
		break;
	case 4:
		color = green;
		break;
	case 5:
		color = blue;
		break;
	case 6:
		color = yellow;
		break;
	case -4142:
		color = NullColor;
		break;
	}

	return 0;
}

int ExcelSourse::SetRowColor( int row,ExcelColor color )
{
	Interior it;
	TCHAR str1[32];
	TCHAR str2[32];
	Transition(row,m_ColNum,str2);
    Transition(row,m_ColNum,str2);
	*range = sheet->GetRange(COleVariant(str1),COleVariant(str2));

	it.AttachDispatch(range->GetInterior());
	it.SetColorIndex(COleVariant((long)color));

	return 0;


}

int ExcelSourse::GetAllRow()
{
	int RowNum = 0;
	Range range_old;
	Range usedRange;
	usedRange.AttachDispatch(sheet->GetUsedRange(),true);
	range_old.AttachDispatch(usedRange.GetRows(),true);
	RowNum = range_old.GetCount();

	usedRange.ReleaseDispatch();
	range_old.ReleaseDispatch();

	return RowNum;

}

int ExcelSourse::GetAllCol()
{

	int ColNum = 0;
	Range range_old;
	Range usedRange;
	usedRange.AttachDispatch(sheet->GetUsedRange(),true);
	range_old.AttachDispatch(usedRange.GetColumns(),true);
	ColNum = range_old.GetCount();

	usedRange.ReleaseDispatch();
	range_old.ReleaseDispatch();

	return ColNum;
}

int ExcelSourse::GetSheetNum()
{
	*sheets = workbook->GetWorksheets();
	int num = sheets->GetCount();

	return num;

}

int ExcelSourse::GetSheetName( int num,LPTSTR lpszSheetName )
{
	if (num <= 0)
	{
		return 1;
	}

	CString da;
	*sheet = sheets->GetItem(_variant_t((long)num));
	da = sheet->GetName();
	lstrcpy(lpszSheetName,da);
	return 0;

}

int ExcelSourse::GetAllSheetName( CStringArray *strarr )
{
	CString da;
	int num = GetSheetNum();
    
	for (int i = 1;i <= num;i++)
	{
		*sheet = sheets->GetItem(_variant_t((long)i));
		da = sheet->GetName();
		strarr->Add(da);
	}
	return 0;
}

VARIANT ExcelSourse::GetSlectRange( int row,int col )
{
	VARIANT allrange;
	if (!((row <= 0)||(col < 0)))
	{
		allrange = Allrage->GetItem(COleVariant((long)row),COleVariant((long)col));
		return allrange;
	}
	allrange.vt = VT_EMPTY;

	return allrange;

}

int ExcelSourse::ClearComment( int nStartRow,int nStartCol )
{
	int nEndrow = m_RowNum;
	TCHAR szRow1[32];
	Transition(nStartRow,nStartCol,szRow1);

	int col2 = m_ColNum;
	TCHAR szRow2[32];
	Transition(nEndrow,col2,szRow2);
    *range = sheet->GetRange(COleVariant(szRow1),COleVariant(szRow2));
	range->ClearComments();

	return 0;

}

int ExcelSourse::ClearAllComment()
{
	Allrage->ClearComments();
	return 0;	

}

int ExcelSourse::ClearAllColor()
{
	Interior it;
	*range = *Allrage;
	it.AttachDispatch(range->GetInterior());
	it.SetColorIndex(COleVariant((long)NullColor));

	return 0;

}

int ExcelSourse::ClearClolar( int nStartRow,int nStartCol )
{
	int nEndrow = m_RowNum;
	TCHAR szRow1[32];
	Transition(nStartRow,nStartCol,szRow1);

	int col2 = m_ColNum;
	TCHAR szRow2[32];
	Transition(nEndrow,col2,szRow2);
	 *range = sheet->GetRange(COleVariant(szRow1),COleVariant(szRow2));
	 Interior it;

	 it.AttachDispatch(range->GetInterior());
	 it.SetColorIndex(COleVariant((long)NullColor));

	 return 0;


}

int ExcelSourse::IsExcelAlradyOpen( LPCTSTR lpszExcelName )
{
	return 0;//IsAlradyOpen(lpszExcelName);

}



int ExcelSourse::ShowExcel( LPCTSTR lpszExcelName,LPCTSTR lpszSheetName,int row,int col )
{
	if ((row <=0)||(col<=0))
	{
		return 1;
	}

	ExcelSourse excel;
	excel.Show(lpszExcelName,lpszSheetName,row,col);
	return 0;
}

int ExcelSourse::ShowExcel( LPCTSTR lpszExcelName,LPCTSTR lpszSheetName )
{
	ExcelSourse excel;
	excel.Show(lpszExcelName,lpszSheetName);

	return 0;
}

int ExcelSourse::ShowExcel( LPCTSTR lpszExcelName )
{
	ExcelSourse excel;
	excel.Show(lpszExcelName);

	return 0;

}


int ExcelSourse::Transition( int row,int col,LPTSTR RCstr )
{
	if ((row<=0)||(col<=0)||(col>256))
	{
		return 1;
	}
	TCHAR szRow1[32];
	if (col > 26)
	{
		_stprintf_s(szRow1,30,_T("%c%c%d"),'A' +(col - 1)/26 - 1,(col - 1)%26 + 'A',row);
	} 
	else
	{
		_stprintf_s(szRow1,30,_T("%c%d"),col - 1 + 'A',row);
	}

	lstrcpy(RCstr,szRow1);
	return 0;

}

int ExcelSourse::Show( LPCTSTR lpszFileName,LPCTSTR lpszSheetName,int row,int col )
{
	TCHAR szRow1[32];
	Transition(row,col,szRow1);
	*workbooks = App->GetWorkbooks();
	*workbook = workbooks->Add(_variant_t(lpszFileName));
	sheets->AttachDispatch(workbook->GetWorksheets(),true);
	sheet->AttachDispatch(sheets->GetItem(_variant_t(lpszSheetName)),true);
	sheet->Activate();
	sheet->Select(_variant_t(true));

    *range = sheet->GetRange(COleVariant(szRow1),COleVariant(szRow1));

	range->Activate();
	range->Select();
	App->Goto(_variant_t(App->GetActiveCell()),_variant_t(true));
	range->Show();
	App->SetVisible(TRUE);
	return 0;
}

int ExcelSourse::Show( LPCTSTR lpszFileName,LPCTSTR lpszSheetName )
{
	*workbooks = App->GetWorkbooks();
	*workbook = workbooks->Add(_variant_t(lpszFileName));
	sheets->AttachDispatch(workbook->GetWorksheets(),true);
	sheet->AttachDispatch(sheets->GetItem(_variant_t(lpszSheetName)),true);
	sheet->Activate();
	sheet->Select(_variant_t(true));

	App->SetVisible(TRUE);

	return 0;

}

int ExcelSourse::Show( LPCTSTR lpszFileName )
{
	*workbooks = App->GetWorkbooks();
	*workbook = workbooks->Add(_variant_t(lpszFileName));
    App->SetVisible(TRUE);
	return 0;
}




