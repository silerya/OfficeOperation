OfficeOperation
===============

对offile包括Excel和Word的相关的，版面设置，数据读写，格式设置，背景设置，等一系列的操作

详细：
           ////////////////////////////////////////////////
           /// \fn     virtual int SetVisible(bool visible);
           /// \brief  设置excel可见与否
           /// \param  [IN]  bool visible  --true为可见
           /// \author siler_ya
           /// \return 成功返回 0
           ////////////////////////////////////////////////
	   virtual int SetVisible(bool visible);

	   ////////////////////////////////////////////////
	   /// \fn     virtual int Open(LPCTSTR lpszFileName);
	   /// \brief  打开excel
	   /// \param  [IN]  LPCTSTR lpszFileName  --excel名称(带路径) 
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int Open(LPCTSTR lpszFileName);

	   ////////////////////////////////////////////////
	   /// \fn     virtual int SaveAs(LPCTSTR lpszFileName);
	   /// \brief  另存excel
	   /// \param  [IN]  LPCTSTR lpszFileName  --excel名称(带路径) 
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int SaveAs(LPCTSTR lpszFileName);
	   ////////////////////////////////////////////////
	   /// \fn     virtual int Save(LPCTSTR lpszFileName);
	   /// \brief  保存excel
	   /// \param  [IN]  LPCTSTR lpszFileName  --excel名称(带路径) 
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int Save(LPCTSTR lpszFileName);
	   ////////////////////////////////////////////////
	   /// \fn     virtual int Save();
	   /// \brief  保存excel
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int Save();

	   ////////////////////////////////////////////////
	   /// \fn     virtual int Close();
	   /// \brief  关闭excel
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int Close();

	   ////////////////////////////////////////////////
	   /// \fn     virtual int SetCell(int row,int col,LPCTSTR strValue);
	   /// \brief  设置单元格内容
	   /// \param  [IN]  int row                 --行号
	   /// \param  [IN]  int col                 --列号
	   /// \param  [IN]  LPCTSTR strValue        --值
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int SetCell(int row,int col,LPCTSTR strValue);
	   ////////////////////////////////////////////////
	   /// \fn     virtual int SetCell(int row,int col,long lValue);
	   /// \brief  设置单元格内容
	   /// \param  [IN]  int row                 --行号
	   /// \param  [IN]  int col                 --列号
	   /// \param  [IN]  long lValue             --值
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int SetCell(int row,int col,long lValue);
	   ////////////////////////////////////////////////
	   /// \fn     virtual int SetCell(int row,int col,int nValue);
	   /// \brief  设置单元格内容
	   /// \param  [IN]  int row                 --行号
	   /// \param  [IN]  int col                 --列号
	   /// \param  [IN]  int nValue              --值
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int SetCell(int row,int col,int nValue);
	   ////////////////////////////////////////////////
	   /// \fn     virtual int SetCell(int row,int col,LPCTSTR strValue);
	   /// \brief  设置单元格内容
	   /// \param  [IN]  int row                 --行号
	   /// \param  [IN]  int col                 --列号
	   /// \param  [IN]  double dValue           --值
	   /// \param  [IN]  int n                   --精确度
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int SetCell(int row,int col,double dValue,int n);
	   //virtual int SetSheetName(LPCTSTR lpszName);

	   ////////////////////////////////////////////////
	   /// \fn     virtual int SelectActiveSheet(LPCTSTR lpszSheetName);
	   /// \brief  设置活动的sheet
	   /// \param  [IN] LPCTSTR lpszSheetName    --sheet名称
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int SelectActiveSheet(LPCTSTR lpszSheetName);

	   ////////////////////////////////////////////////
	   /// \fn     virtual int SelectCurSheet(LPCTSTR lpszSheetName);
	   /// \brief  设置活动的sheet
	   /// \param  [IN] LPCTSTR lpszSheetName    --sheet名称
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int SelectCurSheet(LPCTSTR lpszSheetName);

	   ////////////////////////////////////////////////
	   /// \fn     virtual int SelectCurSheet(int nSheetIndex);
	   /// \brief  设置活动的sheet
	   /// \param  [IN] int nSheetIndex         --sheet序号
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int SelectCurSheet(int nSheetIndex);

	   ////////////////////////////////////////////////
	   /// \fn     virtual int DeleteSheet(LPCTSTR lpszName);
	   /// \brief  删除sheet
	   /// \param  [IN] LPCTSTR lpszName        --sheet名称
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int DeleteSheet(LPCTSTR lpszName);

	   ////////////////////////////////////////////////
	   /// \fn     virtual int GetCell(int row,int col,LPTSTR strValue);
	   /// \brief  获取单元格内容
	   /// \param  [IN]  int row                 --行号
	   /// \param  [IN]  int col                 --列号
	   /// \param  [OUT] LPTSTR strValue         --值
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int GetCell(int row,int col,LPTSTR strValue);

	   ////////////////////////////////////////////////
	   /// \fn     virtual int GetCell(int row,int col,COleVariant *strValue);
	   /// \brief  获取单元格内容
	   /// \param  [IN]  int row                 --行号
	   /// \param  [IN]  int col                 --列号
	   /// \param  [OUT] COleVariant *strValue   --值
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int GetCell(int row,int col,COleVariant *strValue);

	   ////////////////////////////////////////////////
	   /// \fn      virtual int GetRow(int row,int nStartCol,int nEndCol,CStringArray *strArr);
	   /// \brief  获取指定行内容（指定起始列和终止列）
	   /// \param  [IN]  int row                 --行号
	   /// \param  [IN]  int nStartCol           --起始列号
	   /// \param  [IN]  int nEndCol             --终止列号
	   /// \param  [OUT] CStringArray *strArr    --值
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int GetRow(int row,int nStartCol,int nEndCol,CStringArray *strArr);

	   ////////////////////////////////////////////////
	   /// \fn     virtual int GetRow(int row,int nStartCol,int nEndCol,CStringArray *strArr);
	   /// \brief  获取指定列内容（指定起始行和终止行）
	   /// \param  [IN]  int col                 --列号
	   /// \param  [IN]  int nStartRow           --起始行号
	   /// \param  [IN]  int nEndRow             --终止行号
	   /// \param  [OUT] CStringArray *strArr    --值
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int GetCol(int col,int nStartRow,int nEndRow,CStringArray *strArr);

	   ////////////////////////////////////////////////
	   /// \fn     virtual int GetColCell(int col,CStringArray *strArr);
	   /// \brief  获取指定列内容
	   /// \param  [IN]  int col                 --列号
	   /// \param  [OUT] CStringArray *strArr    --值
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int GetColCell(int col,CStringArray *strArr);

	   ////////////////////////////////////////////////
	   /// \fn     virtual int GetRowCell(int row,CStringArray *strArr);
	   /// \brief  获取指定行内容
	   /// \param  [IN]  int row                 --行号
	   /// \param  [OUT] CStringArray *strArr    --值
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int GetRowCell(int row,CStringArray *strArr);
	   
	   ////////////////////////////////////////////////
	   /// \fn     virtual int SetColWidth(int col,int nWidth);
	   /// \brief  设置列宽
	   /// \param  [IN]  int col           --列号 
	   /// \param  [IN]  int nWidth        --列宽
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int SetColWidth(int col,int nWidth);

	   ////////////////////////////////////////////////
	   /// \fn    virtual int SetRowHeight(int row,int nHeight);
	   /// \brief  设置行高
	   /// \param  [IN]  int row           --行号 
	   /// \param  [IN]  int nWidth        --行高
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int SetRowHeight(int row,int nHeight);

	   ////////////////////////////////////////////////
	   /// \fn      virtual int GetCellWidth(int row,int col,int &nWeigth);
	   /// \brief  获取单元格宽度
	   /// \param  [IN]  int row           --行号 
	   /// \param  [IN]  int col           --列号
	   /// \param  [OUT] int &nWeigth      --宽度
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int GetCellWidth(int row,int col,int &nWeigth);

	   ////////////////////////////////////////////////
	   /// \fn      virtual int GetCellHeight(int row,int col,int &nHeight);
	   /// \brief  获取单元格高度
	   /// \param  [IN]  int row           --行号 
	   /// \param  [IN]  int col           --列号
	   /// \param  [OUT] int &nHeight      --高度
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int GetCellHeight(int row,int col,int &nHeight);

	   ////////////////////////////////////////////////
	   /// \fn     virtual int SetCellPostill(int row,int col,LPCTSTR lpszstr);
	   /// \brief  设置单元格批注
	   /// \param  [IN]  int row           --行号 
	   /// \param  [IN]  int col           --列号
	   /// \param  [IN]  LPCTSTR lpszstr   --批注
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
       virtual int SetCellPostill(int row,int col,LPCTSTR lpszstr);

	   ////////////////////////////////////////////////
	   /// \fn     virtual int SetCellPostill(int row,int col,LPCTSTR lpszstr);
	   /// \brief  设置单元格批注
	   /// \param  [IN]  int row           --行号 
	   /// \param  [IN]  int col           --列号
	   /// \param  [IN]  LPCTSTR lpszstr   --批注
	   /// \param  [IN]  LPCTSTR lpszauthor--作者
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int SetCellPostill(int row,int col,LPCTSTR lpszstr,LPCTSTR lpszauthor);

	   ////////////////////////////////////////////////
	   /// \fn     virtual int ClearCellPostill(int row,int col);
	   /// \brief  清空单元格批注
	   /// \param  [IN]  int row           --行号 
	   /// \param  [IN]  int col           --列号
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int ClearCellPostill(int row,int col);

	   ////////////////////////////////////////////////
	   /// \fn    virtual int GetCellPostill(int row,int col,LPTSTR lpszstr);
	   /// \brief  获取单元格批注
	   /// \param  [IN]  int row           --行号 
	   /// \param  [IN]  int col           --列号
	   /// \param  [OUT] LPTSTR lpszstr    --批注内容
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int GetCellPostill(int row,int col,LPTSTR lpszstr);

	   ////////////////////////////////////////////////
	   /// \fn     virtual int SetCellColor(int row,int col,ExcelColor color);
	   /// \brief  设置单元格颜色
	   /// \param  [IN]  int row           --行号 
	   /// \param  [IN]  int col           --列号
	   /// \param  [IN]  ExcelColor colorr  --颜色
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int SetCellColor(int row,int col,ExcelColor color);

	   ////////////////////////////////////////////////
	   /// \fn     virtual int GetCellColor(int row,int col,ExcelColor &color);
	   /// \brief  获取单元格颜色
	   /// \param  [IN]  int row             --行号 
	   /// \param  [IN]  int col             --列号
	   /// \param  [OUT] ExcelColor &colorr  --颜色
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int GetCellColor(int row,int col,ExcelColor &color);

	   ////////////////////////////////////////////////
	   /// \fn     virtual int SetRowColor(int row,ExcelColor color);
	   /// \brief  设置整行填充颜色
	   /// \param  [IN]  int row            --行号 
	   /// \param  [IN]  ExcelColor colorr  --颜色
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int SetRowColor(int row,ExcelColor color);
	  
	   ////////////////////////////////////////////////
	   /// \fn     virtual int GetAllRow();
	   /// \brief  获取已用的行数
	   /// \author siler_ya
	   /// \return int 返回行数
	   ////////////////////////////////////////////////
	   virtual int GetAllRow();

	   ////////////////////////////////////////////////
	   /// \fn     virtual int GetAllCol();
	   /// \brief  获取已用的列数
	   /// \author siler_ya
	   /// \return int 返回列数
	   ////////////////////////////////////////////////
	   virtual int GetAllCol();

	   ////////////////////////////////////////////////
	   /// \fn     virtual int GetSheetNum();
	   /// \brief  获取sheet数目
	   /// \author siler_ya
	   /// \return int 返回sheet数目
	   ////////////////////////////////////////////////
	   virtual int GetSheetNum();

	   ////////////////////////////////////////////////
	   /// \fn     virtual int GetSheetName(int num,LPTSTR lpszSheetName);
	   /// \brief  获取指定索引号的sheet名称 
	   /// \param  [IN]  int num                 --索引号
	   /// \param  [OUT] LPTSTR lpszSheetName    --sheet名称
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int GetSheetName(int num,LPTSTR lpszSheetName);

	   ////////////////////////////////////////////////
	   /// \fn     virtual int GetAllSheetName(CStringArray *strarr);
	   /// \brief  获取所有的sheet名称
	   /// \param  [OUT] CStringArray *strarr   --sheet名称
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int GetAllSheetName(CStringArray *strarr);

	   ////////////////////////////////////////////////
	   /// \fn     virtual VARIANT GetSlectRange(int row,int col);
	   /// \brief  选择指定range
	   /// \param  [IN]  int row  --行
	   /// \param  [IN]  int col  --列
	   /// \author siler_ya
	   /// \return VARIANT
	   ////////////////////////////////////////////////
       virtual VARIANT GetSlectRange(int row,int col);

	   ////////////////////////////////////////////////
	   /// \fn     virtual int ClearComment(int nStartRow,int nStartCol );
	   /// \brief  清空指定起始行列的批注 
	   /// \param  [IN]  int nStartRow    --起始行
	   /// \param  [IN]  int nStartCol    --起始列
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int ClearComment(int nStartRow,int nStartCol );

	   ////////////////////////////////////////////////
	   /// \fn     virtual int ClearAllComment();
	   /// \brief  清空所有批注
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int ClearAllComment();

	   ////////////////////////////////////////////////
	   /// \fn     virtual int ClearAllColor();
	   /// \brief  清空所有填充颜色
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int ClearAllColor();

	   ////////////////////////////////////////////////
	   /// \fn     virtual int ClearClolar(int nStartRow,int nStartCol );
	   /// \brief  清空指定起始行列的颜色 
	   /// \param  [IN]  int nStartRow    --起始行
	   /// \param  [IN]  int nStartCol    --起始列
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int ClearClolar(int nStartRow,int nStartCol );

	   ////////////////////////////////////////////////
	   /// \fn     virtual int GetRow(int row,int nStartCol,int nEndCol,VARIANT &varvalue);
	   /// \brief  获取指定行内容（指定起始列和终止列）
	   /// \param  [IN]  int row                 --行号
	   /// \param  [IN]  int nStartCol           --起始列号
	   /// \param  [IN]  int nEndCol             --终止列号
	   /// \param  [OUT] VARIANT &varvalue       --值
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int GetRow(int row,int nStartCol,int nEndCol,VARIANT &varvalue);

	   ////////////////////////////////////////////////
	   /// \fn     virtual int GetCol(int row,int nStartCol,int nEndCol,VARIANT &varvalue);
	   /// \brief  获取指定列内容（指定起始行和终止行）
	   /// \param  [IN]  int row                 --行号
	   /// \param  [IN]  int nStartCol           --起始列号
	   /// \param  [IN]  int nEndCol             --终止列号
	   /// \param  [OUT] VARIANT &varvalue       --值
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int GetCol(int col,int nStartRow,int nEndRow,VARIANT &varvalue);

	   ////////////////////////////////////////////////
	   /// \fn     virtual int IsExcelAlradyOpen(LPCTSTR lpszExcelName);
	   /// \brief  判断excel是否被占用 
	   /// \param  [IN]  LPCTSTR lpszExcelName       --excel名字（带路径）
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   virtual int IsExcelAlradyOpen(LPCTSTR lpszExcelName);

	   ////////////////////////////////////////////////
	   /// \fn     static int ShowExcel(LPCTSTR lpszExcelName);
	   /// \brief  显示excel
	   /// \param  [IN]  LPCTSTR lpszExcelName       --excel名字（带路径）
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
       static int ShowExcel(LPCTSTR lpszExcelName);

	   ////////////////////////////////////////////////
	   /// \fn     static int Showstatic int ShowExcel(LPCTSTR lpszExcelName,LPCTSTR lpszSheetName);
	   /// \brief  显示excel到指定sheet
	   /// \param  [IN]  LPCTSTR lpszExcelName       --excel名字（带路径）
	   /// \param  [IN]  LPCTSTR lpszSheetName       --sheet名字
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   static int ShowExcel(LPCTSTR lpszExcelName,LPCTSTR lpszSheetName);

	   ////////////////////////////////////////////////
	   /// \fn     static int ShowExcel(LPCTSTR lpszExcelName,LPCTSTR lpszSheetName,int row,int col);
	   /// \brief  显示excel到指定单元格
	   /// \param  [IN]  LPCTSTR lpszExcelName       --excel名字（带路径）
	   /// \param  [IN]  LPCTSTR lpszSheetName       --sheet名字
	   /// \param  [IN]  int row                     --行
	   /// \param  [IN] int co                       --列
	   /// \author siler_ya
	   /// \return 成功返回 0
	   ////////////////////////////////////////////////
	   static int ShowExcel(LPCTSTR lpszExcelName,LPCTSTR lpszSheetName,int row,int col);
