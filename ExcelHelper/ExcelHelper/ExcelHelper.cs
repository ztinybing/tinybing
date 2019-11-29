using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using NPOI.HSSF.UserModel;
using System.IO;
using System.Runtime.InteropServices;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Com.Bing
{
    public class ExcelHelper
    {
        /// <summary>
        /// 电子表格导入到DataTable,默认第一个表单,第一行不作为DataTable的列名
        /// </summary>
        /// <param name="xlsPath">excel路径</param>
        /// <returns>结果DataTable</returns>
        public static DataTable ImportFromExcel(string xlsPath)
        {
            ExcelHelper helper = new ExcelHelper(xlsPath);
            DataTable dt = helper.ExcelToDataTable();
            return dt;
        }
        /// <summary>
        /// 电子表格导入到DataTable
        /// </summary>
        /// <param name="xlsPath">excel路径</param>
        /// <param name="sheetName">表单名称，未找到则默认取第一个表单</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>结果DataTable</returns>
        public static DataTable ImportFromExcel(string xlsPath, string sheetName, bool isFirstRowColumn)
        {
            ExcelHelper helper = new ExcelHelper(xlsPath);
            DataTable dt = helper.ExcelToDataTable(sheetName, isFirstRowColumn);
            return dt;
        }
        public static DataSet ImportDataSetFromExcel(string xlsPath)
        {
            return ImportDataSetFromExcel(xlsPath, false);
        }
        public static DataSet ImportDataSetFromExcel(string xlsPath, bool isFirstRowColumn)
        {
            ExcelHelper helper = new ExcelHelper(xlsPath);
            return helper.ExcelToDataSet(isFirstRowColumn);
        }
        /// <summary>
        /// 电子表格导入到DataTable，默认第一个表单
        /// </summary>
        /// <param name="xlsPath">excel路径</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>结果DataTable</returns>
        public static DataTable ImportFromExcel(string xlsPath, bool isFirstRowColumn)
        {
            ExcelHelper helper = new ExcelHelper(xlsPath);
            DataTable dt = helper.ExcelToDataTable(isFirstRowColumn);
            return dt;
        }

        public static int ExprotToExcel(string xlsPath, DataTable dt, bool isColumnWritten)
        {
            ExcelHelper helper = new ExcelHelper(xlsPath);
            return helper.DataTableToExcel(dt, isColumnWritten);
        }
        public static int ExprotToExcel(string xlsPath, DataSet ds, bool isColumnWritten)
        {
            ExcelHelper helper = new ExcelHelper(xlsPath);
            return helper.DataSetToExcel(ds, isColumnWritten);
        }

        private string fileName = null; //文件名
        private ExcelHelper(string fileName)
        {
            this.fileName = fileName;
        }
        /// <summary>
        /// 将DataSet数据导出到excel中
        /// </summary>
        /// <param name="data">要导入的数据</param>
        /// <param name="isColumnWritten">DataTable的列名是否要导入</param>
        /// <returns>导入数据行数(包含列名那一行)</returns>
        public int DataSetToExcel(DataSet ds, bool isColumnWritten)
        {
            int allCount = 0;
            try
            {
                IWorkbook workbook = CreateWorkBook(fileName);
                if (workbook == null) return -1;
                foreach (DataTable data in ds.Tables)
                {
                    int count = 0;
                    ISheet sheet = workbook.CreateSheet(data.TableName);

                    if (isColumnWritten) //写入DataTable的列名
                    {
                        IRow row = sheet.CreateRow(0);
                        for (int j = 0; j < data.Columns.Count; j++)
                        {
                            row.CreateCell(j).SetCellValue(data.Columns[j].ColumnName);
                        }
                        count = 1;
                    }
                    else
                    {
                        count = 0;
                    }

                    for (int i = 0; i < data.Rows.Count; i++)
                    {
                        IRow row = sheet.CreateRow(count);
                        for (int j = 0; j < data.Columns.Count; ++j)
                        {
                            row.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
                        }
                        count++;
                    }
                    allCount += count;
                }
                using (FileStream fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    workbook.Write(fs); //写入到excel
                }
                return allCount;

            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return -1;
            }

        }
        /// <summary>
        /// 将DataTable数据导出到excel中
        /// </summary>
        /// <param name="data">要导入的数据</param>
        /// <param name="isColumnWritten">DataTable的列名是否要导入</param>
        /// <returns>导入数据行数(包含列名那一行)</returns>
        public int DataTableToExcel(DataTable data, bool isColumnWritten)
        {
            DataSet ds = new DataSet();
            ds.Tables.Add(data);
            return DataSetToExcel(ds, isColumnWritten);
        }

        private IWorkbook CreateWorkBook(string fileName)
        {
            if (fileName.IndexOf(".xlsx") > 0) return new XSSFWorkbook();// 2007版本
            else if (fileName.IndexOf(".xls") > 0) return new HSSFWorkbook(); // 2003版本
            return null;
        }

        #region Excel2DataTable
        /// <summary>
        /// 将excel中的第一张表单数据导入到DataTable中
        /// </summary>
        /// <returns>返回的DataTable</returns>
        public DataTable ExcelToDataTable()
        {
            return ExcelToDataTable(null, false);
        }
        /// <summary>
        /// 将excel中的第一张表单数据导入到DataTable中
        /// </summary>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>返回的DataTable</returns>
        public DataTable ExcelToDataTable(bool isFirstRowColumn)
        {
            return ExcelToDataTable(null, isFirstRowColumn);
        }
        /// <summary>
        /// 将excel中的数据导入到DataSet中
        /// </summary>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>返回的DataSet</returns>
        public DataSet ExcelToDataSet(bool isFirstRowColumn)
        {
            try
            {
                DataSet ds = new DataSet();
                using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = GetWorkBook(fs, fileName);
                    if (workbook == null) return ds;
                    for (int t = 0; t < workbook.Count; t++)
                    {
                        ISheet sheet = workbook.GetSheetAt(t);

                        DataTable dt = new DataTable();

                        dt.TableName = sheet.SheetName;//设置表名

                        int startRow = 0;//数据起始行

                        int maxColumnNum = GetMaxColumnNum(startRow, sheet);//最大列数

                        if (isFirstRowColumn)
                        {
                            IRow firstRow = sheet.GetRow(sheet.FirstRowNum);
                            for (int i = 0; i < maxColumnNum; i++)
                            {
                                ICell cell = firstRow.GetCell(i);
                                string cellValue = string.Empty;
                                if (cell != null) cellValue = GetCellValue(cell).ToString();
                                if (string.IsNullOrEmpty(cellValue)) cellValue = GetColumnLetter(i + 1);
                                else if (dt.Columns.Contains(cellValue))
                                {
                                    cellValue = string.Format("[{0}]重复列{1}", GetColumnLetter(i + 1), cellValue);
                                }
                                DataColumn column = new DataColumn(cellValue);
                                dt.Columns.Add(column);
                            }
                            startRow = sheet.FirstRowNum + 1;
                        }
                        else
                        {
                            startRow = sheet.FirstRowNum;
                            for (int i = 1; i <= maxColumnNum; i++)
                            {
                                dt.Columns.Add(GetColumnLetter(i));
                            }
                        }

                        //最后一列的标号
                        for (int i = startRow; i <= sheet.LastRowNum; i++)
                        {
                            IRow row = sheet.GetRow(i);

                            DataRow dataRow = dt.NewRow();
                            dt.Rows.Add(dataRow);

                            if (row == null) continue;//没有数据的行默认是null　

                            for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                            {
                                ICell cell = row.GetCell(j);
                                if (cell == null) continue; //同理，没有数据的单元格都默认是null
                                dataRow[j] = GetCellValue(cell);
                            }
                        }
                        ds.Tables.Add(dt);
                    }
                }
                return ds;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                DataSet ds = new DataSet();
                DataTable errorDt = new DataTable();
                errorDt.Columns.Add("error");
                DataRow row = errorDt.NewRow();
                row["error"] = ex.Message;
                errorDt.Rows.Add(row);
                ds.Tables.Add(errorDt);
                return ds;
            }
        }
        /// <summary>
        /// 将excel中的数据导入到DataTable中
        /// </summary>
        /// <param name="sheetName">excel工作薄sheet的名称</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>返回的DataTable</returns>
        public DataTable ExcelToDataTable(string sheetName, bool isFirstRowColumn)
        {
            try
            {
                using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    DataTable dt = new DataTable();

                    IWorkbook workbook = GetWorkBook(fs, fileName);
                    if (workbook == null) return dt;
                    ISheet sheet = GetSheet(workbook, sheetName);
                    if (sheet == null) return dt;


                    dt.TableName = sheet.SheetName;//设置表名

                    int startRow = 0;//数据起始行

                    int maxColumnNum = GetMaxColumnNum(startRow, sheet);//最大列数

                    if (isFirstRowColumn)
                    {
                        IRow firstRow = sheet.GetRow(sheet.FirstRowNum);
                        for (int i = 0; i < maxColumnNum; i++)
                        {
                            ICell cell = firstRow.GetCell(i);
                            string cellValue = string.Empty;
                            if (cell != null) cellValue = GetCellValue(cell).ToString();
                            if (string.IsNullOrEmpty(cellValue)) cellValue = GetColumnLetter(i + 1);
                            else if (dt.Columns.Contains(cellValue))
                            {
                                cellValue = string.Format("[{0}]重复列{1}", GetColumnLetter(i + 1), cellValue);
                            }
                            DataColumn column = new DataColumn(cellValue);
                            dt.Columns.Add(column);
                        }
                        startRow = sheet.FirstRowNum + 1;
                    }
                    else
                    {
                        startRow = sheet.FirstRowNum;
                        for (int i = 1; i <= maxColumnNum; i++)
                        {
                            dt.Columns.Add(GetColumnLetter(i));
                        }
                    }

                    //最后一列的标号
                    for (int i = startRow; i <= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);

                        DataRow dataRow = dt.NewRow();
                        dt.Rows.Add(dataRow);

                        if (row == null) continue;//没有数据的行默认是null　

                        for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                        {
                            ICell cell = row.GetCell(j);
                            if (cell == null) continue; //同理，没有数据的单元格都默认是null
                            dataRow[j] = GetCellValue(cell);
                        }
                    }
                    return dt;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                DataTable errorDt = new DataTable();
                errorDt.Columns.Add("error");
                DataRow row = errorDt.NewRow();
                row["error"] = ex.Message;
                errorDt.Rows.Add(row);
                return errorDt;
            }
        }
        private IWorkbook GetWorkBook(FileStream fs, string fileName)
        {
            // 2007版本
            if (fileName.IndexOf(".xlsx") > 0 || fileName.IndexOf(".xlsm") > 0) return new XSSFWorkbook(fs);
            // 2003版本
            else if (fileName.IndexOf(".xls") > 0) return new HSSFWorkbook(fs);
            return null;
        }
        private ISheet GetSheet(IWorkbook workbook, string sheetName)
        {
            if (sheetName != null)
            {
                ISheet sheet = workbook.GetSheet(sheetName);
                if (sheet != null) return sheet;
                //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                return workbook.GetSheetAt(0);
            }
            else
            {
                return workbook.GetSheetAt(0);
            }
        }
        private object GetCellValue(ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.String: return cell.StringCellValue;
                case CellType.Numeric:
                    {
                        if (HSSFDateUtil.IsCellDateFormatted(cell)) return cell.DateCellValue;
                        return cell.NumericCellValue;
                    }
                case CellType.Boolean: return cell.BooleanCellValue;
                case CellType.Formula: return cell.NumericCellValue;
                case CellType.Blank: return string.Empty;
                default: return "ERROR";
            }
        }

        private int GetMaxColumnNum(int startRowIndex, ISheet sheet)
        {
            int maxColumnNum = 0;
            for (int i = startRowIndex; i <= sheet.LastRowNum; ++i)
            {
                IRow row = sheet.GetRow(i);
                if (row == null) continue;
                maxColumnNum = Math.Max(row.LastCellNum, maxColumnNum);
            }
            return maxColumnNum;
        }
        #endregion

        private Dictionary<int, string> columnLetterDict = new Dictionary<int, string>();
        private string GetColumnLetter(int columnIndex)
        {
            if (columnLetterDict.ContainsKey(columnIndex)) return columnLetterDict[columnIndex];
            Stack<char> stack = new Stack<char>();
            int curColumnIndex = columnIndex;
            while (curColumnIndex > 0)
            {
                int remain = curColumnIndex / 26;
                int remainder = curColumnIndex % 26;
                if (remainder == 0)
                {
                    remain -= 1;
                    remainder = 26;
                }
                stack.Push((char)(remainder + 64));
                curColumnIndex = remain;
            }
            StringBuilder sb = new StringBuilder();
            while (stack.Count > 0) sb.Append(stack.Pop());
            columnLetterDict[columnIndex] = sb.ToString();
            return sb.ToString();
        }
    }
}
