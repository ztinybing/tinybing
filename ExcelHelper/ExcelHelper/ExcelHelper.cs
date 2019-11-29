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
        /// ���ӱ���뵽DataTable,Ĭ�ϵ�һ����,��һ�в���ΪDataTable������
        /// </summary>
        /// <param name="xlsPath">excel·��</param>
        /// <returns>���DataTable</returns>
        public static DataTable ImportFromExcel(string xlsPath)
        {
            ExcelHelper helper = new ExcelHelper(xlsPath);
            DataTable dt = helper.ExcelToDataTable();
            return dt;
        }
        /// <summary>
        /// ���ӱ���뵽DataTable
        /// </summary>
        /// <param name="xlsPath">excel·��</param>
        /// <param name="sheetName">�����ƣ�δ�ҵ���Ĭ��ȡ��һ����</param>
        /// <param name="isFirstRowColumn">��һ���Ƿ���DataTable������</param>
        /// <returns>���DataTable</returns>
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
        /// ���ӱ���뵽DataTable��Ĭ�ϵ�һ����
        /// </summary>
        /// <param name="xlsPath">excel·��</param>
        /// <param name="isFirstRowColumn">��һ���Ƿ���DataTable������</param>
        /// <returns>���DataTable</returns>
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

        private string fileName = null; //�ļ���
        private ExcelHelper(string fileName)
        {
            this.fileName = fileName;
        }
        /// <summary>
        /// ��DataSet���ݵ�����excel��
        /// </summary>
        /// <param name="data">Ҫ���������</param>
        /// <param name="isColumnWritten">DataTable�������Ƿ�Ҫ����</param>
        /// <returns>������������(����������һ��)</returns>
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

                    if (isColumnWritten) //д��DataTable������
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
                    workbook.Write(fs); //д�뵽excel
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
        /// ��DataTable���ݵ�����excel��
        /// </summary>
        /// <param name="data">Ҫ���������</param>
        /// <param name="isColumnWritten">DataTable�������Ƿ�Ҫ����</param>
        /// <returns>������������(����������һ��)</returns>
        public int DataTableToExcel(DataTable data, bool isColumnWritten)
        {
            DataSet ds = new DataSet();
            ds.Tables.Add(data);
            return DataSetToExcel(ds, isColumnWritten);
        }

        private IWorkbook CreateWorkBook(string fileName)
        {
            if (fileName.IndexOf(".xlsx") > 0) return new XSSFWorkbook();// 2007�汾
            else if (fileName.IndexOf(".xls") > 0) return new HSSFWorkbook(); // 2003�汾
            return null;
        }

        #region Excel2DataTable
        /// <summary>
        /// ��excel�еĵ�һ�ű����ݵ��뵽DataTable��
        /// </summary>
        /// <returns>���ص�DataTable</returns>
        public DataTable ExcelToDataTable()
        {
            return ExcelToDataTable(null, false);
        }
        /// <summary>
        /// ��excel�еĵ�һ�ű����ݵ��뵽DataTable��
        /// </summary>
        /// <param name="isFirstRowColumn">��һ���Ƿ���DataTable������</param>
        /// <returns>���ص�DataTable</returns>
        public DataTable ExcelToDataTable(bool isFirstRowColumn)
        {
            return ExcelToDataTable(null, isFirstRowColumn);
        }
        /// <summary>
        /// ��excel�е����ݵ��뵽DataSet��
        /// </summary>
        /// <param name="isFirstRowColumn">��һ���Ƿ���DataTable������</param>
        /// <returns>���ص�DataSet</returns>
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

                        dt.TableName = sheet.SheetName;//���ñ���

                        int startRow = 0;//������ʼ��

                        int maxColumnNum = GetMaxColumnNum(startRow, sheet);//�������

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
                                    cellValue = string.Format("[{0}]�ظ���{1}", GetColumnLetter(i + 1), cellValue);
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

                        //���һ�еı��
                        for (int i = startRow; i <= sheet.LastRowNum; i++)
                        {
                            IRow row = sheet.GetRow(i);

                            DataRow dataRow = dt.NewRow();
                            dt.Rows.Add(dataRow);

                            if (row == null) continue;//û�����ݵ���Ĭ����null��

                            for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                            {
                                ICell cell = row.GetCell(j);
                                if (cell == null) continue; //ͬ��û�����ݵĵ�Ԫ��Ĭ����null
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
        /// ��excel�е����ݵ��뵽DataTable��
        /// </summary>
        /// <param name="sheetName">excel������sheet������</param>
        /// <param name="isFirstRowColumn">��һ���Ƿ���DataTable������</param>
        /// <returns>���ص�DataTable</returns>
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


                    dt.TableName = sheet.SheetName;//���ñ���

                    int startRow = 0;//������ʼ��

                    int maxColumnNum = GetMaxColumnNum(startRow, sheet);//�������

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
                                cellValue = string.Format("[{0}]�ظ���{1}", GetColumnLetter(i + 1), cellValue);
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

                    //���һ�еı��
                    for (int i = startRow; i <= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);

                        DataRow dataRow = dt.NewRow();
                        dt.Rows.Add(dataRow);

                        if (row == null) continue;//û�����ݵ���Ĭ����null��

                        for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                        {
                            ICell cell = row.GetCell(j);
                            if (cell == null) continue; //ͬ��û�����ݵĵ�Ԫ��Ĭ����null
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
            // 2007�汾
            if (fileName.IndexOf(".xlsx") > 0 || fileName.IndexOf(".xlsm") > 0) return new XSSFWorkbook(fs);
            // 2003�汾
            else if (fileName.IndexOf(".xls") > 0) return new HSSFWorkbook(fs);
            return null;
        }
        private ISheet GetSheet(IWorkbook workbook, string sheetName)
        {
            if (sheetName != null)
            {
                ISheet sheet = workbook.GetSheet(sheetName);
                if (sheet != null) return sheet;
                //���û���ҵ�ָ����sheetName��Ӧ��sheet�����Ի�ȡ��һ��sheet
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
