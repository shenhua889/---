using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Data;
using System.Windows.Forms;
namespace 创意研发室_待出库统计
{
    /// <summary>
    /// 用NPOI读取EXCEL
    /// </summary>
    public class NPExcel
    {
        private IWorkbook _WorkBook;
        private string _FilePath;
        public List<string> SheetNames { get; set; }
        public NPExcel()
        {
            SheetNames = new List<string>();
        }
        public List<string> LoadFile(string FIlePath)
        {
            _FilePath = FIlePath;
            using (var fs = new FileStream(FIlePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                    _WorkBook = WorkbookFactory.Create(fs);
            }
            return GetSheetNames();
        }
        private List<string> GetSheetNames()
        {
            SheetNames.Clear();
            var count = _WorkBook.NumberOfSheets;
            for(int i=0;i<count;i++)
            {
                SheetNames.Add(_WorkBook.GetSheetName(i));
            }
            return SheetNames;
        }
        /// <summary>
        /// 默认为第一个SheetName
        /// </summary>
        /// <returns></returns>
        public DataTable ExcelToDataTable()
        {
            //GetSheetNames();
            if (SheetNames.Count != 0)
                return ExcelToDataTable(SheetNames[0]);
            else
                return null;
        }
        public DataTable ExcelToDataTable(string SheetName)
        {
            try
            {
                ISheet sheet = _WorkBook.GetSheet(SheetName);
                DataTable dt = new DataTable();
                IRow FirstRow = sheet.GetRow(0);
                //添加列
                int CellCount = FirstRow.LastCellNum;
                for (int i = FirstRow.FirstCellNum; i < CellCount; i++)
                {
                    string CellName = FirstRow.GetCell(i).StringCellValue;
                    DataColumn dc = new DataColumn(CellName);
                    dt.Columns.Add(dc);
                }
                //添加内容
                int StartRowNum = sheet.FirstRowNum + 1;
                int RowCount = sheet.LastRowNum;
                for (int i = StartRowNum; i <= RowCount; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue;
                    DataRow dr = dt.NewRow();
                    for (int j = FirstRow.FirstCellNum; j < CellCount; j++)
                    {
                        if (row.GetCell(j) != null)
                        {
                            dr[j] = row.GetCell(j, MissingCellPolicy.RETURN_NULL_AND_BLANK).ToString();
                        }
                    }
                    dt.Rows.Add(dr);
                }
                return dt;
            }
            catch (Exception e)
            {
                MessageBox.Show("错误表为:" + _FilePath + "\r\n" + "报错信息为:" + e.ToString());
                return null;
            }
        }
        /// <summary>
        /// DataTable导出到EXCEL
        /// </summary>
        /// <param name="Table">表格</param>
        /// <param name="SaveFile">保存文件地址</param>
        /// <returns></returns>
        public string DataTableToExcel(DataTable Table,string SaveFile)
        {
            //IWorkbook WorkBook = new XSSFWorkbook();
            IWorkbook WorkBook = new HSSFWorkbook();
            ISheet Sheet = WorkBook.CreateSheet();
            int RowsCount = Table.Rows.Count;
            int ColsCount = Table.Columns.Count;
            IRow Row = Sheet.CreateRow(0);
            ICell Cell = null;
            for(int i=0;i<ColsCount;i++)
            {
                Cell = Row.CreateCell(i);
                Cell.SetCellValue(Table.Columns[i].ColumnName);
                Sheet.AutoSizeColumn(i);
            }
            for(int i=0;i<RowsCount;i++)
            {
                Row = Sheet.CreateRow(i + 1);
                for(int j=0;j<ColsCount;j++)
                {
                    Cell = Row.CreateCell(j);
                    Cell.SetCellValue(Table.Rows[i][j].ToString());
                }
            }
            //表格宽度
            for(int i=0;i<ColsCount;i++)
            {
                int ColLenght = 0;
                for(int j=0;j<=Sheet.LastRowNum;j++)
                {
                    int CellLenght =Encoding.Default.GetBytes(Sheet.GetRow(j).GetCell(i).StringCellValue).Length;
                    if (CellLenght > ColLenght)
                        ColLenght = CellLenght;
                }
                Sheet.SetColumnWidth(i, ColLenght * 256);
            }
            using (FileStream fs = File.OpenWrite(SaveFile))
            {
                WorkBook.Write(fs);
                return "保存成功";
                
            }
        }
    }
}
