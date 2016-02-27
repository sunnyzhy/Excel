using System;
using NPOI.SS.UserModel;
using System.IO;
using System.Data;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.Util;

namespace G.Excel.Common
{
    public class NpoiHelper:IExcel 
    {
        private volatile static NpoiHelper _instance = null;
        private static readonly object lockFlag = new object();
        private IWorkbook workbook = null;

        private NpoiHelper() { }

        public static NpoiHelper CreateInstance()
        {
            if (_instance == null)
            {
                lock (lockFlag)
                {
                    if (_instance == null)
                        _instance = new NpoiHelper();
                }
            }
            return _instance;
        }

        /// <summary>
        /// 生成Excel
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="workSheetName"></param>
        /// <param name="table"></param>
        public void GenerateExcel(string filePath, string workSheetName, DataTable table)
        {
            string fileName = string.Format("{0}\\Target\\报考.xls", filePath);
            if (workbook == null)
            {
                workbook = new HSSFWorkbook();
            }

            try
            {
                ISheet worksheet = workbook.CreateSheet(workSheetName);
                IRow wsrow = null;
                ICell wscell = null;

                ICellStyle cellstyle = GetCellStyle();
                wsrow = worksheet.CreateRow(0);
                wsrow.Height = 50 * 10;//设置行高
                wscell = wsrow.CreateCell(0);
                wscell.SetCellValue(workSheetName);//设置标题
                wscell.CellStyle = cellstyle;//设置标题样式
                IFont font = workbook.CreateFont();
                font.FontHeightInPoints = 14;//字体大小
                wscell.CellStyle.SetFont(font);//设置标题字体
                worksheet.SetColumnWidth(1, 20 * 256);//设置列宽

                cellstyle = GetCellStyle();
                //填充单元格
                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    wsrow = worksheet.CreateRow(row);
                    for (int col = 0; col < table.Columns.Count; col++)
                    {
                        wscell = wsrow.CreateCell(col);
                        wscell.SetCellValue(table.Rows[row-1][col].ToString());
                        wscell.CellStyle = cellstyle;
                    }
                }

                worksheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, table.Columns.Count - 1));//合并标题
                worksheet.AddMergedRegion(new CellRangeAddress(1, 3, 0, 0));//合并序号
                worksheet.AddMergedRegion(new CellRangeAddress(1, 3, 1, 1));//合并学号
                worksheet.AddMergedRegion(new CellRangeAddress(1, 3, table.Columns.Count - 1, table.Columns.Count - 1));//合并学生的报考总计
                worksheet.AddMergedRegion(new CellRangeAddress(table.Rows.Count, table.Rows.Count, 0, 2)); //合并专业报考总计

                using (FileStream file = new FileStream(fileName, FileMode.OpenOrCreate))
                {
                    workbook.Write(file);　　//创建Excel文件
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 设置单元格样式
        /// </summary>
        /// <returns></returns>
        private ICellStyle GetCellStyle()
        {
            ICellStyle cellstyle = workbook.CreateCellStyle();
            cellstyle.Alignment = HorizontalAlignment.Center;
            cellstyle.VerticalAlignment = VerticalAlignment.Center;
            cellstyle.WrapText = true;
            cellstyle.BorderTop = BorderStyle.Thin;
            cellstyle.TopBorderColor = HSSFColor.Black.Index;
            cellstyle.BorderRight = BorderStyle.Thin;
            cellstyle.RightBorderColor = HSSFColor.Black.Index;
            cellstyle.BorderBottom = BorderStyle.Thin;
            cellstyle.BottomBorderColor = HSSFColor.Black.Index;
            cellstyle.BorderLeft = BorderStyle.Thin;
            cellstyle.LeftBorderColor = HSSFColor.Black.Index;
            return cellstyle;
        }

        /// <summary>
        /// 获取源数据
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="workSheetIndex"></param>
        /// <returns></returns>
        public DataTable GetSourceFromExcel(string filePath, int workSheetIndex)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                return null;
            }

            DataTable table = new DataTable();
            IWorkbook workbook = null;

            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read,FileShare.ReadWrite))
                {
                    workbook = new HSSFWorkbook(fs);
                }
                ISheet worksheet = workbook.GetSheetAt(workSheetIndex-1);
                int rowStart = worksheet.FirstRowNum;
                int rowEnd = worksheet.LastRowNum;
                IRow firstRow = worksheet.GetRow(rowStart);
                int colStart = firstRow.FirstCellNum;
                int colEnd = firstRow.LastCellNum;

                if (worksheet != null)
                {
                    for (int col = colStart; col < colEnd; col++)
                    {
                        table.Columns.Add(new DataColumn(firstRow.GetCell(col).StringCellValue, typeof(string)));
                    }

                    DataRow newRow = null;
                    for (int row = rowStart + 1; row <= rowEnd; row++)
                    {
                        //try
                        //{
                            newRow = table.NewRow();
                            for (int col = colStart; col < colEnd; col++)
                            {
                                newRow[col] = worksheet.GetRow(row).GetCell(col).StringCellValue;
                            }
                            table.Rows.Add(newRow);
                        //}
                        //catch (Exception ee)
                        //{
                        //    int i = 1;
                        //    Console.WriteLine(ee.Message);
                        //}
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return table;
        }

        /// <summary>
        /// 清除缓存的worksheet
        /// </summary>
        public void ClearWorkSheets()
        {
            if (workbook != null)
            {
                int count = workbook.NumberOfSheets;
                for (int i = count-1; i >=0 ; i--)
                {
                    workbook.RemoveSheetAt(i);
                }
            }
        }

    }

}
