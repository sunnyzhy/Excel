using System;
using System.Data;
using System.IO;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace G.Excel.Common
{
    /// <summary>
    /// EPPlus操作类
    /// </summary>
    public class EPPlusHelper : IExcel
    {
        private volatile static EPPlusHelper _instance = null;
        private static readonly object lockFlag = new object();
        private ExcelPackage package = null;

        private EPPlusHelper() { }

        public static EPPlusHelper CreateInstance()
        {
            if (_instance == null)
            {
                lock (lockFlag)
                {
                    if (_instance == null)
                        _instance = new EPPlusHelper();
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
            //if (package == null)
            //{
            //    package = new ExcelPackage(new FileInfo(string.Format("{0}\\Target\\报考.xlsx", filePath)));
            //}
            ExcelPackage package = new ExcelPackage(new FileInfo(string.Format("{0}\\Target\\报考.xlsx", filePath)));

            //判断表单中是否存在名为workSheetName的表单，如果存在，就删除同名的表单
            for (int i = package.Workbook.Worksheets.Count; i >= 1; i--)
            {
                if (package.Workbook.Worksheets[i].Name.Equals(workSheetName))
                {
                    package.Workbook.Worksheets.Delete(i);
                }
            }
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(workSheetName);
            worksheet.Cells.Style.WrapText = true;

            worksheet.Cells[1, 1].Value = workSheetName;//设置标题
            worksheet.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[1, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Row(1).Height = 20;//设置行高
            worksheet.Cells[1, 1].Style.Font.Size = 14;//字体大小
            worksheet.Column(2).Width = 20;//设置列宽

            //填充单元格
            for (int row = 1; row <= table.Rows.Count; row++)
            {
                for (int col = 1; col <= table.Columns.Count; col++)
                {
                    worksheet.Cells[row + 1, col].Value = table.Rows[row - 1][col - 1].ToString();
                    worksheet.Cells[row + 1, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row + 1, col].Style.Fill.BackgroundColor.SetColor(Color.Transparent);
                    worksheet.Cells[row + 1, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[row + 1, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[row + 1, col].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                }
            }

            worksheet.Cells[1, 1, 1, table.Columns.Count].Merge = true; //合并标题
            worksheet.Cells[2, 1, 4, 1].Merge = true; //合并序号
            worksheet.Cells[2, 2, 4, 2].Merge = true; //合并学号
            worksheet.Cells[2, table.Columns.Count, 4, table.Columns.Count].Merge = true; //合并学生的报考总计
            worksheet.Cells[table.Rows.Count + 1, 1, table.Rows.Count + 1, 3].Merge = true; //合并专业报考总计

            try
            {
                package.Save();//保存文件
            }
            catch (Exception ex)
            {
                throw ex;
            }
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
            try
            {
                ExcelPackage package = new ExcelPackage(new FileInfo(filePath));

                ExcelWorksheet worksheet = package.Workbook.Worksheets[workSheetIndex];

                int columnStart = worksheet.Dimension.Start.Column; //列数据的起始索引
                int columnEnd = worksheet.Dimension.End.Column; //列数据的终止索引

                int rowStart = worksheet.Dimension.Start.Row; //行数据的起始索引
                int rowEnd = worksheet.Dimension.End.Row; //行数据的终止索引

                for (int col = columnStart; col <= columnEnd; col++)
                {
                    table.Columns.Add(new DataColumn(worksheet.Cells[rowStart, col].Value.ToString(), typeof(string)));
                }

                DataRow newRow = null;
                for (int row = rowStart + 1; row <= rowEnd; row++)
                {
                    newRow = table.NewRow();
                    for (int col = columnStart; col <= columnEnd; col++)
                    {
                        newRow[col - 1] = worksheet.Cells[row, col].Value;
                    }
                    table.Rows.Add(newRow);
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
            if (package != null)
            {
                int count = package.Workbook.Worksheets.Count;
                for (int i = count - 1; i >= 0; i--)
                {
                    package.Workbook.Worksheets.Delete(i);
                }
            }
        }
    }

}
