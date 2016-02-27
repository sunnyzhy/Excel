using System.Data;

namespace G.Excel.Common
{
    public interface IExcel
    {
        DataTable GetSourceFromExcel(string filePath, int workSheetIndex);

        void GenerateExcel(string filePath, string workSheetName, DataTable table);

        void ClearWorkSheets();
    }
}
