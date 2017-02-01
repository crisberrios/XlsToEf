using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace XlsToEf.Import
{
    public interface IExcelIoWrapper
    {
        Task<IList<string>> GetSheets(string filePath);
        Task<List<Dictionary<string, string>>> GetRows(string filePath, string sheetName);
        Task<IList<string>> GetImportColumnData(XlsxColumnMatcherQuery matcherQuery);
    }

    public class ExcelIoWrapper : IExcelIoWrapper
    {
        public async Task<IList<string>> GetSheets(string filePath)
        {
            var sheetNames = await Task.Run(() =>
            {
                using (var excel = new ExcelPackage(new FileInfo(filePath)))
                {
                    return excel.Workbook.Worksheets.Select(x => x.Name).ToList();
                }
            });

            return sheetNames;
        }

        private async Task<IList<string>> GetColumns(string filePath, string sheetName)
        {
            var colNames = await Task.Run(() =>
            {
                using (var excel = new ExcelPackage(new FileInfo(filePath)))
                {
                    var sheet = excel.Workbook.Worksheets.First(x => x.Name == sheetName);
                    var headerCells =
                        sheet.Cells[
                            sheet.Dimension.Start.Row, sheet.Dimension.Start.Column, 1, sheet.Dimension.End.Column];
                    return headerCells.Select(x => x.Text).ToList();
                }
            });

            return colNames;
        }

        public async Task<IList<string>> GetImportColumnData(XlsxColumnMatcherQuery matcherQuery)
        {
            return await GetColumns(matcherQuery.FilePath, matcherQuery.Sheet);
        }

        public async Task<List<Dictionary<string, string>>> GetRows(string filePath, string sheetName)
        {
            var worksheetRows = await Task.Run(() =>
            {
                using (var excel = new ExcelPackage(new FileInfo(filePath)))
                {
                    var sheet = excel.Workbook.Worksheets.First(x => x.Name == sheetName);
                    var start = sheet.Dimension.Start;
                    var end = sheet.Dimension.End;
                    var rows = new List<Dictionary<string, string>>();
                    var firstDataRow = start.Row + 1;
                    var columnHeaders = sheet.Cells[start.Row, start.Column, start.Row, end.Column].Select(x => x.Text).ToList();

                    for (var rowNum = firstDataRow; rowNum <= end.Row; rowNum++)
                    {
                        var rowDict = new Dictionary<string, string>();
                        var cellValues = new List<string>();

                        for (var col = start.Column; col <= end.Column; col++)
                        { 
                            cellValues.Add(sheet.Cells[rowNum, col].Text ?? string.Empty);
                        }

                        for (var colIndex = 0; colIndex < sheet.Dimension.Columns; colIndex++)
                        {
                            var cellText = cellValues[colIndex];
                            rowDict.Add(columnHeaders[colIndex], cellText);
                        }

                        rows.Add(rowDict);
                    }

                    return rows;
                }
            });
            return worksheetRows;
        }
    }
}