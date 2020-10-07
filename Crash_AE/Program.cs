using System;
using System.Data;
using System.IO;
using System.Linq;
using Microsoft.Extensions.Configuration;
using SqlServerHelper.Core;
using System.Diagnostics;
using System.Collections.Generic;
using OfficeOpenXml;

namespace Crash_AE
{
    class Program
    {
        static void Main(string[] args)
        {
            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsetting.json", optional: true, reloadOnChange: true).Build();
            string filepath = $"{config[$"TargetFile:FilePath"]}";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            DataTabletoExcel dataTabletoExcel = new DataTabletoExcel();
            DataTable Emergency_dt = dataTabletoExcel.GetDataTableFromExcel(filepath, true);
        }
        
    }
    #region -- DataTabletoExcel --
    public class DataTabletoExcel {
        public DataTable GetDataTableFromExcel(string path, bool hasHeader = true)
        {
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }
                var ws = pck.Workbook.Worksheets.First();
                DataTable tbl = new DataTable();
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                {
                    tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                }
                var startRow = hasHeader ? 2 : 1;
                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    DataRow row = tbl.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }
                return tbl;
            }
        }
    }
    #endregion
}
