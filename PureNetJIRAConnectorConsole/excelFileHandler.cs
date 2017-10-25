using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System.Linq;

public class ExcelFileHandler {

    private string fileName = "report.xlsx";

    public ReadExcel()
    {
        var package = new ExcelPackage(new FileInfo(fileName));
 
        ExcelWorksheet workSheet = package.Workbook.Worksheets[1];
        
        // Read source excel file
        DataTable dt = new DataTable();
        dt=package.ToDataTable();
        List<DataRow> listOfRows = new List<DataRow>();
        listOfRows = dt.AsEnumerable().ToList();

        var listOfAccounts = listOfRows.Select(r => new AccountDataRow()
        {
            ID = listOfRows["ID"];
            accountName = listOfRows["AccountName"];
        });

        // Print value of internal object
        for (int i=0, i <= listOfAccounts.Rows.Count, i++)
        {
            Console.WriteLine(listOfAccounts[i].ID);
            Console.WriteLine(listOfAccounts[i].accountName);
        }

    }

    public ExcelPackage createExcelPackage()
        {
            var package = new ExcelPackage();
            package.Workbook.Properties.Title = "Salary Report";
            package.Workbook.Properties.Author = "Vahid N.";
            package.Workbook.Properties.Subject = "Salary Report";
            package.Workbook.Properties.Keywords = "Salary";

            var worksheet = package.Workbook.Worksheets.Add("Employee");

            //First add the headers
            worksheet.Cells[1, 1].Value = "ID";
            worksheet.Cells[1, 2].Value = "Name";
            worksheet.Cells[1, 3].Value = "Gender";
            worksheet.Cells[1, 4].Value = "Salary (in $)";

            //Add values

            var numberformat = "#,##0";
            var dataCellStyleName = "TableNumber";
            var numStyle = package.Workbook.Styles.CreateNamedStyle(dataCellStyleName);
            numStyle.Style.Numberformat.Format = numberformat;

            worksheet.Cells[2, 1].Value = 1000;
            worksheet.Cells[2, 2].Value = "Jon";
            worksheet.Cells[2, 3].Value = "M";
            worksheet.Cells[2, 4].Value = 5000;
            worksheet.Cells[2, 4].Style.Numberformat.Format = numberformat;

            worksheet.Cells[3, 1].Value = 1001;
            worksheet.Cells[3, 2].Value = "Graham";
            worksheet.Cells[3, 3].Value = "M";
            worksheet.Cells[3, 4].Value = 10000;
            worksheet.Cells[3, 4].Style.Numberformat.Format = numberformat;

            worksheet.Cells[4, 1].Value = 1002;
            worksheet.Cells[4, 2].Value = "Jenny";
            worksheet.Cells[4, 3].Value = "F";
            worksheet.Cells[4, 4].Value = 5000;
            worksheet.Cells[4, 4].Style.Numberformat.Format = numberformat;

            // Add to table / Add summary row
            var tbl = worksheet.Tables.Add(new ExcelAddressBase(fromRow: 1, fromCol: 1, toRow: 4, toColumn: 4), "Data");
            tbl.ShowHeader = true;
            tbl.TableStyle = TableStyles.Dark9;
            tbl.ShowTotal = true;
            tbl.Columns[3].DataCellStyleName = dataCellStyleName;
            tbl.Columns[3].TotalsRowFunction = RowFunctions.Sum;
            worksheet.Cells[5, 4].Style.Numberformat.Format = numberformat;

            return package;
        }

        public void SaveFile()
        {
            using (var package = createExcelPackage())
            {
                package.SaveAs(new FileInfo(@fileName));
            }
            
        }

}
