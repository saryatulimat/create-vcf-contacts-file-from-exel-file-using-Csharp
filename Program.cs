using System;
using System.IO;
using OfficeOpenXml;//you need EPPlus library(https://www.nuget.org/packages/EPPlus/)
using System.Linq;
using System.Text;

namespace contact
{
    class Program
    {
        static StreamWriter streamWriter;
        public static void writecontact(contact c)
        {
            streamWriter.WriteLine("BEGIN:VCARD");
            streamWriter.WriteLine("VERSION:2.1");
            streamWriter.WriteLine("N;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:" + c.name);
            streamWriter.WriteLine("FN;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:" + c.name);
            streamWriter.WriteLine("TEL;CELL:" + c.phone);
            streamWriter.WriteLine("END:VCARD");
        }
        public class contact{
            public string name { get; set; }    
            public string phone{get;set;}
        }
        static void Main(string[] args)
        {
            streamWriter = new StreamWriter("StoreVcfPath.vcf");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo("xlsxFilePath.xlsx")))
            {
                var myWorksheet = xlPackage.Workbook.Worksheets.FirstOrDefault(); //select the first sheet ,(in  my file i store the contacts info in the first sheet)
                var totalRows = myWorksheet.Dimension.End.Row;//get row count
                var totalColumns = myWorksheet.Dimension.End.Column; //get Columns count
            
                for (int rowNum = 1; rowNum <= totalRows; rowNum++) //start from first row to the last row
                {
                   var row = myWorksheet.Cells[rowNum,1,rowNum,totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());//read row
                   var cells= row.ToList();
                   if(!cells[4].Equals(string.Empty)){
                       writecontact(new contact{name=cells[0],phone=cells[4]});//here index 0 refers to the col that store name and index 4 Refers to the col that store phone number
                   }
                }
            }
            streamWriter.Flush();
            streamWriter.Close();
        }
    }
}
