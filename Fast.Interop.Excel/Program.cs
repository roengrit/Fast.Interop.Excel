using Fast.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelX = Microsoft.Office.Interop.Excel;

namespace Fast.Interop.Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            var SampleData = new SampleData().GenData(); //สร้างข้อมูลทดสอบ
            object[,] RawData = new object[SampleData.Rows.Count, SampleData.Columns.Count];

            //เก็บข้อมูลใน Array
            for (int i = 0; i <= SampleData.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= SampleData.Columns.Count - 1; j++)
                {
                    RawData[i, j] = SampleData.Rows[i].ItemArray[j].ToString();
                }
            }

            /// ใส่ข้อมูล และ เปิด Excel 
            ExcelX.Application App;
            ExcelX.Workbook WorkBook;
            ExcelX.Worksheet WorkSheet;
            object MisValue = System.Reflection.Missing.Value;
            App = new ExcelX.Application();
            WorkBook = App.Workbooks.Add(MisValue);
            WorkSheet = (ExcelX.Worksheet)WorkBook.Worksheets.get_Item(1);      
                   
            //จุดสำคัญ
            var StartCell = WorkSheet.Cells[1, 1];
            var EndCell = WorkSheet.Cells[SampleData.Rows.Count, SampleData.Columns.Count];
            var WriteRange = WorkSheet.Range[StartCell, EndCell];
            WriteRange.Value2 = RawData; 

            App.Visible = true;
        
        }
    }
}
