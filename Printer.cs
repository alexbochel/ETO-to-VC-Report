using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication2
{
    public class Printer
    {

        private _Application excel;
        private Workbooks wbs;
        private _Workbook wb;
        private _Worksheet ws;
        private const int sheetNumber = 2;

        public Printer(_Application excel, Workbooks wbs,  _Workbook wb)
        {
            this.excel = excel;
            this.wbs = wbs;
            this.wb = wb;
            this.ws = wb.Worksheets[sheetNumber];
            printHeaders();
        }

        public void printDumpRow(Sales sale, int i, int j)
        {
            printCell(i, j, sale.docType);
            j++;
            printCell(i, j, sale.salesDocument);
            j++;
            printCell(i, j, sale.item);
            j++;
            printCell(i, j, sale.material);
            j++;
            printCell(i, j, sale.description);
            j++;
            printCell(i, j, sale.deliveryDate);
        }

        public void printDescriptions(List<StringWithCount> list)
        {
            for (int i = 0; i < list.Count; i++ )
            {
                printCell(i + 2, 8, list[i].str);
                printCell(i + 2, 9, list[i].count.ToString());
            }
        }

        public void printHeaders()
        {
            printCell(1, 1, "Document Type");
            printCell(1, 2, "Sales Document");
            printCell(1, 3, "Item Number");
            printCell(1, 4, "Material");
            printCell(1, 5, "Description");
            printCell(1, 6, "Delivery Date");

            printCell(1, 8, "Description");
            printCell(1, 9, "Count");

            ws.get_Range("A1", "Z1").Font.Bold = true;
        }

        /// <summary>
        /// This method prints data in a cell. 
        /// </summary>
        /// <param name="i"> The "y" coordinate on a plane. </param>
        /// <param name="j"> The "x" coordinate on a plane. </param>
        /// <param name="value"> The data to be printed in the cell. </param>
        private void printCell(int i, int j, string value)
        {
            ws.Cells[i, j].Value2 = value;
        }
    }
}
