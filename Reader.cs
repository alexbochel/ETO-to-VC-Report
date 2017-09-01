using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ConsoleApplication2
{
    /// <summary>
    /// This class is responsible for reading the data from the large excel file that holds the SAP data on all sales 
    /// numbers entered. 
    /// 
    /// @author: Alexander James Bochel
    /// @version: 8/30/2017
    /// 
    /// </summary>
    public class Reader
    {
        public List<Sales> saleList;
        public List<StringWithCount> descStringList;
        public Printer printer;
        
        public int sheet = 1;
        public _Application excel;
        public Workbooks wbs;
        public _Workbook wb;
        public _Worksheet ws;

        public string path = @"c:\users\abochel\documents\visual studio 2012\Projects\ConsoleApplication2\ConsoleApplication2\export.XLSX";

        public int location;

        /// <summary>
        /// This constructor opens the excel file, creates a new list and printer. 
        /// </summary>
        public Reader()
        {
            // Open the Excel Sheet here. 
            excel = new Application();
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];

            descStringList = new List<StringWithCount>();

            printer = new Printer(excel, wbs, wb);
            read();
            printer.printDescriptions(descStringList);
            wb.Worksheets[2].Columns.AutoFit();
            garbageCleanup();
        }

        /// <summary>
        /// This method reads through the large excel sheet and new sales to add to the list. If the sale
        /// alreader exists it simply increases the total count of the sale. 
        /// </summary>
        public void read()
        {
            int i = 2;
            int j = 1;

            int docTypeLine = 1;
            int matLine = 6;
            int saleDocLine = 3;
            int itemLine = 5;
            int descLine = 7;

            
            while (ws.Cells[i, j].Value2 != null)
            {
                // Do not forget to check for nulls or at list beginning. 
                while ((readCell(i, saleDocLine) == readCell(i - 1, saleDocLine)) 
                    && (readCell(i, itemLine) == readCell(i - 1, itemLine))
                    || (notETO(i, matLine)))
                {
                    // Skip and move on to the next line.
                    i++;
                }
                

                Sales sale = new Sales();

                // Add information to the list of sales read. 
                sale.docType = readCell(i, docTypeLine);
                sale.salesDocument = readCell(i, saleDocLine);
                sale.item = readCell(i, itemLine);
                sale.material = readCell(i, matLine);
                sale.description = readCell(i, descLine);

                if (checkExists(sale.description))
                {
                    descStringList[location].count++;
                }
                else
                {
                    // Add a new string with count to the list. 
                    StringWithCount strWC = new StringWithCount();
                    strWC.count = 1;
                    strWC.str = sale.description;
                    descStringList.Add(strWC);
                }

                printer.printDumpRow(sale, i, j);

                i++;
                j = 1;
            }
        }

        /// <summary>
        /// Checks to see if that string already exists.
        /// </summary>
        /// <param name="compStr"> The string being compared to. </param>
        /// <returns> Whether or not the string already exists. </returns> 
        public bool checkExists(string compStr)
        {
            for (int i = 0; i < descStringList.Count; i++ )
            {
                if (descStringList[i].str == compStr)
                {
                    location = i;
                    return true;
                }
            }
            return false;
        }

        public bool notETO(int i, int j)
        {
            if (readCell(i, j).Length > 2)
            {
                return readCell(i, j).Substring(0, 3) != "ETO";
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// This method reads in a cell from excel. 
        /// </summary>
        /// <param name="i"> The "y" coordinate. </param>
        /// <param name="j"> The "x" coordinate. </param>
        /// <returns> The value in the cell as a string. </returns>
        private string readCell(int i, int j)
        {
            if (ws.Cells[i, j].Value2 != null)
            {
                string cell = ws.Cells[i, j].Value2.ToString();

                return cell;
            }
            else
            {
                return "";
            }
        }

        /// <summary>
        /// Releases excel from memory. 
        /// </summary>
        public void garbageCleanup()
        {
            excel.Quit();

            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(excel);
        }
    }
}
