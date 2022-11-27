using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;
using System.Collections;
using System.Reflection;
using System.Linq.Expressions;
using System.Runtime.ExceptionServices;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace ExcelAddIn1
{
    public partial class hello
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Give me a few seconds");
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            //Microsoft.Office.Interop.Excel.Application excelT = new Microsoft.Office.Interop.Excel.Application();
            object misValue = System.Reflection.Missing.Value;
            string filePathV = "C:\\Users\\veuseb\\Desktop\\valu8test1.xlsx";
            Excel.Workbook wbV;
            Excel.Worksheet wsV;
            wbV = excel.Workbooks.Open(filePathV);
            wsV = wbV.Worksheets[1];

            string filePathT = "C:\\Users\\veuseb\\Desktop\\SagaTarget2022.xlsm";
            Excel.Workbook wbT;
            Excel.Worksheet wsT;
            wbT = excel.Workbooks.Open(filePathT);
            wsT = wbT.Worksheets[2];


            Dictionary<string, string[]> dict = ReadComp(excel, wbV, wsV);
            Stack<string> columnValues = GetDataType(excel);
            //excel.Visible = true;

            Colourize(excel, wsT);
            WriteTargetComps(excel, dict, columnValues, wsT);
            wsT.Unprotect();
            wbT.Unprotect();
            //wbT.Save();
           
            //wbV.Close(false, misValue, misValue);
            MessageBox.Show("You are all set");

        }

        private void Colourize(Microsoft.Office.Interop.Excel.Application excel, Worksheet ws)
        {
            int[] cols= new int[] {2, 6, 8, 9};
            var columnHeadingsRange = ws.Range[ws.Cells[1, 1], ws.Cells[11, 36]];
            columnHeadingsRange.Interior.Color = XlRgbColor.rgbBlack;
            foreach(var col in cols) { Range fillColumns = ws.Cells[10, col];
                fillColumns.Interior.Color = XlRgbColor.rgbYellow;
                fillColumns.Font.Color = XlRgbColor.rgbBlack;
                fillColumns.Value = "Fill";
            }
            
            excel.Visible = true;
            MessageBox.Show("A few more seconds");
        }
        private int GetIndex(Microsoft.Office.Interop.Excel.Application excel, string name, Worksheet ws)
        {
            //string filePath = "C:\\Users\\veuseb\\Desktop\\valu8test1.xlsx";
            //Excel.Workbook wb;
            //Excel.Worksheet ws;
            //wb = excel.Workbooks.Open(filePath);
            //ws = wb.Worksheets[1];
            object misValue = System.Reflection.Missing.Value;
            Range UsedRange = ws.UsedRange;
            int colCount = UsedRange.Columns.Count;

            int varIndex = 1;
            try
            {
                for (int i = 1; i < colCount; i++)
                {
                    Range cellHead = ws.Cells[14, i];
                    string ofCellHead = cellHead.Value;
                    if (ofCellHead == name)
                    {
                        varIndex = i;
                        Range temp = ws.Cells[1, 2];
                        temp.Value = varIndex;
                        return varIndex;
                    }
                }
                return 0;
            }
            catch (ArgumentException e)
            {
                return -1;
            }
        }
        //Reads data from Valu8 file
        private Dictionary<string, string[]> ReadComp(Microsoft.Office.Interop.Excel.Application excel, Workbook wb, Worksheet ws)
        {
            //string filePath = "C:\\Users\\veuseb\\Desktop\\valu8test1.xlsx";
            Dictionary<string, string[]> hash = new Dictionary<string, string[]>();

            //Excel.Workbook wb;
            //Excel.Worksheet ws;
            object misValue = System.Reflection.Missing.Value;
            //wb = excel.Workbooks.Open(filePath);
            //ws = wb.Worksheets[1];
            Range UsedRange = ws.UsedRange;
            int lastUsedRow = UsedRange.Row + UsedRange.Rows.Count - 1;
            int orgNr = 1;
            int companyNameIndex = GetIndex(excel, "Name", ws);
            int orgNrIndex = GetIndex(excel, "Registration no.", ws);
            int industryIndex = GetIndex(excel, "SIC/NACE industry", ws);
            int companyTypeIndex = GetIndex(excel, "Company type", ws);

            for (int i = 15; i < lastUsedRow + 1; i += 21)
            {
                string[] temp;
                Range cellValue = ws.Cells[i, companyNameIndex]; //[row, column] --> company name
                Range cellKey = ws.Cells[i, orgNrIndex]; //reg.nr.
                Range cellValue2 = ws.Cells[i, industryIndex]; //industry
                Range cellValue3 = ws.Cells[i, companyTypeIndex];//company type

                string ofCellKey = cellKey.Value;//reg.nr.
                string ofCellValue = cellValue.Value;//company name
                string ofcellValue2 = cellValue2.Value;//industry
                string ofcellValue3 = cellValue3.Value;//company type
                if (hash.ContainsKey(ofCellKey) && ofCellKey != "") { continue; }
                else
                {
                    if (ofCellKey == "")
                    {
                        ofCellKey = orgNr.ToString() + "00000000";
                        orgNr++;
                    }
                    temp = new string[3] { ofCellValue, ofcellValue2, ofcellValue3 };
                }
                var list = new List<string>();
                for (int j = i; j < 22 + i - 1; j++)
                {
                    Range cellDataValueForComp = ws.Cells[j, 17];
                    list.Add(cellDataValueForComp.Value.ToString());
                }
                var array = list.ToArray();
                string[] combined = temp.Concat(array).ToArray();
                hash.Add(ofCellKey, combined);
            }
            
            return hash;
        }

        //generates all columns
        private Stack<string> GetDataType(Microsoft.Office.Interop.Excel.Application excel)
        {
            string filePath = "C:\\Users\\veuseb\\Desktop\\valu8test1.xlsx";
            Stack<string> typeStack = new Stack<string>();
            Excel.Workbook wb;
            Excel.Worksheet ws;
            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[1];
            object misValue = System.Reflection.Missing.Value;

            Range UsedRange = ws.UsedRange;
            int lastUsedRow = UsedRange.Row + UsedRange.Rows.Count - 1;

            for (int i = 15; i < lastUsedRow + 1; i++)
            {
                string newColumn;
                Range cellValue = ws.Cells[i, 15]; //[row, column]
                Range cellTime = ws.Cells[i, 16]; //[row, column]
                var ofCellTime = cellTime.Value.ToString();

                if (ofCellTime == "0")
                {
                    newColumn = cellValue.Value;
                }
                else
                {
                    newColumn = cellValue.Value + " " + cellTime.Value.ToString().Substring(0, 10);
                }
                if (!(typeStack.Contains(newColumn)))
                {
                    typeStack.Push(newColumn);
                }
                else
                {
                    break;
                }
            }
            
            wb.Close(false, misValue, misValue);
            return typeStack;
        }
        private void WriteTargetComps(Microsoft.Office.Interop.Excel.Application excel, Dictionary<string, string[]> companies, Stack<string> columnValues, Worksheet ws)
        {
            //string filePath = "C:\\Users\\veuseb\\Desktop\\SagaTarget2022.xlsm";
            //Excel.Workbook wb;
            //Excel.Worksheet ws;
            //wb = excel.Workbooks.Open(filePath);
            //ws = wb.Worksheets[2];
            int j = 0;

            for (int i = 16; i < columnValues.Count() + 16; i++)
            {
                Range cell = ws.Cells[11, i];
                cell.Value = columnValues.ElementAt(j);
                j++;
            }
            for (int i = 0; i < companies.Count(); i++)
            {
                var item = companies.ElementAt(i);
                int index = i + 12;
                Range cellCompName = ws.Cells[index, 3];
                Range cellCompNr = ws.Cells[index, 4];
                Range cellCompIndustry = ws.Cells[index, 12];
                Range cellCompType = ws.Cells[index, 14];//company type
                Range employees = ws.Cells[index, 16]; //Employees
                Range equityRatio = ws.Cells[index, 17];
                Range shareholdersEquirt = ws.Cells[index, 18];
                Range totalAssets = ws.Cells[index, 19];
                Range netProfitMargin = ws.Cells[index, 20];
                Range ebitMargin = ws.Cells[index, 21];
                Range ebitDaMargin31122018 = ws.Cells[index, 22];
                Range ebitDaMargin31122019 = ws.Cells[index, 23];
                Range ebitDaMargin31122020 = ws.Cells[index, 24];
                Range ebitDaMargin31122021 = ws.Cells[index, 25];
                Range ebitDaMargin31122022 = ws.Cells[index, 26];
                Range ebitDaMargin = ws.Cells[index, 27];
                Range netSalesGrowth = ws.Cells[index, 28];
                Range ebit = ws.Cells[index, 29];
                Range ebitDa31122018 = ws.Cells[index, 30];
                Range ebitDa31122019 = ws.Cells[index, 31];
                Range ebitDa31122020 = ws.Cells[index, 32];
                Range ebitDa31122021 = ws.Cells[index, 33];
                Range ebitDa31122022 = ws.Cells[index, 34];
                Range ebitDa = ws.Cells[index, 35];
                Range netSales = ws.Cells[index, 36];

                //Employees Equity ratio	Shareholders equity	Total assets	Net profit margin	EBIT margin	EBITDA margin 31.12.2018	
                //EBITDA margin 31.12.2019	EBITDA margin 31.12.2020	EBITDA margin 31.12.2021	EBITDA margin 31.12.2022	
                //EBITDA margin	Net sales growth	EBIT	EBITDA 31.12.2018	EBITDA 31.12.2019	EBITDA 31.12.2020	
                //EBITDA 31.12.2021	EBITDA 31.12.2022	EBITDA	Net sales

                cellCompName.Value = item.Key;
                cellCompNr.Value = item.Value[0];
                cellCompIndustry.Value = item.Value[1];
                cellCompType.Value = item.Value[2];
                employees.Value = item.Value[23];
                equityRatio.Value = item.Value[22];
                shareholdersEquirt.Value = item.Value[21];
                totalAssets.Value = item.Value[20];
                netProfitMargin.Value = item.Value[19];
                ebitMargin.Value = item.Value[18];
                ebitDaMargin31122018.Value = item.Value[17];
                ebitDaMargin31122019.Value = item.Value[16];
                ebitDaMargin31122020.Value = item.Value[15];
                ebitDaMargin31122021.Value = item.Value[14];
                ebitDaMargin31122022.Value = item.Value[13];
                ebitDaMargin.Value = item.Value[12];
                netSalesGrowth.Value = item.Value[11];
                ebit.Value = item.Value[10];
                ebitDa31122018.Value = item.Value[9];
                ebitDa31122019.Value = item.Value[8];
                ebitDa31122020.Value = item.Value[7];
                ebitDa31122021.Value = item.Value[6];
                ebitDa31122022.Value = item.Value[5];
                ebitDa.Value = item.Value[4];
                netSales.Value = item.Value[3];
            }
            ws.Columns.AutoFit();
        }

    }

}

//-combined    { string[24]}
//string[]
//        [0] "REACH SUBSEA ASA"  string
//        [1] "n/a"   string
//        [2] "Independent"   string
//        [3] "628,030000003052"  string
//        [4] "267,207000004992"  string
//        [5] "n/a"   string
//        [6] "n/a"   string
//        [7] "267,207000004992"  string
//        [8] "171,40599999894"   string
//        [9] "274,999999998763"  string
//        [10] "51,3170000045363"  string
//        [11] "0,23545994095876"  string
//        [12] "0,425468528579357" string
//        [13] "n/a"   string
//        [14] "n/a"   string
//        [15] "0,425468528579357" string
//        [16] "0,337189698959664" string
//        [17] "0,407407407403828" string
//        [18] "0,081711064764879" string
//        [19] "0,0686320717102542"    string
//        [20] "363,012999998116"  string
//        [21] "210,154000002117"  string
//        [22] "0,578915906601712" string
//        [23] "n/a"   string

//Net sales
//EBITDA
//EBITDA
//EBITDA
//EBITDA
//EBITDA
//EBITDA
//EBIT
//Net sales growth
//EBITDA margin
//EBITDA margin
//EBITDA margin
//EBITDA margin
//EBITDA margin
//EBITDA margin
//EBIT margin
//Net profit margin
//Total assets
//Shareholders equity
//Equity ratio
//Employees

