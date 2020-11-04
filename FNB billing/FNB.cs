﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelTools = Microsoft.Office.Tools.Excel;

namespace FNB_billing
{
    class FNB
    {
        internal static void QSummary() //Create a summry of the quote sheets
        {
            Excel.Application xlAp = Globals.ThisAddIn.Application;
            Excel.Workbook XlWb = xlAp.ActiveWorkbook;
            Excel.Worksheet QSSht;
            string QSummShtName = "QuoteSummary";

            try
            {
                Globals.ThisAddIn.LogTrackInfo("FNBQSummary");
                xlAp.ScreenUpdating = false;
                if (ExistSheet(QSummShtName))
                {
                    xlAp.DisplayAlerts = false;
                    XlWb.Worksheets[QSummShtName].delete();
                    xlAp.DisplayAlerts = true;
                }
                QSSht = XlWb.Worksheets.Add();
                QSSht.Name = QSummShtName;
                int Col = 0;
                int Row = 4;
                QSSht.Cells[1, 1].value = "Do not change. Change data in the quote sheets";
                QSSht.Cells[2, 1].value = "Summary of all quotes, purchase orders and invoice details on the quote sheets";
                QSSht.Cells[2, 1].Font.Size = 14;
                QSSht.Cells[Row, Col += 1].value = "Country";
                QSSht.Cells[Row, Col += 1].value = "Province";
                QSSht.Cells[Row, Col += 1].value = "City";
                QSSht.Cells[Row, Col += 1].value = "Branch";
                QSSht.Cells[Row, Col += 1].value = "Quote Total";
                QSSht.Cells[Row, Col += 1].value = "Po No";
                QSSht.Cells[Row, Col += 1].value = "Po Date";
                QSSht.Cells[Row, Col += 1].value = "Po Status";
                QSSht.Cells[Row, Col += 1].value = "Po Amount";
                QSSht.Cells[Row, Col += 1].value = "Q Total";
                QSSht.Cells[Row, Col += 1].value = "Inv No";
                QSSht.Cells[Row, Col += 1].value = "Inv Date";
                QSSht.Cells[Row, Col += 1].value = "Inv Amount";

                foreach (Excel.Worksheet Sht in XlWb.Worksheets)
                {
                    if (IsQuoteSht(Sht))
                    {
                        Col = 0;
                        Row += 1;
                        QSSht.Cells[Row, Col += 1].value = "=" + Sht.Name + "!QCountry";
                        QSSht.Cells[Row, Col += 1].value = "=" + Sht.Name + "!QProvince";
                        QSSht.Cells[Row, Col += 1].value = "=" + Sht.Name + "!QCity";
                        QSSht.Cells[Row, Col += 1].value = "=" + Sht.Name + "!QBranch";
                        QSSht.Cells[Row, Col += 1].value = "=" + Sht.Name + "!QTotal";
                        QSSht.Cells[Row, Col += 1].value = "=" + Sht.Name + "!PoNo";
                        QSSht.Cells[Row, Col += 1].value = "=" + Sht.Name + "!PoDate";
                        QSSht.Cells[Row, Col += 1].value = "=" + Sht.Name + "!PoStatus";
                        QSSht.Cells[Row, Col += 1].value = "=" + Sht.Name + "!PoAmount";
                        QSSht.Cells[Row, Col += 1].value = "=" + Sht.Name + "!InvNo";
                        QSSht.Cells[Row, Col += 1].value = "=" + Sht.Name + "!InvDate";
                        QSSht.Cells[Row, Col += 1].value = "=" + Sht.Name + "!InvAmount";

                        //MessageBox.Show(Sheet.Name + " is a quote sheet");
                    }
                    else
                    {
                        //MessageBox.Show(Sheet.Name + " is not a quote sheet");
                    }
                }
                Excel.ListObject QSummList = QSSht.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, QSSht.Cells[Row, 1].CurrentRegion, false,
                    Excel.XlYesNoGuess.xlYes);
                QSummList.Name = "Tab" + QSummShtName;
                QSummList.ListColumns["Quote Total"].DataBodyRange.NumberFormat = "R# ##0.00";
                QSummList.ListColumns["Po Amount"].DataBodyRange.NumberFormat = "R# ##0.00";
                QSummList.ListColumns["Inv Amount"].DataBodyRange.NumberFormat = "R# ##0.00";
                QSummList.ListColumns["Po Date"].DataBodyRange.NumberFormat = "yyyy-mm-dd";
                QSummList.ListColumns["Inv Date"].DataBodyRange.NumberFormat = "yyyy-mm-dd";
                QSummList.DataBodyRange.Columns.AutoFit();
                QSSht.Protect(
                    DrawingObjects: true,
                    Contents: true,
                    Scenarios: true,
                    AllowFormattingCells: true,
                    AllowFormattingColumns: true,
                    AllowFormattingRows: true,
                    AllowInsertingColumns: true,
                    AllowInsertingRows: true,
                    AllowSorting: true,
                    AllowFiltering: true,
                    AllowUsingPivotTables: true);

                xlAp.ScreenUpdating = true;
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.ExMsg(ex);
            }
        }
        internal static void PoImport() //Imports PO info from export.csv query in Excel
        {
            Excel.Application xlAp = Globals.ThisAddIn.Application;
            Excel.Workbook XlWb = xlAp.ActiveWorkbook;
            string PoExportTableName = "PoExport";
            string PoExportShtName;
            int PoExportRow = 0;
            Excel.ListObject PoExportTable;
            Excel.Range PoExportBody;

            try
            {
                Globals.ThisAddIn.LogTrackInfo("FNBPoImport");
                PoExportShtName = ExistListObject(PoExportTableName);
                if (PoExportShtName == "") //Test if PoExport table exists
                {
                    MessageBox.Show("The PoExport table is required");
                    return;
                }
                xlAp.ScreenUpdating = false;
                PoExportTable = XlWb.Worksheets[PoExportShtName].ListObjects[PoExportShtName];
                PoExportBody = PoExportTable.DataBodyRange;

                foreach (Excel.Worksheet Sht in XlWb.Worksheets)
                {
                    if (IsQuoteSht(Sht))
                    {
                        if (Sht.Range["PoNo"].Text != "") //Po No is known
                        {
                            PoExportRow = (int)xlAp.WorksheetFunction.Match(Sht.Range["PoNo"].Text,
                                PoExportBody.Cells[1, 1].EntireColumn, 0);
                        }
                        else
                        {
                            var FindRange = PoExportBody.Find(Sht.Range["QBranch"]);
                            if (FindRange != null)
                            {
                                PoExportRow = FindRange.Row;
                            }
                        }
                        if (PoExportRow > 0)
                        {
                            Sht.Range["PoNo"].Value = PoExportBody.Worksheet.Cells[PoExportRow, PoExportTable.ListColumns.Item["Purchase Order Number"].Index].Value;
                            Sht.Range["PoDate"].Value = PoExportBody.Worksheet.Cells[PoExportRow, PoExportTable.ListColumns.Item["Purchase Order Date"].Index].Value;
                            Sht.Range["PoStatus"].Value = PoExportBody.Worksheet.Cells[PoExportRow, PoExportTable.ListColumns.Item["Purchase Order Status"].Index].Value;
                            Sht.Range["PoAmount"].Value = PoExportBody.Worksheet.Cells[PoExportRow, PoExportTable.ListColumns.Item["Total"].Index].Value;
                        }
                        else
                        {
                            MessageBox.Show("Cannot find Po info for " + Sht.Name);
                        }
                    }
                    xlAp.ScreenUpdating = true;
                }
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.ExMsg(ex);
            }
        }
        internal static void QHideBillRows() //Hide unsused bill rows on the active quote sheet
        {
            Excel.Application xlAp = Globals.ThisAddIn.Application;
            Excel.Workbook XlWb = xlAp.ActiveWorkbook;
            Excel.Worksheet QSSht = XlWb.ActiveSheet;
            int BillStartRow;
            int BillEndRow;

            if (IsQuoteSht(QSSht)) //Do nothing if it is not a quote sheet
            {
                BillStartRow = QSSht.Range["A:A"].Find(What: "#BillStart").Row;
                BillEndRow = QSSht.Range["A:A"].Find(What: "#BillEnd").Row;
                for (int BillRow = BillStartRow + 1; BillRow < BillEndRow; BillRow += 1)
                {
                    if (QSSht.Cells[BillRow,4].text.Length>0 && (QSSht.Cells[BillRow,6].text=="" || QSSht.Cells[BillRow,6].value==0))
                    {
                        QSSht.Cells[BillRow, 1].EntireRow.Hidden = true;
                    }
                    else
                    {
                        QSSht.Cells[BillRow, 1].EntireRow.Hidden = false;
                    }
                }
            }
        }
        internal static void QUnHideBillRows() //Unhide all bill rows
        {
            Excel.Application xlAp = Globals.ThisAddIn.Application;
            Excel.Workbook XlWb = xlAp.ActiveWorkbook;
            Excel.Worksheet QSSht = XlWb.ActiveSheet;
            int BillStartRow;
            int BillEndRow;

            if (IsQuoteSht(QSSht)) //Do nothing if it is not a quote sheet
            {
                BillStartRow = QSSht.Range["A:A"].Find(What: "#BillStart").Row;
                BillEndRow = QSSht.Range["A:A"].Find(What: "#BillEnd").Row;
                for (int BillRow = BillStartRow + 1; BillRow < BillEndRow; BillRow += 1)
                {
                    QSSht.Cells[BillRow, 1].EntireRow.Hidden = false;
                }
            }
        }
        internal static bool ExistSheet(string SheetName) // Returns true if a sheet exists in the workbook
        {
            Excel.Application xlAp = Globals.ThisAddIn.Application;
            Excel.Workbook XlWb = xlAp.ActiveWorkbook;

            foreach (Excel.Worksheet Sht in XlWb.Worksheets) // Loop through all the worksheets
            {
                if (Sht.Name == SheetName)
                {
                    return true;
                }
            }
            return false;
        }
        internal static bool IsQuoteSht(Excel.Worksheet XlSh) // Returns true if a sheet complies with all the quote requirements
        {
            string[] LineCodes = { "#BillStart", "BillEnd", "#BSTStart", "#BSTEnd" };
            string[] CellNames = { "PoDate", "PoNo", "PoStatus", "PoAmount", "QBranch","InvNo","InvDate","InvNo"}; //Check only required names
            Excel.Range LineCodeRange = XlSh.Range["A:A"];

            foreach (string LineCode in LineCodes) // Loop through line codes
            {
                var FindRange = LineCodeRange.Find(What: LineCode);
                if (FindRange == null) { return false; }
            }
            foreach (string CellName in CellNames) // Loop through cell names
            {
                try
                { 
                    var NamedRange = XlSh.Names.Item(CellName); 
                }
                catch 
                { 
                    return false;
                }
            }
            return true;
        }
        internal static string ExistListObject(string ListName) // Returns sheet name if a list object exist in the workbook
        {
            Excel.Application xlAp = Globals.ThisAddIn.Application;
            Excel.Workbook XlWb = xlAp.ActiveWorkbook;

            foreach (Excel.Worksheet Sht in XlWb.Worksheets) // Loop through all the worksheets
            {
                foreach (Excel.ListObject ListObj in Sht.ListObjects) // Loop through each table in the worksheet
                {
                    if (ListObj.Name == ListName)
                    {
                        return Sht.Name;
                    }
                }
            }
            return "";
        }
    }
}