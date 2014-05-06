using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace EventTimeNormalizer
{
    class Program
    {
        private static log4net.ILog log = log4net.LogManager.GetLogger(typeof(Program));

        private static bool fFirstRowHeader = true;
        static void Main(string[] args)
        {
            using (System.IO.Stream s = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("EventTimeNormalizer.log4net configuration.xml"))
            {
                log4net.Config.XmlConfigurator.Configure(s);
            }
            log.DebugFormat("{0:S} v{1:S} started", System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString());
            SpreadsheetDocument sd = SpreadsheetDocument.Open(@"D:\GIT\EventTimeNormalizer\EventTimeNormalizer\sample\data.xlsx", false);

            WorkbookPart wp = sd.Parts.First(item => item.OpenXmlPart is WorkbookPart).OpenXmlPart as WorkbookPart;
            Sheet sheet = wp.Workbook.Descendants<Sheet>().First();
            WorksheetPart worksheetPart1 = (WorksheetPart)(wp.GetPartById(sheet.Id));
            SheetData sheetData = worksheetPart1.Worksheet.GetFirstChild<SheetData>();

            SharedStringTable sharedStringTable = wp.SharedStringTablePart.SharedStringTable;


            int iPos = 0;
            int iValueCell = 2;
            int iDataCell = 3;

            List<RowDateValue> lRDV = new List<RowDateValue>();

            foreach (Row row in sheetData.AsEnumerable())
            {
                if (iPos == 0 && fFirstRowHeader)
                {
                    iPos++;
                    continue;
                }

                lRDV.Add(new RowDateValue(iPos++));

                Cell cell = row.ElementAt(iValueCell) as Cell;
                double dVal = double.Parse(cell.CellValue.Text);
                lRDV.Last().Value = dVal;

                cell = row.ElementAt(iDataCell) as Cell;
                if (cell.DataType == CellValues.Date)
                {
                    string s = cell.CellValue.Text;
                    DateTime dt = DateTime.Parse(s);
                    lRDV.Last().Date = dt;
                }
                else if (cell.DataType == CellValues.SharedString)
                {
                    SharedStringItem ssi = sharedStringTable.ChildElements[int.Parse(cell.CellValue.Text)] as SharedStringItem;
                    DateTime dt = DateTime.Parse(ssi.Text.Text);
                    lRDV.Last().Date = dt;
                }

                else if ((cell.StyleIndex != null) && (cell.StyleIndex == 2))
                {
                    DateTime dt = DateTime.FromOADate(double.Parse(cell.CellValue.InnerXml));
                    lRDV.Last().Date = dt;
                }
            }

            // sort them by date
            lRDV.Sort((RowDateValue r1, RowDateValue r2) =>
            {
                int i = r1.Date.CompareTo(r2.Date);
                if (i != 0)
                    return i;

                return r1.OriginalPosition.CompareTo(r2.OriginalPosition);
            });

            bool fFirst = true;
            foreach (Row row in sheetData.AsEnumerable())
            {
                if (fFirst && fFirstRowHeader)
                {
                    fFirst = false;
                    continue;
                }

                StringBuilder sb = new StringBuilder();
                StringBuilder sb2 = new StringBuilder();

                foreach (Cell cell in row)
                {
                    sb.Append(cell.CellReference.Value + "-");

                    if (cell.DataType != null)
                    {
                        if (cell.DataType == CellValues.SharedString)
                        {
                            SharedStringItem ssi = sharedStringTable.ChildElements[int.Parse(cell.CellValue.Text)] as SharedStringItem;
                            sb2.Append(ssi.Text.Text + "-");
                        }
                        else if (cell.DataType == CellValues.Date)
                        {
                            sb2.Append(cell.CellValue.Text);
                        }
                        else
                        {
                            int jj = 0;
                        }
                    }
                    else
                    {
                        if ((cell.StyleIndex != null) && (cell.StyleIndex == 2))
                        {
                            DateTime dt = DateTime.FromOADate(double.Parse(cell.CellValue.InnerXml));
                            sb2.Append(dt.ToString());
                        }
                        else
                        {
                            if (cell.CellValue != null)
                                sb2.Append(cell.CellValue.InnerXml);
                        }
                    }

                }

                log.Debug(sb.ToString());
                log.Debug(sb2.ToString());
            }


            //WorksheetPart wps = wp.WorksheetParts.First();

            //Worksheet ws = wps.Worksheet;

            //SheetData sheetData = ws.GetFirstChild<SheetData>();





        }
    }
}
