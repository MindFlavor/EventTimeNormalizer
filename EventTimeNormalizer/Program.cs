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

                        }
                    } else
                    {
                        DateTime dt = DateTime.FromOADate(double.Parse(cell.CellValue.InnerXml));

                        sb2.Append(cell.CellValue.InnerXml);
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
