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
            //SpreadsheetDocument sd = SpreadsheetDocument.Open(@"D:\GIT\EventTimeNormalizer\EventTimeNormalizer\sample\data - Copy.xlsx", false);

            WorkbookPart wp = sd.Parts.First(item => item.OpenXmlPart is WorkbookPart).OpenXmlPart as WorkbookPart;
            Sheet sheet = wp.Workbook.Descendants<Sheet>().First();
            WorksheetPart worksheetPart1 = (WorksheetPart)(wp.GetPartById(sheet.Id));
            SheetData sheetData = worksheetPart1.Worksheet.GetFirstChild<SheetData>();

            SharedStringTable sharedStringTable = wp.SharedStringTablePart.SharedStringTable;

            int iPos = 0;
            int iValueCell = 2;
            int iDataCell = 3;

            int[] lKeyCells = new int[] { 0, 1 };

            List<DateValueGroup> lDVGInput = new List<DateValueGroup>();

            foreach (Row row in sheetData.AsEnumerable())
            {
                if (iPos == 0 && fFirstRowHeader)
                {
                    iPos++;
                    continue;
                }

                #region Find correct DateValueGroup based on the keys
                object[] keys = new object[lKeyCells.Length];
                for (int i = 0; i < lKeyCells.Length; i++)
                    keys[i] = ExtractValueFromCell(sharedStringTable, row.ElementAt(i) as Cell);

                DateValueGroup dvg = lDVGInput.FirstOrDefault(item => item.MatchKeys(keys));
                if (dvg == null)
                {
                    dvg = new DateValueGroup(keys);
                    lDVGInput.Add(dvg);
                }
                #endregion

                DateValuePair dvp = new DateValuePair(iPos++);

                object oDate = ExtractValueFromCell(sharedStringTable, row.ElementAt(iDataCell) as Cell);
                if (oDate is DateTime)
                    dvp.Date = (DateTime)oDate;
                else
                    dvp.Date = DateTime.Parse(oDate.ToString());

                object oValue = ExtractValueFromCell(sharedStringTable, row.ElementAt(iValueCell) as Cell);
                if (oValue is double)
                    dvp.Value = (double)oValue;
                else
                    dvp.Value = double.Parse(oValue.ToString());

                dvg.Add(dvp);
            }

            Parallel.ForEach(lDVGInput, dvg => { dvg.SortDateValues(); });

            #region Find starting time and end time
            DateTime dtStart = DateTime.MaxValue;
            DateTime dtEnd = DateTime.MinValue;

            foreach (DateValueGroup dvg in lDVGInput)
            {
                if (dtStart > dvg[0].Date)
                    dtStart = dvg[0].Date;
                if (dtEnd < dvg[dvg.Values.Count - 1].Date)
                    dtEnd = dvg[dvg.Values.Count - 1].Date;
            }

            log.InfoFormat("Min event time is {0:S}, max is {1:S}. Timespan is {2:S}.",
                dtStart.ToString(), dtEnd.ToString(), ((dtEnd - dtStart).ToString()));

            long lTotalSteps = (int)((dtEnd - dtStart).TotalMilliseconds / 1000); // VARIABLE
            lTotalSteps *= lDVGInput.Count; // for each group
            #endregion

            //double dLastShownPerc = -10.0D;
            //double dSteps = 0.0D;

            List<DateValueGroup> lDVGOutput = new List<DateValueGroup>();

            foreach (DateValueGroup dvgInput in lDVGInput)
            {
                DateValueGroup dvgOutput = Normalize(dtStart, dtEnd, 
                    new TimeSpan(0,0,1),
                    
                    dvgInput);

                lDVGOutput.Add(dvgOutput);
            }
        }

        public static DateValueGroup Normalize(
            DateTime dtStart, DateTime dtEnd,
            TimeSpan tsStep,
            DateValueGroup dvgInput)
        {
            //// Move first value to start
            //dvgInput.Values[0].Date = dtStart;

            // Add last value to end
            //dvgInput.Add(new DateValuePair(-1) { Value = dvgInput.Values[//dvgInput.Values.Count - 1].Value, Date = dtEnd });

            DateValueGroup dvgOutput = new DateValueGroup(dvgInput.Keys);

            DateTime dtCurrent = dtStart;

            long lPos = 0;

            double dLastValue = dvgInput.Values[0].Value;
            DateTime dLastDate = dvgInput.Values[0].Date;

            int iSrcPos = 1;


            while (dtCurrent < dtEnd)
            {
                DateTime dtNext = dtCurrent.AddSeconds(1); // VARIABILE
                double dAccumulated = 0;

                while ((dtCurrent < dtNext))// && (iSrcPos < dvgInput.Values.Count))
                {
                    DateTime dtToAnalyze;
                    double dPreviousValue;

                    if (iSrcPos < dvgInput.Values.Count)
                    {
                        dtToAnalyze = dvgInput.Values[iSrcPos].Date;
                        dPreviousValue = dvgInput.Values[iSrcPos - 1].Value;
                    }
                    else
                    {
                        dtToAnalyze = DateTime.MaxValue;
                        dPreviousValue = dvgInput.Values[dvgInput.Values.Count - 1].Value;
                    }

                    if (dtToAnalyze < dtNext)
                    {
                        double deltaMSPerc = (dtToAnalyze - dtCurrent).TotalMilliseconds / 1000; // VARIABILE
                        dAccumulated += dPreviousValue * deltaMSPerc;

                        dtCurrent = dtToAnalyze;
                        iSrcPos++;
                    }
                    else
                    {
                        double deltaMSPerc = (dtNext - dtCurrent).TotalMilliseconds / 1000; // VARIABILE
                        dAccumulated += dPreviousValue * deltaMSPerc;

                        DateValuePair dvp = new DateValuePair(lPos++) { Date = dtNext.Subtract(TimeSpan.FromSeconds(1)) }; // VARIABILE!
                        dvp.Value = dAccumulated;
                        dvgOutput.Add(dvp);

                        //#region Calculate % completed
                        //dSteps++;
                        //double dCurPerc = (dSteps / lTotalSteps) * 100;
                        //if (dCurPerc - dLastShownPerc > 1.0D)
                        //{
                        //    log.InfoFormat("{0:N0}% completed.", dCurPerc);
                        //    dLastShownPerc = dCurPerc;
                        //}
                        //#endregion

                        dAccumulated = 0.0D;
                        dtCurrent = dtNext;
                    }
                }
            }

            return dvgOutput;
        }

        public static object ExtractValueFromCell(SharedStringTable sharedStringTable, Cell cell)
        {
            if (cell.DataType != null)
            {
                if (cell.DataType == CellValues.Date)
                {
                    string s = cell.CellValue.Text;
                    DateTime dt = DateTime.Parse(s);
                    return dt;
                }
                else if (cell.DataType == CellValues.SharedString)
                {
                    SharedStringItem ssi = sharedStringTable.ChildElements[int.Parse(cell.CellValue.Text)] as SharedStringItem;
                    return ssi.Text.Text;
                }

                else if ((cell.StyleIndex != null) && (cell.StyleIndex == 2))
                {
                    DateTime dt = DateTime.FromOADate(double.Parse(cell.CellValue.InnerXml));
                    return dt;
                }
                else if (cell.CellValue != null)
                {
                    return cell.CellValue.InnerText;
                }
            }
            else
            {
                return cell.InnerText;
            }

            return null;
        }
    }
}
