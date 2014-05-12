using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace EventTimeNormalizer
{
    class Program
    {
        private static log4net.ILog log = log4net.LogManager.GetLogger(typeof(Program));

        static void Main(string[] args)
        {
            using (System.IO.Stream s = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("EventTimeNormalizer.log4net configuration.xml"))
            {
                log4net.Config.XmlConfigurator.Configure(s);
            }

            Parameters par = new Parameters();


            if (!CommandLine.Parser.Default.ParseArguments(args, par))
            {
                Console.WriteLine("Syntax error.\nSyntax:");

                var ht = CommandLine.Text.HelpText.AutoBuild(par);
                ht.Copyright = "Copyright 2014 ITPCfSQL";
                Console.WriteLine(ht.ToString());
                return;
            }

            int[] lKeyCells;
            {
                string[] sKeyCells = par.KeyCellsCSV.Split(new char[] { ',' });
                lKeyCells = new int[sKeyCells.Length];

                for (int i = 0; i < sKeyCells.Length; i++)
                {
                    if (!int.TryParse(sKeyCells[i], out lKeyCells[i]))
                    {
                        Console.WriteLine("Syntax error.\nCannot convert element {0:N0} (\"{1:S}\") of key CSV into a number.\n", i, sKeyCells[i]);
                        var ht = CommandLine.Text.HelpText.AutoBuild(par);
                        ht.Copyright = "Copyright 2014 ITPCfSQL";
                        Console.WriteLine(ht.ToString());
                        return;
                    }
                }
            }


            if (par.Verbose)
                ((log4net.Repository.Hierarchy.Logger)log.Logger).Level = log4net.Core.Level.Debug;

            log.DebugFormat("{0:S} v{1:S} started", System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString());

            SpreadsheetDocument sd;
            try
            {
                sd = SpreadsheetDocument.Open(par.InputExcelFile, false);
                //sd = SpreadsheetDocument.Open(@"D:\GIT\EventTimeNormalizer\EventTimeNormalizer\sample\data.xlsx", false);
                //sd = SpreadsheetDocument.Open(@"D:\GIT\EventTimeNormalizer\EventTimeNormalizer\sample\data - Copy.xlsx", false);            
            }
            catch (Exception exce)
            {
                Console.WriteLine("Cannot access " + par.InputExcelFile + " as an excel file. Exception was: " + exce.Message);
                return;
            }

            WorkbookPart wp = sd.Parts.First(item => item.OpenXmlPart is WorkbookPart).OpenXmlPart as WorkbookPart;
            Sheet sheetToRead = wp.Workbook.Descendants<Sheet>().First();
            WorksheetPart worksheetPart1 = (WorksheetPart)(wp.GetPartById(sheetToRead.Id));
            SheetData sheetData = worksheetPart1.Worksheet.GetFirstChild<SheetData>();

            SharedStringTable sharedStringTable = wp.SharedStringTablePart.SharedStringTable;

            #region Data read
            int iPos = 0;

            List<DateValueGroup> lDVGInput = new List<DateValueGroup>();

            foreach (Row row in sheetData.AsEnumerable())
            {
                if (iPos == 0 && par.UseFirstRowHeader)
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

                object oDate = ExtractValueFromCell(sharedStringTable, row.ElementAt(par.DataColumn) as Cell);
                if (oDate is DateTime)
                    dvp.Date = (DateTime)oDate;
                else
                    dvp.Date = DateTime.Parse(oDate.ToString());

                object oValue = ExtractValueFromCell(sharedStringTable, row.ElementAt(par.ValueColumn) as Cell);
                if (oValue is double)
                    dvp.Value = (double)oValue;
                else
                    dvp.Value = double.Parse(oValue.ToString());

                dvg.Add(dvp);
            }
            #endregion

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

            TimeSpan tsStep = new TimeSpan(0, 0, 1);

            long lTotalSteps = (int)((dtEnd - dtStart).TotalMilliseconds / tsStep.TotalMilliseconds);
            lTotalSteps *= lDVGInput.Count; // for each group
            #endregion

            //double dLastShownPerc = -10.0D;
            //double dSteps = 0.0D;

            #region Normalization
            List<DateValueGroup> lDVGOutput = new List<DateValueGroup>();

            foreach (DateValueGroup dvgInput in lDVGInput)
            {
                DateValueGroup dvgOutput = Normalize(dtStart, dtEnd,
                    new TimeSpan(0, 0, 1),

                    dvgInput);

                lDVGOutput.Add(dvgOutput);
            }
            #endregion

            GenerateOutput(par);
        }

        public static DateValueGroup Normalize(
            DateTime dtStart, DateTime dtEnd,
            TimeSpan tsStep,
            DateValueGroup dvgInput)
        {
            DateValueGroup dvgOutput = new DateValueGroup(dvgInput.Keys);

            DateTime dtCurrent = dtStart;

            long lPos = 0;

            double dLastValue = dvgInput.Values[0].Value;
            DateTime dLastDate = dvgInput.Values[0].Date;

            int iSrcPos = 1;


            while (dtCurrent < dtEnd)
            {
                DateTime dtNext = dtCurrent.Add(tsStep);
                double dAccumulated = 0;

                while ((dtCurrent < dtNext))
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
                        double deltaMSPerc = (dtToAnalyze - dtCurrent).TotalMilliseconds / tsStep.TotalMilliseconds;
                        dAccumulated += dPreviousValue * deltaMSPerc;

                        dtCurrent = dtToAnalyze;
                        iSrcPos++;
                    }
                    else
                    {
                        double deltaMSPerc = (dtNext - dtCurrent).TotalMilliseconds / tsStep.TotalMilliseconds;
                        dAccumulated += dPreviousValue * deltaMSPerc;

                        DateValuePair dvp = new DateValuePair(lPos++) { Date = dtNext.Subtract(tsStep) }; // VARIABILE!
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

        public static void GenerateOutput(Parameters par)
        {
            SpreadsheetDocument objExcelDoc = SpreadsheetDocument.Create(par.OutputExcelFile, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
            WorkbookPart wbp = objExcelDoc.AddWorkbookPart();
            WorksheetPart wsp = wbp.AddNewPart<WorksheetPart>();

            Workbook wb = new Workbook();

            FileVersion fv = new FileVersion();
            fv.ApplicationName = "Microsoft Office Excel";

            Worksheet workSheet = new Worksheet();

            Columns columns = new Columns();
            for (int i = 1; i <= 4; i++)
            {
                Column column = new Column();
                column.Min = Convert.ToUInt32(i);
                column.Max = Convert.ToUInt32(i);
                column.Width = 50;
                column.CustomWidth = false;
                columns.Append(column);
            }

            workSheet.Append(columns);


            SheetData sheetData = new SheetData();
            //{
            //    sheetData.Append(CreateContent(index, dr, columnSize.Count()));
            //    index++;
            //}

            workSheet.Append(sheetData);

            wsp.Worksheet = workSheet;

            Sheets sheets = new Sheets();
            Sheet sheet = new Sheet();
            sheet.Name = "nameee";
            sheet.SheetId = 1;
            sheet.Id = wbp.GetIdOfPart(wsp);

            sheets.Append(sheet);
            wb.Append(fv);
            wb.Append(sheets);

            objExcelDoc.WorkbookPart.Workbook = wb;
            objExcelDoc.WorkbookPart.Workbook.Save();
            objExcelDoc.Close();
        }

        protected static void AddPartXml(OpenXmlPart part, string xml)
        {
            using (System.IO.Stream stream = part.GetStream())
            {
                byte[] buffer = (new UTF8Encoding()).GetBytes(xml);
                stream.Write(buffer, 0, buffer.Length);
            }
        }

        private static Stylesheet CreateStylesheet()
        {

            Stylesheet ss = new Stylesheet();

            Fonts fts = new Fonts();

            DocumentFormat.OpenXml.Spreadsheet.Font ft = new DocumentFormat.OpenXml.Spreadsheet.Font();

            FontName ftn = new FontName();

            ftn.Val = StringValue.FromString("Calibri");

            DocumentFormat.OpenXml.Spreadsheet.FontSize ftsz = new DocumentFormat.OpenXml.Spreadsheet.FontSize();

            ftsz.Val = DoubleValue.FromDouble(11);

            ft.FontName = ftn;

            ft.FontSize = ftsz;

            fts.Append(ft);

            ft = new DocumentFormat.OpenXml.Spreadsheet.Font();

            ftn = new FontName();

            ftn.Val = StringValue.FromString("Palatino Linotype");

            ftsz = new DocumentFormat.OpenXml.Spreadsheet.FontSize();

            ftsz.Val = DoubleValue.FromDouble(18);

            ft.FontName = ftn;

            ft.FontSize = ftsz;

            fts.Append(ft);

            fts.Count = UInt32Value.FromUInt32((uint)fts.ChildElements.Count);

            Fills fills = new Fills();

            Fill fill;

            PatternFill patternFill;

            fill = new Fill();

            patternFill = new PatternFill();

            patternFill.PatternType = PatternValues.None;

            fill.PatternFill = patternFill;

            fills.Append(fill);

            fill = new Fill();

            patternFill = new PatternFill();

            patternFill.PatternType = PatternValues.Gray125;

            fill.PatternFill = patternFill;

            fills.Append(fill);

            fill = new Fill();

            patternFill = new PatternFill();

            patternFill.PatternType = PatternValues.Solid;

            patternFill.ForegroundColor = new ForegroundColor();

            patternFill.ForegroundColor.Rgb = HexBinaryValue.FromString("CDCDCD");

            patternFill.BackgroundColor = new BackgroundColor();

            patternFill.BackgroundColor.Rgb = patternFill.ForegroundColor.Rgb;

            fill.PatternFill = patternFill;

            fills.Append(fill);

            fills.Count = UInt32Value.FromUInt32((uint)fills.ChildElements.Count);

            Borders borders = new Borders();

            Border border = new Border();

            border.LeftBorder = new LeftBorder();

            border.RightBorder = new RightBorder();

            border.TopBorder = new TopBorder();

            border.BottomBorder = new BottomBorder();

            border.DiagonalBorder = new DiagonalBorder();

            borders.Append(border);

            border = new Border();

            border.LeftBorder = new LeftBorder();

            border.LeftBorder.Style = BorderStyleValues.Thin;

            border.RightBorder = new RightBorder();

            border.RightBorder.Style = BorderStyleValues.Thin;

            border.TopBorder = new TopBorder();

            border.TopBorder.Style = BorderStyleValues.Thin;

            border.BottomBorder = new BottomBorder();

            border.BottomBorder.Style = BorderStyleValues.Thin;

            border.DiagonalBorder = new DiagonalBorder();

            borders.Append(border);

            borders.Count = UInt32Value.FromUInt32((uint)borders.ChildElements.Count);

            CellStyleFormats csfs = new CellStyleFormats();

            CellFormat cf = new CellFormat();

            cf.NumberFormatId = 0;

            cf.FontId = 0;

            cf.FillId = 0;

            cf.BorderId = 0;

            csfs.Append(cf);

            csfs.Count = UInt32Value.FromUInt32((uint)csfs.ChildElements.Count);

            uint iExcelIndex = 164;

            NumberingFormats nfs = new NumberingFormats();

            CellFormats cfs = new CellFormats();

            NumberingFormat nfForcedText = new NumberingFormat();

            nfForcedText.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);

            nfForcedText.FormatCode = StringValue.FromString("@");

            nfs.Append(nfForcedText);

            cf = new CellFormat();

            cf.FontId = 0;

            cf.FillId = 0;

            cf.BorderId = 0;

            cf.FormatId = 0;

            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);

            cfs.Append(cf);

            cf = new CellFormat();

            cf.FontId = 0;

            cf.FillId = 0;

            cf.BorderId = 1;

            cf.FormatId = 0;

            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);

            cfs.Append(cf);

            cf = new CellFormat();

            cf.FontId = 0;

            cf.FillId = 0;

            cf.BorderId = 0;

            cf.FormatId = 0;

            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);

            cfs.Append(cf);

            cf = new CellFormat();

            cf.NumberFormatId = nfForcedText.NumberFormatId;

            cf.FontId = 0;

            cf.FillId = 0;

            cf.BorderId = 0;

            cf.FormatId = 0;

            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);

            cfs.Append(cf);

            cf = new CellFormat();

            cf.NumberFormatId = nfForcedText.NumberFormatId;

            cf.FontId = 1;

            cf.FillId = 0;

            cf.BorderId = 0;

            cf.FormatId = 0;

            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);

            cfs.Append(cf);

            cf = new CellFormat();

            cf.FontId = 0;

            cf.FillId = 0;

            cf.BorderId = 1;

            cf.FormatId = 0;

            cfs.Append(cf);

            cf = new CellFormat();

            cf.FontId = 0;

            cf.FillId = 2;

            cf.BorderId = 1;

            cf.FormatId = 0;

            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);

            cfs.Append(cf);

            cf = new CellFormat();

            cf.NumberFormatId = nfForcedText.NumberFormatId;

            cf.FontId = 0;

            cf.FillId = 2;

            cf.BorderId = 1;

            cf.FormatId = 0;

            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);

            cfs.Append(cf);

            ss.Append(nfs);

            ss.Append(fts);

            ss.Append(fills);

            ss.Append(borders);

            ss.Append(csfs);

            ss.Append(cfs);

            CellStyles css = new CellStyles();

            CellStyle cs = new CellStyle();

            cs.Name = StringValue.FromString("Normal");

            cs.FormatId = 0;

            cs.BuiltinId = 0;

            css.Append(cs);

            css.Count = UInt32Value.FromUInt32((uint)css.ChildElements.Count);

            ss.Append(css);

            DifferentialFormats dfs = new DifferentialFormats();

            dfs.Count = 0;

            ss.Append(dfs);

            TableStyles tss = new TableStyles();

            tss.Count = 0;

            tss.DefaultTableStyle = StringValue.FromString("TableStyleMedium9");

            tss.DefaultPivotStyle = StringValue.FromString("PivotStyleLight16");

            ss.Append(tss);

            return ss;

        }

    }
}
