﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using ITPCfSQL.EventTimeNormalizer;

namespace EventTimeNormalizer
{
    class Program
    {
        private static log4net.ILog log = log4net.LogManager.GetLogger(typeof(Program));

        private static long lTotalToShow;
        private static long lTotalShown;
        private static DateTime dtLastShown;

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
                log.Error("Cannot access " + par.InputExcelFile + " as an excel file. Exception was: " + exce.Message);
                return;
            }

            log.Info("Loading data");

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

                try
                {
                    object oDate = ExtractValueFromCell(sharedStringTable, row.ElementAt(par.DataColumn) as Cell);
                    if (oDate is DateTime)
                        dvp.Date = (DateTime)oDate;
                    else
                    {
                        DateTime dtTemp;
                        if (!DateTime.TryParse(oDate.ToString(), out dtTemp))
                            dtTemp = DateTime.FromOADate(double.Parse(oDate.ToString()));

                        dvp.Date = dtTemp;
                        //dvp.Date = DateTime.Parse(oDate.ToString());
                    }

                    object oValue = ExtractValueFromCell(sharedStringTable, row.ElementAt(par.ValueColumn) as Cell);
                    if (oValue is double)
                        dvp.Value = (double)oValue;
                    else
                        dvp.Value = double.Parse(oValue.ToString());

                    dvg.Add(dvp);
                }
                catch (Exception exce)
                {
                    log.WarnFormat("Exception in row {0:N}: {1:S}",  row.RowIndex, exce.Message);
                    //--iPos;
                }

            }
            #endregion

            #region Remove empty groups
            {
                List<DateValueGroup> lTemp = new List<DateValueGroup>();
                foreach (DateValueGroup dvg in lDVGInput)
                {
                    if (dvg.Values.Count == 0)
                        continue;

                    lTemp.Add(dvg);
                }

                lDVGInput = lTemp;
            }
            #endregion

            log.Info("Data load completed");

            log.Info("Sorting groups");
            Parallel.ForEach(lDVGInput, dvg => { dvg.SortDateValues(); });
            log.Info("Sorting groups completed");


            #region Test
            StreamNormalizer sn = new StreamNormalizer(new TimeSpan(0, 0, 1));
            sn.Start(lDVGInput[0][0]);
            for(int i=1;i<lDVGInput[0].Values.Count; i++)
            {
                var ret = sn.Push(lDVGInput[0][i]);
            }
            #endregion

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

            TimeSpan tsStep = new TimeSpan(0, 0, par.StepInSeconds);

            long lTotalSteps = (int)((dtEnd - dtStart).TotalMilliseconds / tsStep.TotalMilliseconds);
            log.InfoFormat("Normalization will create {0:N0} steps.", lTotalSteps);
            if (lTotalSteps > 1048576) // Excel limit
            {
                log.ErrorFormat("{0:N0} is above excel limit ({1:N0}). Try to shorten the timeframe or to increase the step in seconds. Aborting", lTotalSteps, 1048576);
                return;
            }
            //lTotalSteps *= lDVGInput.Count; // for each group
            #endregion

            //double dLastShownPerc = -10.0D;
            //double dSteps = 0.0D;

            log.Info("Starting normalization");
            #region Normalization

            List<DateValueGroup> lDVGOutput = new List<DateValueGroup>();

            Task<DateValueGroup>[] Tasks = new Task<DateValueGroup>[lDVGInput.Count];

            lTotalToShow = 100 * Tasks.Length;
            lTotalShown = 0;
            dtLastShown = DateTime.MinValue;

            for (int i = 0; i < lDVGInput.Count; i++)
            {
                Normalizer norm = new Normalizer(tsStep);
                norm.OnePercentStep += norm_OnePercentStep;

                DateValueGroup dvg = lDVGInput[i];

                Task<DateValueGroup> t = Task.Run(() => { return norm.Normalize(dtStart, dtEnd, dvg); });
                Tasks[i] = t;

                //Tasks[i] = Task.FromResult(norm.Normalize(dtStart, dtEnd, dvg));
            }

            Task.WaitAll(Tasks);

            log.Info("Normalization completed.");

            for (int i = 0; i < Tasks.Length; i++)
            {
                lDVGOutput.Add(Tasks[i].Result);
            }

            #endregion

            GenerateOutput(par, lDVGOutput);
        }

        static void norm_OnePercentStep(object sender)
        {
            System.Threading.Interlocked.Increment(ref lTotalShown);

            lock (log)
            {
                if ((DateTime.Now - dtLastShown).TotalMilliseconds > 100) // every 100 ms
                {
                    log.InfoFormat("{0:N2}% completed.", ((double)lTotalShown) / ((double)lTotalToShow) * 100);
                    dtLastShown = DateTime.Now;
                }
                else
                {
                    log.DebugFormat("{0:N2}% completed.", ((double)lTotalShown) / ((double)lTotalToShow) * 100);
                }
            }
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
                //Console.WriteLine(cell.CellReference + " - StyleIndex == " + cell.StyleIndex);
                return cell.InnerText;
            }

            return null;
        }

        public static void GenerateOutput(Parameters par, List<DateValueGroup> lDVGOutput)
        {
            log.Info("Output generation started.");

            SpreadsheetDocument objExcelDoc = SpreadsheetDocument.Create(par.OutputExcelFile, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
            WorkbookPart wbp = objExcelDoc.AddWorkbookPart();
            WorksheetPart wsp = wbp.AddNewPart<WorksheetPart>();

            Workbook wb = new Workbook();

            FileVersion fv = new FileVersion();
            fv.ApplicationName = "Microsoft Office Excel";

            Worksheet workSheet = new Worksheet();

            #region Stylesheet
            WorkbookStylesPart wbsp = wbp.AddNewPart<WorkbookStylesPart>();
            wbsp.Stylesheet = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };

            wbsp.Stylesheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            wbsp.Stylesheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            NumberingFormats numberingFormats1 = new NumberingFormats() { Count = (UInt32Value)2U };
            //NumberingFormat numberingFormat1 = new NumberingFormat() { NumberFormatId = (UInt32Value)164U, FormatCode = "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy" };
            NumberingFormat numberingFormat2 = new NumberingFormat() { NumberFormatId = (UInt32Value)165U, FormatCode = "[$-F400]h:mm:ss\\ AM/PM" };
            NumberingFormat numberingFormat1 = new NumberingFormat() { NumberFormatId = (UInt32Value)167U, FormatCode = "m/d/yy\\ h:mm;@" };


            numberingFormats1.Append(numberingFormat1);
            numberingFormats1.Append(numberingFormat2);

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)1U, KnownFonts = true };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            fonts1.Append(font1);

            Fills fills1 = new Fills() { Count = (UInt32Value)2U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            fills1.Append(fill1);
            fills1.Append(fill2);

            Borders borders1 = new Borders() { Count = (UInt32Value)1U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            borders1.Append(border1);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)4U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)14U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true };
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)164U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true };
            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)165U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true };

            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);
            cellFormats1.Append(cellFormat5);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            wbsp.Stylesheet.Append(numberingFormats1);
            wbsp.Stylesheet.Append(fonts1);
            wbsp.Stylesheet.Append(fills1);
            wbsp.Stylesheet.Append(borders1);
            wbsp.Stylesheet.Append(cellStyleFormats1);
            wbsp.Stylesheet.Append(cellFormats1);
            wbsp.Stylesheet.Append(cellStyles1);
            wbsp.Stylesheet.Append(differentialFormats1);
            wbsp.Stylesheet.Append(tableStyles1);

            wbsp.Stylesheet.Save();
            #endregion

            #region Columns definition
            Columns columns = new Columns();

            for (int i = 1; i <= lDVGOutput.Count + 1; i++)
            {
                Column column = new Column();
                column.Min = Convert.ToUInt32(i);
                column.Max = Convert.ToUInt32(i);
                column.Width = 50;
                column.CustomWidth = false;
                columns.Append(column);
            }

            workSheet.Append(columns);
            #endregion

            SheetData sheetData = new SheetData();

            #region Add header
            {
                Row objRow = new Row();

                objRow.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue("DateTime") });

                foreach (DateValueGroup dvg in lDVGOutput)
                {
                    StringBuilder sb = new StringBuilder();
                    for (int i = 0; i < dvg.Keys.Length; i++)
                    {
                        sb.Append(dvg.Keys[i].ToString());
                        if (i + 1 < dvg.Keys.Length)
                            sb.Append("-");
                    }

                    Cell objCell = new Cell() { DataType = CellValues.String, CellValue = new CellValue(sb.ToString()) };
                    objRow.Append(objCell);
                }

                sheetData.Append(objRow);
            }
            #endregion

            #region Add data
            {
                dtLastShown = DateTime.MinValue;
                lTotalShown = 0;
                lTotalToShow = lDVGOutput[0].Values.Count;

                for (int i = 0; i < lDVGOutput[0].Values.Count; i++)
                {
                    Row objRow = new Row();

                    objRow.Append(new Cell()
                    {
                        StyleIndex = 1,
                        CellValue = new CellValue(lDVGOutput[0].Values[i].Date.ToOADate().ToString())
                    });

                    foreach (DateValueGroup dvg in lDVGOutput)
                    {
                        Cell objCell = new Cell() { DataType = CellValues.Number, CellValue = new CellValue(dvg.Values[i].Value.ToString()) };
                        objRow.Append(objCell);
                    }

                    sheetData.Append(objRow);

                    #region Progress report
                    lTotalShown++;
                    if ((DateTime.Now - dtLastShown).TotalMilliseconds > 1000) // every s
                    {
                        log.InfoFormat("{0:N2}% completed.", ((double)lTotalShown) / ((double)lTotalToShow) * 100);
                        dtLastShown = DateTime.Now;
                    }
                    else
                    {
                        log.DebugFormat("{0:N2}% completed.", ((double)lTotalShown) / ((double)lTotalToShow) * 100);
                    }
                    #endregion
                }
            }
            #endregion

            log.Info("Appending sheet data");
            workSheet.Append(sheetData);

            wsp.Worksheet = workSheet;

            Sheets sheets = new Sheets();
            Sheet sheet = new Sheet();
            sheet.Name = "NormalizedData";
            sheet.SheetId = 1;
            sheet.Id = wbp.GetIdOfPart(wsp);

            log.Info("Appending sheet");
            sheets.Append(sheet);
            wb.Append(fv);
            wb.Append(sheets);

            log.Info("Appending Workbook");
            objExcelDoc.WorkbookPart.Workbook = wb;

            log.Info("Saving Workbook");
            objExcelDoc.WorkbookPart.Workbook.Save();
            
            log.Info("Closing excel");
            objExcelDoc.Close();
        }
    }
}
