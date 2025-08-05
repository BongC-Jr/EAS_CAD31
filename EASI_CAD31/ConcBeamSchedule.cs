using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/**
 * Google search 'C# Sheets V4'
 * 
 * Install Google.Apis.Sheets.V4
 * NuGet\Install-Package Google.Apis.Sheets.v4 -Version 1.60.0.2909
 */

using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Management;
using Newtonsoft.Json.Linq;

namespace EASI_CAD31
{
    public class ConcBeamSchedule
    {
        //https://docs.google.com/spreadsheets/d/1JnrIZvtSw9wcBCe9DjWcMk7eo1RCmf7uV5gQnpcumtE/edit#gid=0
        static string spreadSheetIDG = "1JnrIZvtSw9wcBCe9DjWcMk7eo1RCmf7uV5gQnpcumtE";

        static int startRowG = 5;
        static int endRowG = 10;
        static double textHgtG = 250;
        static double rowHgtScaleG = 2.5;

        public static void UpdateSheet(List<IList<object>> lioData, string spreadSheetID)
        {
            int iiNextAvailRow = NextAvailabeRow(spreadSheetID);
            iiNextAvailRow = iiNextAvailRow < 5 ? 5 : iiNextAvailRow;

            string strShtRng = "Beam Schedule!B" + iiNextAvailRow.ToString() + ":P" + (lioData.Count + iiNextAvailRow).ToString();
            ValueRange vrUpdateCell = new ValueRange();
            vrUpdateCell.MajorDimension = "ROWS";
            vrUpdateCell.Values = lioData;
            SpreadsheetsResource.ValuesResource.UpdateRequest urUpdateCell = DataGlobal.sheetsService.Spreadsheets.Values.Update(vrUpdateCell, spreadSheetID, strShtRng);
            urUpdateCell.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            UpdateValuesResponse exeUpdateCell = urUpdateCell.Execute();
        }

        public static int NextAvailabeRow(string spreadSheetID)
        {
            string strAddr = "Beam Schedule!C1";
            SpreadsheetsResource.ValuesResource.GetRequest grNextRow = DataGlobal.sheetsService.Spreadsheets.Values.Get(spreadSheetID, strAddr);
            ValueRange vrNextRow = grNextRow.Execute();
            IList<IList<object>> iioNextRow = vrNextRow.Values;
            return Convert.ToInt32(iioNextRow[0][0]);
        }


        [CommandMethod("GBS_GenerateBeamScheduleInGoogleSheet")]
        public static void GenerateBeamScheduleInGoogleSheet(string[] args)
        {
            Document actDoc = Application.DocumentManager.MdiActiveDocument;
            Database aCurDB = actDoc.Database;

            try
            {
                //// Define request parameters.
                //String spreadsheetId = "1JnrIZvtSw9wcBCe9DjWcMk7eo1RCmf7uV5gQnpcumtE";
                //String range = "Class Data!A2:E";
                //SpreadsheetsResource.ValuesResource.GetRequest request =
                //    service.Spreadsheets.Values.Get(spreadsheetId, range);
                //
                //// Prints the names and majors of students in a sample spreadsheet:
                //// https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
                //ValueRange response = request.Execute();
                //IList<IList<Object>> values = response.Values;
                //if (values == null || values.Count == 0)
                //{
                //    //Console.WriteLine("No data found.");
                //    actDoc.Editor.WriteMessage("\nNo data found.");
                //    return;
                //}
                ////Console.WriteLine("Name, Major");
                //actDoc.Editor.WriteMessage("\nName, Major");
                //foreach (var row in values)
                //{
                //    // Print columns A and E, which correspond to indices 0 and 4.
                //    //Console.WriteLine("{0}, {1}", row[0], row[4]);
                //    actDoc.Editor.WriteMessage("\n" + row[0].ToString() + "; " + row[4].ToString());
                //}
                actDoc.Editor.WriteMessage("\nMake a copy of the Google Workbook //https://docs.google.com/spreadsheets/d/1JnrIZvtSw9wcBCe9DjWcMk7eo1RCmf7uV5gQnpcumtE/edit#gid=0");
                actDoc.Editor.WriteMessage("\nThe following are the allowed text colors:");
                actDoc.Editor.WriteMessage("\n200, 1, 2, 3, 4, 5, 6, 7 and 32");

                PromptStringOptions psoGSheetID = new PromptStringOptions("");
                psoGSheetID.DefaultValue = spreadSheetIDG;
                psoGSheetID.Message = "\nEnter Google Spreadsheet ID: ";
                psoGSheetID.AllowSpaces = false;
                PromptResult prGSheetID = actDoc.Editor.GetString(psoGSheetID);

                string spreadSheetID = prGSheetID.StringResult;
                spreadSheetIDG = spreadSheetID;


                //Create a TypedValue array to define the filter criteria
                TypedValue[] tvBeamInf = new TypedValue[17];
                tvBeamInf.SetValue(new TypedValue((int)DxfCode.Start, "*TEXT"), 0);
                tvBeamInf.SetValue(new TypedValue((int)DxfCode.Operator, "<OR"), 1);
                tvBeamInf.SetValue(new TypedValue((int)DxfCode.LayerName, "S-BeamMarks"), 2);
                tvBeamInf.SetValue(new TypedValue((int)DxfCode.LayerName, "S-GBRebars"), 3);
                tvBeamInf.SetValue(new TypedValue((int)DxfCode.Operator, "OR>"), 4);
                tvBeamInf.SetValue(new TypedValue((int)DxfCode.Operator, "<OR"), 5);
                tvBeamInf.SetValue(new TypedValue((int)DxfCode.Color, 200), 6);
                tvBeamInf.SetValue(new TypedValue((int)DxfCode.Color, 1), 7);
                tvBeamInf.SetValue(new TypedValue((int)DxfCode.Color, 2), 8);
                tvBeamInf.SetValue(new TypedValue((int)DxfCode.Color, 3), 9);
                tvBeamInf.SetValue(new TypedValue((int)DxfCode.Color, 4), 10);
                tvBeamInf.SetValue(new TypedValue((int)DxfCode.Color, 5), 11);
                tvBeamInf.SetValue(new TypedValue((int)DxfCode.Color, 6), 12);
                tvBeamInf.SetValue(new TypedValue((int)DxfCode.Color, 7), 13);
                tvBeamInf.SetValue(new TypedValue((int)DxfCode.Color, 32), 14);
                tvBeamInf.SetValue(new TypedValue((int)DxfCode.Color, 8), 15);
                tvBeamInf.SetValue(new TypedValue((int)DxfCode.Operator, "OR>"), 16);

                SelectionFilter sfBeamInf = new SelectionFilter(tvBeamInf);
                PromptSelectionResult psrBeamInf = actDoc.Editor.GetSelection(sfBeamInf);

                if (psrBeamInf.Status != PromptStatus.OK)
                {
                    actDoc.Editor.WriteMessage("\nSelection cancelled.");
                    return;
                }

                SelectionSet ssBeamInf = psrBeamInf.Value;

                actDoc.Editor.WriteMessage("\nNumber of selected objects: " + ssBeamInf.Count.ToString());

                using (Transaction trBeamInf = aCurDB.TransactionManager.StartTransaction())
                {
                    List<string> liBeamMarks = new List<string>();
                    //Put beam marks in a list and sort it.
                    foreach (SelectedObject soBeamInf in ssBeamInf)
                    {
                        if (soBeamInf != null)
                        {
                            Entity enBeamInf = trBeamInf.GetObject(soBeamInf.ObjectId, OpenMode.ForRead) as Entity;
                            Regex rgxBeamInf = new Regex(@"^C?[BG][0-9A-Z][0-9A-Z\.]{0,5}-[0-9]+");

                            if (enBeamInf is DBText)
                            {
                                DBText dbtBeamInf = enBeamInf as DBText;
                                string textCont = dbtBeamInf.TextString;


                                if (rgxBeamInf.Match(textCont).Success)
                                {
                                    liBeamMarks.Add(textCont);
                                }
                            }
                            if (enBeamInf is MText)
                            {
                                MText mtxBeamInf = enBeamInf as MText;
                                string mtxtCont = mtxBeamInf.Contents;

                                if (rgxBeamInf.Match(mtxtCont).Success)
                                {
                                    liBeamMarks.Add(mtxtCont);
                                }
                            }
                        }
                    }

                    liBeamMarks.Sort();
                    //for(int ii=0; ii<liBeamMarks.Count; ii++)
                    //{
                    //    actDoc.Editor.WriteMessage("\nBeam mark " + ii + "; " + liBeamMarks[ii]);
                    //}

                    string[,] arr2BeamSched = new string[liBeamMarks.Count, 15];
                    //actDoc.Editor.WriteMessage("\nTotal row: " + liBeamMarks.Count);

                    /**
                     * 10 June 2022
                     * Once the list of beam marks is sorted, iterate thru the selection again. Get the extended data of each S-GBMarks; and based on the extended data
                     * value, get the index number in the list. The index number will become the basis of row number position to write in the google sheet.
                     * */
                    foreach (SelectedObject soBeamInf2 in ssBeamInf)
                    {
                        if (soBeamInf2 != null)
                        {
                            Entity entBeamInf2 = trBeamInf.GetObject(soBeamInf2.ObjectId, OpenMode.ForRead) as Entity;
                            ResultBuffer rbBeamInf = entBeamInf2.GetXDataForApplication("SEAS_GBRebInf");
                            if (rbBeamInf != null)
                            {
                                foreach (TypedValue tvExDa in rbBeamInf)
                                {
                                    if ((int)tvExDa.TypeCode == 1000)
                                    {
                                        string exdaValue = (string)tvExDa.Value;
                                        int rawRow = liBeamMarks.IndexOf(exdaValue);
                                        //Additional code here
                                        int iiRow = rawRow;
                                        //actDoc.Editor.WriteMessage("\niiRow: "+ iiRow.ToString());

                                        /**
                                         * 10 June 2022
                                         * Identify the column assignment of each text in the Google Sheet by its contents. Afterwhich, the text content shall be written
                                         * in the specified Google Sheet
                                         * */
                                        string acText = "";
                                        if (entBeamInf2 is DBText)
                                        {
                                            DBText dbTxt2 = entBeamInf2 as DBText;
                                            acText = dbTxt2.TextString;
                                        }
                                        if (entBeamInf2 is MText)
                                        {
                                            MText mText2 = entBeamInf2 as MText;
                                            acText = mText2.Contents;
                                        }

                                        string celVal = "-NV-";
                                        Regex rgxEnd1 = new Regex(@"^E1:[123][0-9]-[1-9][0-9]{0,2}\/[1-9][0-9]{0,2}");
                                        if (rgxEnd1.Match(acText).Success && iiRow >= 0)
                                        {
                                            Regex rgxEnd1Reb = new Regex(@"[1-9][0-9]{0,2}\/[1-9][0-9]{0,2}$");
                                            celVal = rgxEnd1Reb.Match(acText).Value;
                                            string[] sep1 = { "/" };
                                            string[] arrCelVal = celVal.Split(sep1, StringSplitOptions.RemoveEmptyEntries);

                                            arr2BeamSched[iiRow, 7] = arrCelVal[0];
                                            arr2BeamSched[iiRow, 8] = arrCelVal[1];

                                            Regex rgxDia1 = new Regex(@"[1-9][0-9]?-");
                                            string gbDia1 = rgxDia1.Match(acText).Value.Replace("-", "");
                                            arr2BeamSched[iiRow, 6] = gbDia1;
                                        }

                                        Regex rgxEnd2 = new Regex(@"^E2:[123][0-9]-[1-9][0-9]{0,2}\/[1-9][0-9]{0,2}");
                                        if (rgxEnd2.Match(acText).Success && iiRow >= 0)
                                        {
                                            Regex rgxEnd2Reb = new Regex(@"[1-9][0-9]{0,2}\/[1-9][0-9]{0,2}$");
                                            celVal = rgxEnd2Reb.Match(acText).Value;
                                            string[] sep1 = { "/" };
                                            string[] arrCelVal = celVal.Split(sep1, StringSplitOptions.RemoveEmptyEntries);

                                            arr2BeamSched[iiRow, 11] = arrCelVal[0];
                                            arr2BeamSched[iiRow, 12] = arrCelVal[1];
                                        }

                                        Regex rgxMid = new Regex(@"^MS:[123][0-9]-[1-9][0-9]{0,2}\/[1-9][0-9]{0,2}");
                                        if (rgxMid.Match(acText).Success && iiRow >= 0)
                                        {
                                            Regex rgxMidReb = new Regex(@"[1-9][0-9]{0,2}\/[1-9][0-9]{0,2}$");
                                            celVal = rgxMidReb.Match(acText).Value;
                                            string[] sep1 = { "/" };
                                            string[] arrCelVal = celVal.Split(sep1, StringSplitOptions.RemoveEmptyEntries);

                                            arr2BeamSched[iiRow, 9] = arrCelVal[0];
                                            arr2BeamSched[iiRow, 10] = arrCelVal[1];
                                        }

                                        Regex rgxBeamMark = new Regex(@"^C?[BG][0-9A-Z\.][0-9A-Z\.]{0,6}-[0-9][0-9]{0,2}");
                                        if (rgxBeamMark.Match(acText).Success && iiRow >= 0)
                                        {
                                            //Get beam level
                                            Regex rgxLevel = new Regex(@"[BG][0-9A-Z\.][0-9A-Z\.]{0,6}-");
                                            string gbLevel = rgxLevel.Match(acText).Value.Replace("-", "");
                                            //actDoc.Editor.WriteMessage("\nGB Level: " + gbLevel);
                                            gbLevel = gbLevel.Replace("B", "");
                                            gbLevel = gbLevel.Replace("G", "");
                                            arr2BeamSched[iiRow, 0] = gbLevel;

                                            //Get beam mark
                                            if (exdaValue == acText)
                                            {
                                                arr2BeamSched[iiRow, 1] = acText;
                                            }
                                            else
                                            {
                                                actDoc.Editor.WriteMessage("\nData value " + exdaValue + " do not match with text value " + acText);
                                            }

                                            //Get beam number
                                            Regex rgxgbNum = new Regex(@"-[0-9][0-9]{0,2}");
                                            string gbNum = rgxgbNum.Match(acText).Value.Replace("-", "");
                                            arr2BeamSched[iiRow, 2] = gbNum;
                                        }

                                        Regex rgxGBDim = new Regex(@"^\([0-9][0-9]{1,3}X[0-9][0-9]{1,3}\)$");
                                        if (rgxGBDim.Match(acText).Success && iiRow >= 0)
                                        {
                                            //Beam width
                                            Regex rgxWidth = new Regex(@"^\([0-9][0-9]{1,3}");
                                            string gbWidth = rgxWidth.Match(acText).Value.Replace("(", "");
                                            arr2BeamSched[iiRow, 3] = gbWidth;

                                            //Beam depth
                                            Regex rgxDepth = new Regex(@"[0-9][0-9]{1,3}\)$");
                                            string gbDepth = rgxDepth.Match(acText).Value.Replace(")", "");
                                            arr2BeamSched[iiRow, 4] = gbDepth;
                                        }

                                        Regex rgxLu = new Regex(@"Lu=[0-9\.]+");
                                        if (rgxLu.Match(acText).Success && iiRow >= 0)
                                        {
                                            string gbLu = Regex.Replace(acText, @"Lu=", "").ToString();
                                            arr2BeamSched[iiRow, 5] = gbLu;
                                        }

                                        Regex rgxStirr = new Regex(@"^STIRR\.:\(?[A-Z][A-Z]?\)?.*");
                                        if (rgxStirr.Match(acText).Success && iiRow >= 0)
                                        {
                                            Regex rgxGetStirr = new Regex(@":\(?[A-Z][A-Z]?\(?");
                                            string gbStirr = rgxGetStirr.Match(acText).Value.Replace(":", "");
                                            gbStirr = gbStirr.Replace("(", "");
                                            gbStirr = gbStirr.Replace(")", "");
                                            arr2BeamSched[iiRow, 14] = gbStirr;

                                        }

                                        if (iiRow >= 0)
                                        {
                                            arr2BeamSched[iiRow, 13] = "";
                                        }

                                    }
                                }
                            }
                            else
                            {
                                actDoc.Editor.WriteMessage("\nExtended data not available.");
                                entBeamInf2.Highlight();
                            }


                        }
                        else
                        {
                            actDoc.Editor.WriteMessage("\nSelection object empty.");
                        }
                    }

                    int iN = arr2BeamSched.GetLength(0);
                    int jN = arr2BeamSched.GetLength(1);

                    List<IList<object>> lioValues = new List<IList<object>>();
                    for (int ii = 0; ii < iN; ii++)
                    {
                        List<object> loValues = new List<object>();
                        for (int jj = 0; jj < jN; jj++)
                        {
                            loValues.Add(arr2BeamSched[ii, jj]);
                        }
                        lioValues.Add(loValues);
                    }

                    UpdateSheet(lioValues, spreadSheetID);

                    trBeamInf.Commit();
                }

                actDoc.Editor.WriteMessage("\n\n\nBeam schedule created in Google Sheet.");
            }
            catch (FileNotFoundException e)
            {
                //Console.WriteLine(e.Message);
                actDoc.Editor.WriteMessage("\n" + e.Message);
            }
        }

        [CommandMethod("GBSC_GenerateBeamScheduleInCAD")]
        public static void GenerateBeamScheduleInCAD(string[] args)
        {
            Document actDoc = Application.DocumentManager.MdiActiveDocument;
            Database aCurDB = actDoc.Database;


            try
            {
                PromptIntegerOptions pioStartRow = new PromptIntegerOptions("");
                pioStartRow.DefaultValue = startRowG;
                pioStartRow.Message = "\nEnter start row: ";
                pioStartRow.AllowNegative = false;
                pioStartRow.AllowZero = false;
                pioStartRow.AllowNone = true;
                PromptIntegerResult pirStartRow = actDoc.Editor.GetInteger(pioStartRow);
                int startRow = pirStartRow.Value;
                startRowG = startRow;

                PromptIntegerOptions pioEndRow = new PromptIntegerOptions("");
                pioEndRow.DefaultValue = endRowG;
                pioEndRow.Message = "\nEnter end row: ";
                pioEndRow.AllowNegative = false;
                pioEndRow.AllowZero = false;
                pioEndRow.AllowNone = true;
                PromptIntegerResult pirEndRow = actDoc.Editor.GetInteger(pioEndRow);
                int endRow = pirEndRow.Value;
                endRowG = endRow;

                PromptStringOptions psoGSheetID = new PromptStringOptions("");
                psoGSheetID.DefaultValue = spreadSheetIDG;
                psoGSheetID.Message = "\nEnter Google Spreadsheet ID: ";
                psoGSheetID.AllowSpaces = false;
                PromptResult prGSheetID = actDoc.Editor.GetString(psoGSheetID);

                PromptPointOptions ppoPnt1 = new PromptPointOptions("\nSpecify point 1: ");
                PromptPointResult pprPnt1 = actDoc.Editor.GetPoint(ppoPnt1);
                if (pprPnt1.Status != PromptStatus.OK)
                {
                    actDoc.Editor.WriteMessage("\nSelection cancelled.");
                    return;
                }
                Point3d p3dPnt1 = pprPnt1.Value;

                PromptPointOptions ppoPnt2 = new PromptPointOptions("\nSpecify point 2: ");
                ppoPnt2.BasePoint = p3dPnt1;
                ppoPnt2.UseBasePoint = true;
                PromptPointResult pprPnt2 = actDoc.Editor.GetPoint(ppoPnt2);
                if (pprPnt2.Status != PromptStatus.OK)
                {
                    actDoc.Editor.WriteMessage("\nSelection cancelled.");
                    return;
                }
                Point3d p3dPnt2 = pprPnt2.Value;

                PromptPointOptions ppoPnt3 = new PromptPointOptions("\nSpecify point 3: ");
                ppoPnt3.BasePoint = p3dPnt1;
                ppoPnt3.UseBasePoint = true;
                PromptPointResult pprPnt3 = actDoc.Editor.GetPoint(ppoPnt3);
                if (pprPnt3.Status != PromptStatus.OK)
                {
                    actDoc.Editor.WriteMessage("\nSelection cancelled.");
                    return;
                }
                Point3d p3dPnt3 = pprPnt3.Value;

                Vector3d v3dPnt12 = p3dPnt2 - p3dPnt1;
                v3dPnt12 = v3dPnt12.GetNormal();

                Vector3d v3dPnt23 = p3dPnt3 - p3dPnt2;
                v3dPnt23 = v3dPnt23.GetNormal();

                PromptDoubleOptions pdoTextHgt = new PromptDoubleOptions("\nEnter text height: ");
                pdoTextHgt.DefaultValue = textHgtG;
                pdoTextHgt.AllowNegative = false;
                pdoTextHgt.AllowZero = false;
                pdoTextHgt.AllowNone = true;
                PromptDoubleResult pdrTextHgt = actDoc.Editor.GetDouble(pdoTextHgt);
                if (pdrTextHgt.Status != PromptStatus.OK)
                {
                    actDoc.Editor.WriteMessage("\nSelection cancelled.");
                    return;
                }
                double textHgt = pdrTextHgt.Value;
                textHgtG = textHgt;


                string spreadSheetID = prGSheetID.StringResult;
                spreadSheetIDG = spreadSheetID;

                IList<IList<object>> iioShtRngData = GetRange(startRow, endRow, spreadSheetID);
                IList<IList<object>> iioShtRngColW = GetRange(4, 4, spreadSheetID); //Column widths
                IList<IList<object>> iioShtRngHead = GetRange(3, 3, spreadSheetID); //Header

                iioShtRngData = iioShtRngHead.Concat(iioShtRngData).ToList();

                double horDist = p3dPnt1.DistanceTo(p3dPnt2);
                double verDist = Math.Round(p3dPnt2.DistanceTo(p3dPnt3));
                actDoc.Editor.WriteMessage($"\nVertical distance: {verDist}");
                using (Transaction trBeamSch = aCurDB.TransactionManager.StartTransaction())
                {
                    BlockTable blktbl;
                    blktbl = trBeamSch.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord bltrec;
                    bltrec = trBeamSch.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    int ii = 0;
                    double iiVertDist = 0.0;
                    while (ii < iioShtRngData.Count &&  iiVertDist < verDist)
                    {
                        Point3d p3diiP1 = p3dPnt1 + v3dPnt23 * ii * (rowHgtScaleG * textHgt);
                        Point3d p3diiP2 = p3dPnt2 + v3dPnt23 * ii * (rowHgtScaleG * textHgt);
                        Line lnHor = new Line(p3diiP1, p3diiP2);
                        lnHor.ColorIndex = 1;
                        bltrec.AppendEntity(lnHor);
                        trBeamSch.AddNewlyCreatedDBObject(lnHor, true);

                        Point3d p3djjP1 = p3diiP1;
                        Point3d p3djjP2 = p3diiP1 + v3dPnt23 * (rowHgtScaleG * textHgt);
                        iiVertDist = Math.Round(p3dPnt1.DistanceTo(p3djjP2));
                        
                        actDoc.Editor.WriteMessage($"\n{ii}");
                        Line lnVerj0 = new Line(p3djjP1, p3djjP2);
                        lnVerj0.ColorIndex = 1;
                        bltrec.AppendEntity(lnVerj0);
                        trBeamSch.AddNewlyCreatedDBObject(lnVerj0, true);

                        for (int jj = 1; jj < iioShtRngData[ii].Count; jj++)
                        {
                            actDoc.Editor.WriteMessage($"; {iioShtRngData[ii][jj].ToString()}");
                            //Double colWidthScale = iioShtRngColW[0][jj].ToString() == "" ? 1.0 : Convert.ToDouble(iioShtRngColW[0][jj].ToString(), CultureInfo.InvariantCulture);
                            Double colWidthScale = Convert.ToDouble(iioShtRngColW[0][jj]);
                            
                            p3djjP1 = p3djjP1 + v3dPnt12 * (horDist * colWidthScale / 100.0);
                            p3djjP2 = p3djjP2 + v3dPnt12 * (horDist * colWidthScale / 100.0);

                            Line lnVerjk = new Line(p3djjP1, p3djjP2);
                            lnVerjk.ColorIndex = 1;
                            bltrec.AppendEntity(lnVerjk);
                            trBeamSch.AddNewlyCreatedDBObject(lnVerjk, true);

                            Point3d txInsPt = p3djjP1 + v3dPnt23 * (rowHgtScaleG * textHgt / 2) + v3dPnt12 * (0.6 * textHgt) - v3dPnt12 * (horDist * colWidthScale / 100.0);
                            DBText acText = new DBText();
                            acText.SetDatabaseDefaults();
                            acText.Position = txInsPt;
                            acText.Height = textHgt;
                            acText.TextString = iioShtRngData[ii][jj].ToString();

                            acText.ColorIndex = 3;
                            acText.HorizontalMode = TextHorizontalMode.TextCenter;
                            acText.VerticalMode = TextVerticalMode.TextVerticalMid;
                            acText.AlignmentPoint = txInsPt;
                            acText.Justify = AttachmentPoint.MiddleLeft;
                            acText.WidthFactor = 1.0;

                            bltrec.AppendEntity(acText);
                            trBeamSch.AddNewlyCreatedDBObject(acText, true);
                        }
                        actDoc.Editor.WriteMessage($"; ii-Distance= {iiVertDist}");

                        ii++;
                    }
                    trBeamSch.Commit();
                }

                

            }
            catch (FileNotFoundException e)
            {
                //Console.WriteLine(e.Message);
                actDoc.Editor.WriteMessage("\n" + e.Message);
            }


        }

        public static IList<IList<object>> GetRange(int startRow, int endRow, string spreadSheetID)
        {
            string shtRange = "Beam Schedule!A" + startRow.ToString() + ":R" + endRow.ToString();
            SpreadsheetsResource.ValuesResource.GetRequest grBeamSched = DataGlobal.sheetsService.Spreadsheets.Values.Get(spreadSheetID, shtRange);
            ValueRange vrBeamSched = grBeamSched.Execute();
            IList<IList<object>> iioBeamSched = vrBeamSched.Values;

            return iioBeamSched;
        }
    }
}
