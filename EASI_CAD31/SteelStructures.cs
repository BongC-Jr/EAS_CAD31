using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace EASI_CAD31
{
    public class SteelStructures
    {
        /**
         * Required statements
         */
        static Document actDoc = Application.DocumentManager.MdiActiveDocument;
        static Database aCurDB = actDoc.Database;
        static CCData ccData = new CCData();
        static SEASTools seasTools = new SEASTools();

        //Default sheet address
        static string shAddrG = "Sheet1!A2";

        //GSCS command
        static double colHgtG = 3200.00;

        // Default google sheet
        /* https://docs.google.com/spreadsheets/d/1ZGy2rFLEkVEQF4DunspHXn5umcu14udCRT31YycHrQs/edit#gid=0 */
        static string gshtIdG = "1ZGy2rFLEkVEQF4DunspHXn5umcu14udCRT31YycHrQs";

        /**
         * Date added: 21 February 2024
         * Added by: Engr Bernardo Cabebe Jr.
         * Venue: 3407 Cityland Pasong Tamo Tower
         */
        [CommandMethod("GSCS_GetSteelColumnSections")]
        public static void GetSteelColumnSections()
        {
            try
            {
                if (!ccData.isLicenseActive())
                {
                    actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                    return;
                }
                actDoc.Editor.WriteMessage("\nCopy google worksheet https://docs.google.com/spreadsheets/d/1ZGy2rFLEkVEQF4DunspHXn5umcu14udCRT31YycHrQs/edit#gid=0");
                actDoc.Editor.WriteMessage("\nUse color 145 for column MULTILINE TEXT");

                //Select column texts
                TypedValue[] tvFilterN2w = new TypedValue[2];
                tvFilterN2w[0] = new TypedValue((int)DxfCode.Start, "MTEXT");
                tvFilterN2w[1] = new TypedValue((int)DxfCode.Color, 145);
                SelectionFilter sfMTxt = new SelectionFilter(tvFilterN2w);

                PromptSelectionOptions psoMTxt = new PromptSelectionOptions();
                psoMTxt.MessageForAdding = "\nSelect steel column text: ";

                PromptSelectionResult psrMTxt = actDoc.Editor.GetSelection(psoMTxt, sfMTxt);
                if (psrMTxt.Status != PromptStatus.OK) return;

                SelectionSet ssMTxt = psrMTxt.Value;
                actDoc.Editor.WriteMessage("\nSelected text: {0}", ssMTxt.Count);

                //Get column height
                PromptDoubleOptions pdOptN2w = new PromptDoubleOptions("\nEnter column height: ");
                pdOptN2w.DefaultValue = colHgtG;
                pdOptN2w.AllowNegative = false;
                pdOptN2w.AllowZero = false;

                PromptDoubleResult pdrCHgt = actDoc.Editor.GetDouble(pdOptN2w);
                if (pdrCHgt.Status != PromptStatus.OK) return;
                double colHgt = pdrCHgt.Value;
                colHgtG = colHgt;

                //Enter range address
                PromptStringOptions psoSheetAddr = new PromptStringOptions("");
                psoSheetAddr.DefaultValue = shAddrG;
                psoSheetAddr.Message = "\nEnter cell address: ";
                psoSheetAddr.AllowSpaces = false;
                PromptResult prSheetAddr = actDoc.Editor.GetString(psoSheetAddr);

                string sheetAddr = prSheetAddr.StringResult;
                shAddrG = sheetAddr;
                actDoc.Editor.WriteMessage("\nCell address: {0}", shAddrG);

                PromptStringOptions psoGSheetID = new PromptStringOptions("");
                psoGSheetID.DefaultValue = gshtIdG;
                psoGSheetID.Message = "\nEnter Google Spreadsheet ID: ";
                psoGSheetID.AllowSpaces = false;
                PromptResult prGSheetID = actDoc.Editor.GetString(psoGSheetID);

                string spreadSheetID = prGSheetID.StringResult;
                gshtIdG = spreadSheetID;
                actDoc.Editor.WriteMessage("\nSpreadsheet ID: {0}", gshtIdG);

                object[,] arr2ObjN2w = new object[ssMTxt.Count, 2];
                int aaRow = 0;
                using (Transaction transctionN2w = aCurDB.TransactionManager.StartTransaction())
                {
                    foreach(SelectedObject selObjN2w in ssMTxt)
                    {
                        Entity entityN2w = (Entity)transctionN2w.GetObject(selObjN2w.ObjectId, OpenMode.ForRead);
                        // string objTypeN2w = entityN2w.GetType().Name.ToUpper();
                        // actDoc.Editor.WriteMessage("\nObject type: {0}", objTypeN2w);

                        MText mTxtN2w = (MText)entityN2w;
                        string txtCont = mTxtN2w.Contents;
                        arr2ObjN2w[aaRow,0] = txtCont;
                        arr2ObjN2w[aaRow, 1] = colHgt;
                        aaRow++;

                        mTxtN2w.UpgradeOpen();
                        mTxtN2w.ColorIndex = 3;

                    }

                    List<IList<object>> lioValuesN2w = new  List<IList<object>>();
                    for(int aa=0; aa < arr2ObjN2w.GetLength(0); aa++)
                    {
                        List<object> loValueN2w = new List<object>();
                        for(int bb=0; bb<arr2ObjN2w.GetLength(1); bb++)
                        {
                            loValueN2w.Add(arr2ObjN2w[aa, bb]);
                        }
                        lioValuesN2w.Add(loValueN2w);
                    }

                    string strShtRng = sheetAddr;
                    ValueRange vrUpdateCell = new ValueRange();
                    vrUpdateCell.MajorDimension = "ROWS";
                    vrUpdateCell.Values = lioValuesN2w;
                    SpreadsheetsResource.ValuesResource.UpdateRequest urUpdateCell = DataGlobal.sheetsService.Spreadsheets.Values.Update(vrUpdateCell, spreadSheetID, strShtRng);
                    urUpdateCell.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                    UpdateValuesResponse exeUpdateCell = urUpdateCell.Execute();

                    transctionN2w.Commit();
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                actDoc.Editor.WriteMessage("\nError H2Q: {0}", e.Message);
            }
        }


        /**
         * Date added: 07 Feb 2024
         * Added by: Engr Bernardo Cabebe Jr.
         * Venue: 3407 Cityland Pasong Tamo Tower
         */
        [CommandMethod("GSBS_GetSteelBeamSectionDetails")]
        public static void GetSteelBeamSectionDetails()
        {
            try
            {
                if (!ccData.isLicenseActive())
                {

                    actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                    return;
                }
                actDoc.Editor.WriteMessage("\nCopy google worksheet https://docs.google.com/spreadsheets/d/1ZGy2rFLEkVEQF4DunspHXn5umcu14udCRT31YycHrQs/edit#gid=0");
                actDoc.Editor.WriteMessage("\nUse color 115 for beam POLYLINE and TEXT");

                TypedValue[] tvFilterH6t = new TypedValue[2];
                tvFilterH6t[0] = new TypedValue((int)DxfCode.Start, "*POLYLINE*");
                tvFilterH6t[1] = new TypedValue((int)DxfCode.Color, 115);
                SelectionFilter sfLine = new SelectionFilter(tvFilterH6t);

                PromptSelectionOptions psoLine = new PromptSelectionOptions();
                psoLine.MessageForAdding = "\nSelect steel beam lines: ";

                PromptSelectionResult psrLine = actDoc.Editor.GetSelection(psoLine, sfLine);
                if (psrLine.Status != PromptStatus.OK) return;

                SelectionSet ssLine = psrLine.Value;
                actDoc.Editor.WriteMessage("\nSelected objects: {0}", ssLine.Count);


                PromptStringOptions psoSheetAddr = new PromptStringOptions("");
                psoSheetAddr.DefaultValue = shAddrG;
                psoSheetAddr.Message = "\nEnter cell address: ";
                psoSheetAddr.AllowSpaces = false;
                PromptResult prSheetAddr = actDoc.Editor.GetString(psoSheetAddr);

                string sheetAddr = prSheetAddr.StringResult;
                shAddrG = sheetAddr;
                actDoc.Editor.WriteMessage("\nCell address: {0}", shAddrG);

                PromptStringOptions psoGSheetID = new PromptStringOptions("");
                psoGSheetID.DefaultValue = gshtIdG;
                psoGSheetID.Message = "\nEnter Google Spreadsheet ID: ";
                psoGSheetID.AllowSpaces = false;
                PromptResult prGSheetID = actDoc.Editor.GetString(psoGSheetID);

                string spreadSheetID = prGSheetID.StringResult;
                gshtIdG = spreadSheetID;
                actDoc.Editor.WriteMessage("\nSpreadsheet ID: {0}", gshtIdG);

                string[,] arr2Value = new string[ssLine.Count,2];
                int iiRow = 0;
                using (Transaction trLine = aCurDB.TransactionManager.StartTransaction())
                {
                    bool executeP3w = true;
                    foreach (SelectedObject soLine in ssLine)
                    {
                        Polyline lineObj = (Polyline)trLine.GetObject(soLine.ObjectId, OpenMode.ForRead);

                        double lineLength = lineObj.Length;

                        Curve curvLine = (Curve)lineObj;
                        double curvLen = curvLine.GetDistanceAtParameter(curvLine.EndParam) - curvLine.GetDistanceAtParameter(curvLine.StartParam);
                        actDoc.Editor.WriteMessage("\nLength= {0}", curvLen);
                        double fenceLen = 0.5; ///for fence selection
                        if(curvLen > 500)
                        {
                            fenceLen = 800;
                            //If the beam is shorter than the fence length
                            if (0.5 * curvLen < fenceLen)
                            {
                                fenceLen = 0.45 * curvLen;
                            }
                        }
                        else
                        {
                            fenceLen = 0.8;
                            //If the beam is shorter than the fence length
                            if (0.5*curvLen < fenceLen)
                            {
                                fenceLen = 0.45 * curvLen;
                            }
                            
                        }

                        Point3d p3dMidP = curvLine.GetPointAtDist(0.5 * curvLen);
                        Point3d p3dRefP1 = curvLine.GetPointAtDist(0.5 * curvLen - fenceLen);
                        Point3d p3dRefP2 = curvLine.GetPointAtDist(0.5 * curvLen + fenceLen);


                        Point3d p3dFencP1 = RotatePoint(p3dRefP1, p3dMidP, 45);
                        Point3d p3dFencP2 = RotatePoint(p3dRefP2, p3dMidP, 45);
                        actDoc.Editor.WriteMessage("\nFence Pt1: {0}", p3dFencP1);
                        actDoc.Editor.WriteMessage("\nFence Pt2: {0}", p3dFencP2);

                        // //This line of code is use to check the location of points
                        // BlockTable blkTble = trLine.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                        // BlockTableRecord blkTRec = trLine.GetObject(blkTble[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        // 
                        // using (DBPoint acPt1 = new DBPoint(p3dFencP1))
                        // {
                        //     blkTRec.AppendEntity(acPt1);
                        //     trLine.AddNewlyCreatedDBObject(acPt1, true);
                        //     acPt1.ColorIndex = 1;
                        // }
                        // 
                        // using (DBPoint acPt2 = new DBPoint(p3dFencP2))
                        // {
                        //     blkTRec.AppendEntity(acPt2);
                        //     trLine.AddNewlyCreatedDBObject(acPt2, true);
                        //     acPt2.ColorIndex = 2;
                        // }

                        TypedValue[] tvFilterTxt = new TypedValue[2];
                        tvFilterH6t[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                        tvFilterH6t[1] = new TypedValue((int)DxfCode.Color, 115);
                        SelectionFilter sfTxt = new SelectionFilter(tvFilterH6t);

                        Point3dCollection p3dCTxt = new Point3dCollection()
                        {
                            p3dFencP1,
                            p3dFencP2
                        };

                        PromptSelectionResult psrTxt = actDoc.Editor.SelectFence(p3dCTxt, sfTxt);
                        if (psrTxt.Status != PromptStatus.OK)
                        {
                            actDoc.Editor.WriteMessage("\nText selection exit.");
                            curvLine.Highlight();
                            executeP3w = false;
                        }
                        else
                        {
                            actDoc.Editor.WriteMessage("\nText selected.");
                        }

                        if (executeP3w)
                        {
                            SelectionSet ssTxt = psrTxt.Value;
                            actDoc.Editor.WriteMessage("\nText: {0}", ssTxt.Count);

                            DBText txtObj = (DBText)trLine.GetObject(ssTxt[0].ObjectId, OpenMode.ForRead);
                            string txCont = txtObj.TextString;

                            arr2Value[iiRow, 0] = txCont;
                            arr2Value[iiRow, 1] = curvLen.ToString();
                            actDoc.Editor.WriteMessage("\nSection: {0}", txCont);

                            iiRow++;

                            //Change the colors of lines and text
                            lineObj.UpgradeOpen();
                            lineObj.ColorIndex = 2;

                            txtObj.UpgradeOpen();
                            txtObj.ColorIndex = 3;
                        }
                    }

                    if (executeP3w)
                    {
                        List<IList<object>> lioValues = new List<IList<object>>();
                        for (int ii = 0; ii < arr2Value.GetLength(0); ii++)
                        {
                            List<object> loValue = new List<object>();
                            for (int jj = 0; jj < arr2Value.GetLength(1); jj++)
                            {
                                loValue.Add(arr2Value[ii, jj]);
                            }
                            lioValues.Add(loValue);
                        }

                        string strShtRng = sheetAddr;
                        ValueRange vrUpdateCell = new ValueRange();
                        vrUpdateCell.MajorDimension = "ROWS";
                        vrUpdateCell.Values = lioValues;
                        SpreadsheetsResource.ValuesResource.UpdateRequest urUpdateCell = DataGlobal.sheetsService.Spreadsheets.Values.Update(vrUpdateCell, spreadSheetID, strShtRng);
                        urUpdateCell.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                        UpdateValuesResponse exeUpdateCell = urUpdateCell.Execute();
                    }

                    trLine.Commit();
                }

            }
            catch(Autodesk.AutoCAD.Runtime.Exception e)
            {
                actDoc.Editor.WriteMessage("\nError: {0}",e.Message);
            }
        }


        /**
         * Date added: 20 Jun 2025
         * Added by: Engr Bernardo Cabebe Jr.
         * Venue: 3rd Floor CPTT
         * Rationale: This command is brought up by the estimate of 
         *            steel sections of a warehouse project
         */
        [CommandMethod("GSSD_GetSteelSectionDetails")]
        public static void GetSteelSectionDetails()
        {
            try
            {
                if (!ccData.isLicenseActive())
                {

                    actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                    return;
                }
                actDoc.Editor.WriteMessage("\nCopy google worksheet https://docs.google.com/spreadsheets/d/1ZGy2rFLEkVEQF4DunspHXn5umcu14udCRT31YycHrQs/edit#gid=0");
                actDoc.Editor.WriteMessage("\nUse color 204 for steel POLYLINE and TEXT");

                TypedValue[] tvFilterH6t = new TypedValue[2];
                tvFilterH6t[0] = new TypedValue((int)DxfCode.Start, "*POLYLINE*");
                tvFilterH6t[1] = new TypedValue((int)DxfCode.Color, 204);
                SelectionFilter sfLine = new SelectionFilter(tvFilterH6t);

                PromptSelectionOptions psoLine = new PromptSelectionOptions();
                psoLine.MessageForAdding = "\nSelect steel section lines: ";

                PromptSelectionResult psrLine = actDoc.Editor.GetSelection(psoLine, sfLine);
                if (psrLine.Status != PromptStatus.OK) return;

                SelectionSet ssLine = psrLine.Value;
                actDoc.Editor.WriteMessage("\nSelected objects: {0}", ssLine.Count);


                PromptStringOptions psoSheetAddr = new PromptStringOptions("");
                psoSheetAddr.DefaultValue = shAddrG;
                psoSheetAddr.Message = "\nEnter cell address: ";
                psoSheetAddr.AllowSpaces = false;
                PromptResult prSheetAddr = actDoc.Editor.GetString(psoSheetAddr);

                string sheetAddr = prSheetAddr.StringResult;
                shAddrG = sheetAddr;
                actDoc.Editor.WriteMessage("\nCell address: {0}", shAddrG);

                PromptStringOptions psoGSheetID = new PromptStringOptions("");
                psoGSheetID.DefaultValue = gshtIdG;
                psoGSheetID.Message = "\nEnter Google Spreadsheet ID: ";
                psoGSheetID.AllowSpaces = false;
                PromptResult prGSheetID = actDoc.Editor.GetString(psoGSheetID);

                string spreadSheetID = prGSheetID.StringResult;
                gshtIdG = spreadSheetID;
                actDoc.Editor.WriteMessage("\nSpreadsheet ID: {0}", gshtIdG);

                string[,] arr2Value = new string[ssLine.Count, 2];
                int iiRow = 0;
                using (Transaction trLine = aCurDB.TransactionManager.StartTransaction())
                {
                    bool executeP3w = true;
                    foreach (SelectedObject soLine in ssLine)
                    {
                        Polyline lineObj = (Polyline)trLine.GetObject(soLine.ObjectId, OpenMode.ForRead);

                        double lineLength = lineObj.Length;

                        Curve curvLine = (Curve)lineObj;
                        double curvLen = curvLine.GetDistanceAtParameter(curvLine.EndParam) - curvLine.GetDistanceAtParameter(curvLine.StartParam);
                        actDoc.Editor.WriteMessage("\nLength= {0}", curvLen);
                        double fenceLen = 0.5; ///for fence selection
                        if (curvLen > 500)
                        {
                            fenceLen = 800;
                            //If the beam is shorter than the fence length
                            if (0.5 * curvLen < fenceLen)
                            {
                                fenceLen = 0.45 * curvLen;
                            }
                        }
                        else
                        {
                            fenceLen = 0.8;
                            //If the beam is shorter than the fence length
                            if (0.5 * curvLen < fenceLen)
                            {
                                fenceLen = 0.45 * curvLen;
                            }

                        }

                        Point3d p3dMidP = curvLine.GetPointAtDist(0.5 * curvLen);
                        Point3d p3dRefP1 = curvLine.GetPointAtDist(0.5 * curvLen - fenceLen);
                        Point3d p3dRefP2 = curvLine.GetPointAtDist(0.5 * curvLen + fenceLen);


                        Point3d p3dFencP1 = RotatePoint(p3dRefP1, p3dMidP, 45);
                        Point3d p3dFencP2 = RotatePoint(p3dRefP2, p3dMidP, 45);
                        actDoc.Editor.WriteMessage("\nFence Pt1: {0}", p3dFencP1);
                        actDoc.Editor.WriteMessage("\nFence Pt2: {0}", p3dFencP2);

                        // //This line of code is use to check the location of points
                        // BlockTable blkTble = trLine.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                        // BlockTableRecord blkTRec = trLine.GetObject(blkTble[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        // 
                        // using (DBPoint acPt1 = new DBPoint(p3dFencP1))
                        // {
                        //     blkTRec.AppendEntity(acPt1);
                        //     trLine.AddNewlyCreatedDBObject(acPt1, true);
                        //     acPt1.ColorIndex = 1;
                        // }
                        // 
                        // using (DBPoint acPt2 = new DBPoint(p3dFencP2))
                        // {
                        //     blkTRec.AppendEntity(acPt2);
                        //     trLine.AddNewlyCreatedDBObject(acPt2, true);
                        //     acPt2.ColorIndex = 2;
                        // }

                        TypedValue[] tvFilterTxt = new TypedValue[2];
                        tvFilterH6t[0] = new TypedValue((int)DxfCode.Start, "TEXT");
                        tvFilterH6t[1] = new TypedValue((int)DxfCode.Color, 204);
                        SelectionFilter sfTxt = new SelectionFilter(tvFilterH6t);

                        Point3dCollection p3dCTxt = new Point3dCollection()
                        {
                            p3dFencP1,
                            p3dFencP2
                        };

                        PromptSelectionResult psrTxt = actDoc.Editor.SelectFence(p3dCTxt, sfTxt);
                        if (psrTxt.Status != PromptStatus.OK)
                        {
                            actDoc.Editor.WriteMessage("\nText selection exit.");
                            curvLine.Highlight();
                            executeP3w = false;
                        }
                        else
                        {
                            actDoc.Editor.WriteMessage("\nText selected.");
                        }

                        if (executeP3w)
                        {
                            SelectionSet ssTxt = psrTxt.Value;
                            actDoc.Editor.WriteMessage("\nText: {0}", ssTxt.Count);

                            DBText txtObj = (DBText)trLine.GetObject(ssTxt[0].ObjectId, OpenMode.ForRead);
                            string txCont = txtObj.TextString;

                            arr2Value[iiRow, 0] = txCont;
                            arr2Value[iiRow, 1] = curvLen.ToString();
                            actDoc.Editor.WriteMessage("\nSection: {0}", txCont);

                            iiRow++;

                            //Change the colors of lines and text
                            lineObj.UpgradeOpen();
                            lineObj.ColorIndex = 64;

                            txtObj.UpgradeOpen();
                            txtObj.ColorIndex = 64;
                        }
                    }

                    if (executeP3w)
                    {
                        List<IList<object>> lioValues = new List<IList<object>>();
                        for (int ii = 0; ii < arr2Value.GetLength(0); ii++)
                        {
                            List<object> loValue = new List<object>();
                            for (int jj = 0; jj < arr2Value.GetLength(1); jj++)
                            {
                                loValue.Add(arr2Value[ii, jj]);
                            }
                            lioValues.Add(loValue);
                        }

                        string strShtRng = sheetAddr;
                        ValueRange vrUpdateCell = new ValueRange();
                        vrUpdateCell.MajorDimension = "ROWS";
                        vrUpdateCell.Values = lioValues;
                        SpreadsheetsResource.ValuesResource.UpdateRequest urUpdateCell = DataGlobal.sheetsService.Spreadsheets.Values.Update(vrUpdateCell, spreadSheetID, strShtRng);
                        urUpdateCell.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                        UpdateValuesResponse exeUpdateCell = urUpdateCell.Execute();
                    }

                    trLine.Commit();
                }

            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                actDoc.Editor.WriteMessage("\nError: {0}", e.Message);
            }
        }


        private static Point3d RotatePoint(Point3d pointToRotate, Point3d centerPoint, double angleInDegrees)
        {
            double angRad = angleInDegrees * (Math.PI / 180);
            double cosTheta = Math.Cos(angRad);
            double sinTheta = Math.Sin(angRad);

            double X =
                    (double)
                    (cosTheta * (pointToRotate.X - centerPoint.X) -
                    sinTheta * (pointToRotate.Y - centerPoint.Y) + centerPoint.X);
            double Y =
                    (double)
                    (sinTheta * (pointToRotate.X - centerPoint.X) +
                    cosTheta * (pointToRotate.Y - centerPoint.Y) + centerPoint.Y);
            

            Point3d rotPt = new Point3d(X, Y, pointToRotate.Z);
            return rotPt;
        }
    }
}
