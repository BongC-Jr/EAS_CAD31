using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace EASI_CAD31
{
    public class ColumnSchedule
    {
        static Document actDoc = Application.DocumentManager.MdiActiveDocument;
        static Database aCurDB = actDoc.Database;

        static CCData ccData = new CCData();

        //Column schedule
        //https://docs.google.com/spreadsheets/d/1dJbHXn0HQHa_N_eQE81fbP3nUzA7S3qguddc0JwlcQU/edit?gid=1602764100#gid=1602764100
        static string csGSheetIDG = "1dJbHXn0HQHa_N_eQE81fbP3nUzA7S3qguddc0JwlcQU";
        static int rowG = 7;

        [CommandMethod("DCS_DrawColumnSection")]
        public static void DrawColumnSection()
        {
            actDoc.Editor.WriteMessage("\nMake a copy of the google worksheet https://docs.google.com/spreadsheets/d/1dJbHXn0HQHa_N_eQE81fbP3nUzA7S3qguddc0JwlcQU/edit?gid=1602764100#gid=1602764100");
            PromptPointOptions ppo = new PromptPointOptions("");
            ppo.Message = "\nEnter insertion point: ";

            PromptPointResult ppr = actDoc.Editor.GetPoint(ppo);
            Point3d insertionPt = ppr.Value;

            if (ppr.Status != PromptStatus.OK)
            {
                actDoc.Editor.WriteMessage("\nQuit...");
                return;
            }

            PromptIntegerOptions pioRowp24t = new PromptIntegerOptions("");
            pioRowp24t.Message = "\nEnter row number: ";
            pioRowp24t.DefaultValue = rowG;

            PromptIntegerResult pirRowp24t = actDoc.Editor.GetInteger(pioRowp24t);
            if (pirRowp24t.Status != PromptStatus.OK) return;
            int shRowp24t = pirRowp24t.Value;
            rowG = shRowp24t;

            actDoc.Editor.WriteMessage("\nCopy google sheet Column Schedule template from: https://docs.google.com/spreadsheets/d/1dJbHXn0HQHa_N_eQE81fbP3nUzA7S3qguddc0JwlcQU/edit?gid=1602764100#gid=1602764100");
            PromptStringOptions psoCSp24t = new PromptStringOptions("");
            psoCSp24t.Message = "\nEnter Column Schedule Google Sheet ID:";
            psoCSp24t.DefaultValue = csGSheetIDG;

            PromptResult prCSp24t = actDoc.Editor.GetString(psoCSp24t);
            if (prCSp24t.Status != PromptStatus.OK) return;
            string csSheetID = prCSp24t.StringResult;
            csGSheetIDG = csSheetID;

            //Get the data from the google sheet Column Schedule
            string strCellAddr = "ColSched!G" + shRowp24t.ToString() + ":T" + shRowp24t.ToString();
            SpreadsheetsResource.ValuesResource.GetRequest grColDet = DataGlobal.sheetsService.Spreadsheets.Values.Get(csSheetID, strCellAddr);
            ValueRange vrColDet = grColDet.Execute();
            IList<IList<object>> iioColDet = vrColDet.Values;

            double colWid = Convert.ToDouble(iioColDet[0][0]);
            double colHgt = Convert.ToDouble(iioColDet[0][1]);

            Point2d pvtx1 = new Point2d(insertionPt.X, insertionPt.Y);
            Point2d pvtx2 = new Point2d(pvtx1.X + colWid, pvtx1.Y);
            Point2d pvtx3 = new Point2d(pvtx2.X, pvtx2.Y - colHgt);
            Point2d pvtx4 = new Point2d(pvtx3.X - colWid, pvtx3.Y);
            Point2d pvtx5 = new Point2d(insertionPt.X, insertionPt.Y);

            //Create column outline
            using (Transaction trColSec = aCurDB.TransactionManager.StartTransaction())
            {
                BlockTable blktbl;
                blktbl = trColSec.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord bltrec;
                bltrec = trColSec.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                Polyline pColSec = new Polyline();
                pColSec.SetDatabaseDefaults();
                pColSec.AddVertexAt(0, pvtx1, 0, 0, 0);
                pColSec.AddVertexAt(1, pvtx2, 0, 0, 0);
                pColSec.AddVertexAt(2, pvtx3, 0, 0, 0);
                pColSec.AddVertexAt(3, pvtx4, 0, 0, 0);
                pColSec.AddVertexAt(4, pvtx5, 0, 0, 0);

                bltrec.AppendEntity(pColSec);
                trColSec.AddNewlyCreatedDBObject(pColSec, true);

                trColSec.Commit();
            }


            double verBarDia = Convert.ToDouble(iioColDet[0][6]);
            double tieBarDia = Convert.ToDouble(iioColDet[0][7]);
            double concCover = 40.0;

            //Circle center points
            Point3d cirCPt1 = new Point3d(pvtx1.X + concCover + tieBarDia + verBarDia * 0.5, pvtx1.Y - concCover - tieBarDia - verBarDia * 0.5, 0.0);
            Point3d cirCPt2 = new Point3d(pvtx2.X - concCover - tieBarDia - verBarDia * 0.5, pvtx2.Y - concCover - tieBarDia - verBarDia * 0.5, 0.0);
            Point3d cirCPt3 = new Point3d(pvtx3.X - concCover - tieBarDia - verBarDia * 0.5, pvtx3.Y + concCover + tieBarDia + verBarDia * 0.5, 0.0);
            Point3d cirCPt4 = new Point3d(pvtx4.X + concCover + tieBarDia + verBarDia * 0.5, pvtx3.Y + concCover + tieBarDia + verBarDia * 0.5, 0.0);
            Point3d[] cirPts = { cirCPt1, cirCPt2, cirCPt3, cirCPt4 };
            using (Transaction trVertBar = aCurDB.TransactionManager.StartTransaction())
            {
                BlockTable blktbl;
                blktbl = trVertBar.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord bltrec;
                bltrec = trVertBar.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                for (int i = 0; i < cirPts.Length; i++)
                {
                    Circle cVerBar = new Circle();
                    cVerBar.SetDatabaseDefaults();
                    cVerBar.Center = cirPts[i];
                    cVerBar.Diameter = verBarDia;

                    bltrec.AppendEntity(cVerBar);
                    trVertBar.AddNewlyCreatedDBObject(cVerBar, true);
                }

                trVertBar.Commit();
            }


            Vector3d s1Vector = cirCPt1.GetVectorTo(cirCPt2);
            int bBarNum = Convert.ToInt16(iioColDet[0][3]);
            int bAddBar = bBarNum - 1;
            using (Transaction trAddVerBar = aCurDB.TransactionManager.StartTransaction())
            {
                BlockTable blktbl;
                blktbl = trAddVerBar.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord bltrec;
                bltrec = trAddVerBar.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                Point3d jCirPt = cirCPt1;
                for (int j = 0; j < bAddBar - 1; j++)
                {
                    jCirPt = jCirPt + s1Vector / bAddBar;

                    Circle cAddBar = new Circle();
                    cAddBar.SetDatabaseDefaults();
                    cAddBar.Center = jCirPt;
                    cAddBar.Diameter = verBarDia;

                    bltrec.AppendEntity(cAddBar);
                    trAddVerBar.AddNewlyCreatedDBObject(cAddBar, true);
                }

                trAddVerBar.Commit();
            }

            Vector3d s2Vector = cirCPt3.GetVectorTo(cirCPt4);
            using (Transaction trAddVerBar = aCurDB.TransactionManager.StartTransaction())
            {
                BlockTable blktbl;
                blktbl = trAddVerBar.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord bltrec;
                bltrec = trAddVerBar.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                Point3d kCirPt = cirCPt3;
                for (int k = 0; k < bAddBar - 1; k++)
                {
                    kCirPt = kCirPt + s2Vector / bAddBar;

                    Circle cAddBar = new Circle();
                    cAddBar.SetDatabaseDefaults();
                    cAddBar.Center = kCirPt;
                    cAddBar.Diameter = verBarDia;

                    bltrec.AppendEntity(cAddBar);
                    trAddVerBar.AddNewlyCreatedDBObject(cAddBar, true);
                }

                trAddVerBar.Commit();
            }


            Vector3d s3Vector = cirCPt4.GetVectorTo(cirCPt1);
            Vector3d s4Vector = cirCPt3.GetVectorTo(cirCPt2);
            int hBarNum = Convert.ToInt16(iioColDet[0][4]);
            int hAddBar = hBarNum - 1;
            using (Transaction trAddVerBar = aCurDB.TransactionManager.StartTransaction())
            {
                BlockTable blktbl;
                blktbl = trAddVerBar.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord bltrec;
                bltrec = trAddVerBar.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                Point3d aa1CirPt = cirCPt4;
                Point3d aa2CirPt = cirCPt3;
                for (int aa = 0; aa < hAddBar - 1; aa++)
                {
                    aa1CirPt = aa1CirPt + s3Vector / hAddBar;

                    Circle cAddBarInH1 = new Circle();
                    cAddBarInH1.SetDatabaseDefaults();
                    cAddBarInH1.Center = aa1CirPt;
                    cAddBarInH1.Diameter = verBarDia;

                    bltrec.AppendEntity(cAddBarInH1);
                    trAddVerBar.AddNewlyCreatedDBObject(cAddBarInH1, true);


                    aa2CirPt = aa2CirPt + s4Vector / hAddBar;

                    Circle cAddBarInH2 = new Circle();
                    cAddBarInH2.SetDatabaseDefaults();
                    cAddBarInH2.Center = aa2CirPt;
                    cAddBarInH2.Diameter = verBarDia;

                    bltrec.AppendEntity(cAddBarInH2);
                    trAddVerBar.AddNewlyCreatedDBObject(cAddBarInH2, true);
                }

                trAddVerBar.Commit();
            }

            actDoc.Editor.WriteMessage("\nColumn section created");
        }

        [CommandMethod("DCT_DrawCloseTie")]
        public static void DrawCloseTie()
        {
            TypedValue[] tvFilterB3r = new TypedValue[1];
            tvFilterB3r[0] = new TypedValue((int)DxfCode.Start, "CIRCLE");
            SelectionFilter sfVerBar = new SelectionFilter(tvFilterB3r);

            PromptSelectionOptions psoVerBar = new PromptSelectionOptions();
            psoVerBar.MessageForAdding = "\nSelect vertical bars: ";

            PromptIntegerOptions pioRowp24t = new PromptIntegerOptions("");
            pioRowp24t.Message = "\nEnter row number: ";
            pioRowp24t.DefaultValue = rowG;

            PromptIntegerResult pirRowp24t = actDoc.Editor.GetInteger(pioRowp24t);
            if (pirRowp24t.Status != PromptStatus.OK) return;
            int shRowp24t = pirRowp24t.Value;
            rowG = shRowp24t;

            actDoc.Editor.WriteMessage("\nCopy google sheet Column Schedule template from: https://docs.google.com/spreadsheets/d/1dJbHXn0HQHa_N_eQE81fbP3nUzA7S3qguddc0JwlcQU/edit?gid=1602764100#gid=1602764100");
            PromptStringOptions psoCSp24t = new PromptStringOptions("");
            psoCSp24t.Message = "\nEnter Column Schedule Google Sheet ID:";
            psoCSp24t.DefaultValue = csGSheetIDG;

            PromptResult prCSp24t = actDoc.Editor.GetString(psoCSp24t);
            if (prCSp24t.Status != PromptStatus.OK) return;
            string csSheetID = prCSp24t.StringResult;
            csGSheetIDG = csSheetID;

            //Get the data from the google sheet Column Schedule
            string strCellAddr = "ColSched!G" + shRowp24t.ToString() + ":T" + shRowp24t.ToString();
            SpreadsheetsResource.ValuesResource.GetRequest grColDet = DataGlobal.sheetsService.Spreadsheets.Values.Get(csSheetID, strCellAddr);
            ValueRange vrColDet = grColDet.Execute();
            IList<IList<object>> iioColDet = vrColDet.Values;

            double colWid = Convert.ToDouble(iioColDet[0][0]);
            double colHgt = Convert.ToDouble(iioColDet[0][1]);
            double verBarDia = Convert.ToDouble(iioColDet[0][6]);
            double tieBarDia = Convert.ToDouble(iioColDet[0][7]);

            PromptSelectionResult psrVerBar = actDoc.Editor.GetSelection(psoVerBar, sfVerBar);
            if (psrVerBar.Status != PromptStatus.OK)
            {
                actDoc.Editor.WriteMessage("\nQuit...");
                return;
            }

            SelectionSet ssVerBar = psrVerBar.Value;

            if(ssVerBar.Count != 2)
            {
                actDoc.Editor.WriteMessage("\nSelect 2 vertical bars only. Quit...");
                return;
            }

            using (Transaction trCloseTie = aCurDB.TransactionManager.StartTransaction())
            {
                BlockTable blktbl;
                blktbl = trCloseTie.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord bltrec;
                bltrec = trCloseTie.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                Circle cVerBar1 = trCloseTie.GetObject(ssVerBar[0].ObjectId, OpenMode.ForRead) as Circle;
                Circle cVerBar2 = trCloseTie.GetObject(ssVerBar[1].ObjectId, OpenMode.ForRead) as Circle;
                Point3d cirCPt1 = cVerBar1.Center;
                Point3d cirCPt2 = cVerBar2.Center;

                Matrix3d curUCSMatrix = actDoc.Editor.CurrentUserCoordinateSystem;
                CoordinateSystem3d curUCS = curUCSMatrix.CoordinateSystem3d;

                Point3d plpt02u = new Point3d(cirCPt1.X + verBarDia * 0.5 + tieBarDia, cirCPt1.Y, cirCPt1.Z);
                Point3d plpt02 = plpt02u.TransformBy(Matrix3d.Rotation(0.785398, curUCS.Zaxis, cirCPt1));

                Point3d plpt01u = new Point3d(plpt02.X + 6 * tieBarDia, plpt02.Y, plpt02.Z);
                Point3d plpt01 = plpt01u.TransformBy(Matrix3d.Rotation(-0.785398, curUCS.Zaxis, plpt02));

                Point3d appPt1 = new Point3d(cirCPt1.X, cirCPt2.Y, cirCPt1.Z);
                Point3d appPt2 = new Point3d(cirCPt2.X, cirCPt1.Y, cirCPt1.Z);
                Point3d plpt04 = new Point3d(appPt1.X - verBarDia * 0.5 - tieBarDia, appPt1.Y , appPt1.Z);
                Point3d plpt05 = new Point3d(appPt1.X, appPt1.Y - verBarDia * 0.5 - tieBarDia, appPt1.Z);
                Point3d plpt06 = new Point3d(cirCPt2.X, cirCPt2.Y - verBarDia * 0.5 - tieBarDia, cirCPt2.Z);
                Point3d plpt07 = new Point3d(cirCPt2.X + verBarDia * 0.5 + tieBarDia, cirCPt2.Y, cirCPt2.Z);
                Point3d plpt08 = new Point3d(appPt2.X + verBarDia * 0.5 + tieBarDia, appPt2.Y, appPt2.Z);
                Point3d plpt09 = new Point3d(appPt2.X, appPt2.Y + verBarDia * 0.5 + tieBarDia, appPt2.Z);
                Point3d plpt10 = new Point3d(cirCPt1.X, cirCPt1.Y + verBarDia * 0.5 + tieBarDia, cirCPt1.Z);
                Point3d plpt11u = new Point3d(cirCPt1.X - verBarDia * 0.5 - tieBarDia, cirCPt1.Y, cirCPt1.Z);
                Point3d plpt11 = plpt11u.TransformBy(Matrix3d.Rotation(0.785398, curUCS.Zaxis, cirCPt1));
                Point3d plpt12u = new Point3d(plpt11.X, plpt11.Y - 6 * tieBarDia, plpt11.Z);
                Point3d plpt12 = plpt12u.TransformBy(Matrix3d.Rotation(0.785398, curUCS.Zaxis, plpt11));

                Polyline plSeg1;
                using (plSeg1 = new Polyline())
                {
                    plSeg1.SetDatabaseDefaults();
                    plSeg1.AddVertexAt(0, new Point2d(plpt01.X, plpt01.Y), 0, 0, 0);

                    double arcRad = verBarDia * 0.5 + tieBarDia;
                    Point3d plpt03 = new Point3d(cirCPt1.X - verBarDia * 0.5 - tieBarDia, cirCPt1.Y, cirCPt1.Z);
                    plSeg1.AddVertexAt(1, new Point2d(plpt02.X, plpt02.Y), 0.668179, 0, 0);
                    plSeg1.AddVertexAt(2, new Point2d(plpt03.X, plpt03.Y), 0, 0, 0);

                    plSeg1.AddVertexAt(3, new Point2d(plpt04.X, plpt04.Y), 0.414214, 0, 0);
                    plSeg1.AddVertexAt(4, new Point2d(plpt05.X, plpt05.Y), 0, 0, 0);

                    plSeg1.AddVertexAt(5, new Point2d(plpt06.X, plpt06.Y), 0.414214, 0, 0);
                    plSeg1.AddVertexAt(6, new Point2d(plpt07.X, plpt07.Y), 0, 0, 0);

                    plSeg1.AddVertexAt(7, new Point2d(plpt08.X, plpt08.Y), 0.414214, 0, 0);
                    plSeg1.AddVertexAt(8, new Point2d(plpt09.X, plpt09.Y), 0, 0, 0);

                    plSeg1.AddVertexAt(9, new Point2d(plpt10.X, plpt10.Y), 0.668179, 0, 0);
                    plSeg1.AddVertexAt(10, new Point2d(plpt11.X, plpt11.Y), 0, 0, 0);

                    plSeg1.AddVertexAt(11, new Point2d(plpt12.X, plpt12.Y), 0, 0, 0);

                    bltrec.AppendEntity(plSeg1);
                    trCloseTie.AddNewlyCreatedDBObject(plSeg1, true);
                }

                trCloseTie.Commit();
            }

        }


        [CommandMethod("CSTV_ColumnSetTextVerticalBar")]
        public static void ColumnSetTextVerticalBar()
        {
            TypedValue[] tvFilterH1q = new TypedValue[4];
            tvFilterH1q[0] = new TypedValue((int)DxfCode.Operator, "<OR");
            tvFilterH1q[1] = new TypedValue((int)DxfCode.Start, "MTEXT");
            tvFilterH1q[2] = new TypedValue((int)DxfCode.Start, "TEXT");
            tvFilterH1q[3] = new TypedValue((int)DxfCode.Operator, "OR>");

            SelectionFilter sfColText = new SelectionFilter(tvFilterH1q);
            PromptSelectionOptions psoColText = new PromptSelectionOptions();
            psoColText.MessageForAdding = "\nSelect text vertical bar: ";
            
            PromptSelectionResult psrColText = actDoc.Editor.GetSelection(psoColText, sfColText);
            if (psrColText.Status != PromptStatus.OK)
            {
                actDoc.Editor.WriteMessage("\nQuit...");
                return;
            }
            
            SelectionSet ssColText = psrColText.Value;
            if (ssColText.Count != 1)
            {
                actDoc.Editor.WriteMessage("\nSelect only one text vertical bar. Quit...");
                return;
            }


            PromptIntegerOptions pioRowp24t = new PromptIntegerOptions("");
            pioRowp24t.Message = "\nEnter row number: ";
            pioRowp24t.DefaultValue = rowG;

            PromptIntegerResult pirRowp24t = actDoc.Editor.GetInteger(pioRowp24t);
            if (pirRowp24t.Status != PromptStatus.OK) return;
            int shRowp24t = pirRowp24t.Value;
            rowG = shRowp24t;

            PromptStringOptions psoCSp24t = new PromptStringOptions("");
            psoCSp24t.Message = "\nEnter Column Schedule Google Sheet ID:";
            psoCSp24t.DefaultValue = csGSheetIDG;

            PromptResult prCSp24t = actDoc.Editor.GetString(psoCSp24t);
            if (prCSp24t.Status != PromptStatus.OK) return;
            string csSheetID = prCSp24t.StringResult;
            csGSheetIDG = csSheetID;

            //Get the data from the google sheet Column Schedule
            string strCellAddr = "ColSched!G" + shRowp24t.ToString() + ":T" + shRowp24t.ToString();
            SpreadsheetsResource.ValuesResource.GetRequest grColDet = DataGlobal.sheetsService.Spreadsheets.Values.Get(csSheetID, strCellAddr);
            ValueRange vrColDet = grColDet.Execute();
            IList<IList<object>> iioColDet = vrColDet.Values;

            string barDia = Convert.ToString(iioColDet[0][6]);
            string barNum = Convert.ToString(iioColDet[0][5]);
            string colVB = barNum + " - %%C" + barDia + "mm";


            using (Transaction trColText = aCurDB.TransactionManager.StartTransaction())
            {
                BlockTable blktbl;
                blktbl = trColText.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord bltrec;
                bltrec = trColText.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                Entity textObj = trColText.GetObject(ssColText[0].ObjectId, OpenMode.ForWrite) as Entity;
                if (textObj is DBText)
                {
                    DBText dbText = (DBText)textObj;
                    dbText.TextString = colVB;
                }
                else if (textObj is MText)
                {
                    MText mText = (MText)textObj;
                    mText.Contents = colVB;
                }
                else
                {
                    actDoc.Editor.WriteMessage("\nSelected object is not a text. Quit...");
                    return;
                }

                trColText.Commit();
            }

        }



        [CommandMethod("CSTJ_ColumnSetTextJointTies")]
        public static void ColumnSetTextJointTies()
        {
            TypedValue[] tvFilterH1q = new TypedValue[4];
            tvFilterH1q[0] = new TypedValue((int)DxfCode.Operator, "<OR");
            tvFilterH1q[1] = new TypedValue((int)DxfCode.Start, "MTEXT");
            tvFilterH1q[2] = new TypedValue((int)DxfCode.Start, "TEXT");
            tvFilterH1q[3] = new TypedValue((int)DxfCode.Operator, "OR>");

            SelectionFilter sfColText = new SelectionFilter(tvFilterH1q);
            PromptSelectionOptions psoColText = new PromptSelectionOptions();
            psoColText.MessageForAdding = "\nSelect text confinement ties: ";

            PromptSelectionResult psrColText = actDoc.Editor.GetSelection(psoColText, sfColText);
            if (psrColText.Status != PromptStatus.OK)
            {
                actDoc.Editor.WriteMessage("\nQuit...");
                return;
            }

            SelectionSet ssColText = psrColText.Value;
            if (ssColText.Count != 1)
            {
                actDoc.Editor.WriteMessage("\nSelect only one text vertical bar. Quit...");
                return;
            }


            PromptIntegerOptions pioRowp24t = new PromptIntegerOptions("");
            pioRowp24t.Message = "\nEnter row number: ";
            pioRowp24t.DefaultValue = rowG;

            PromptIntegerResult pirRowp24t = actDoc.Editor.GetInteger(pioRowp24t);
            if (pirRowp24t.Status != PromptStatus.OK) return;
            int shRowp24t = pirRowp24t.Value;
            rowG = shRowp24t;

            PromptStringOptions psoCSp24t = new PromptStringOptions("");
            psoCSp24t.Message = "\nEnter Column Schedule Google Sheet ID:";
            psoCSp24t.DefaultValue = csGSheetIDG;

            PromptResult prCSp24t = actDoc.Editor.GetString(psoCSp24t);
            if (prCSp24t.Status != PromptStatus.OK) return;
            string csSheetID = prCSp24t.StringResult;
            csGSheetIDG = csSheetID;

            //Get the data from the google sheet Column Schedule
            string strCellAddr = "ColSched!G" + shRowp24t.ToString() + ":T" + shRowp24t.ToString();
            SpreadsheetsResource.ValuesResource.GetRequest grColDet = DataGlobal.sheetsService.Spreadsheets.Values.Get(csSheetID, strCellAddr);
            ValueRange vrColDet = grColDet.Execute();
            IList<IList<object>> iioColDet = vrColDet.Values;

            string tieDia = Convert.ToString(iioColDet[0][7]);
            string jointS = Convert.ToString(iioColDet[0][11]);
            string colJT = tieDia + "%%C @" + jointS + " O.C.";


            using (Transaction trColText = aCurDB.TransactionManager.StartTransaction())
            {
                BlockTable blktbl;
                blktbl = trColText.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord bltrec;
                bltrec = trColText.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                Entity textObj = trColText.GetObject(ssColText[0].ObjectId, OpenMode.ForWrite) as Entity;
                if (textObj is DBText)
                {
                    DBText dbText = (DBText)textObj;
                    dbText.TextString = colJT;
                }
                else if (textObj is MText)
                {
                    MText mText = (MText)textObj;
                    mText.Contents = colJT;
                }
                else
                {
                    actDoc.Editor.WriteMessage("\nSelected object is not a text. Quit...");
                    return;
                }

                trColText.Commit();
            }
        }


        [CommandMethod("CSTC_ColumnSetTextColTies")]
        public static void ColumnSetTextColTies()
        {
            TypedValue[] tvFilterH1q = new TypedValue[4];
            tvFilterH1q[0] = new TypedValue((int)DxfCode.Operator, "<OR");
            tvFilterH1q[1] = new TypedValue((int)DxfCode.Start, "MTEXT");
            tvFilterH1q[2] = new TypedValue((int)DxfCode.Start, "TEXT");
            tvFilterH1q[3] = new TypedValue((int)DxfCode.Operator, "OR>");

            SelectionFilter sfColText = new SelectionFilter(tvFilterH1q);
            PromptSelectionOptions psoColText = new PromptSelectionOptions();
            psoColText.MessageForAdding = "\nSelect text confinement ties: ";

            PromptSelectionResult psrColText = actDoc.Editor.GetSelection(psoColText, sfColText);
            if (psrColText.Status != PromptStatus.OK)
            {
                actDoc.Editor.WriteMessage("\nQuit...");
                return;
            }

            SelectionSet ssColText = psrColText.Value;
            if (ssColText.Count != 1)
            {
                actDoc.Editor.WriteMessage("\nSelect only one text vertical bar. Quit...");
                return;
            }


            PromptIntegerOptions pioRowp24t = new PromptIntegerOptions("");
            pioRowp24t.Message = "\nEnter row number: ";
            pioRowp24t.DefaultValue = rowG;

            PromptIntegerResult pirRowp24t = actDoc.Editor.GetInteger(pioRowp24t);
            if (pirRowp24t.Status != PromptStatus.OK) return;
            int shRowp24t = pirRowp24t.Value;
            rowG = shRowp24t;

            PromptStringOptions psoCSp24t = new PromptStringOptions("");
            psoCSp24t.Message = "\nEnter Column Schedule Google Sheet ID:";
            psoCSp24t.DefaultValue = csGSheetIDG;

            PromptResult prCSp24t = actDoc.Editor.GetString(psoCSp24t);
            if (prCSp24t.Status != PromptStatus.OK) return;
            string csSheetID = prCSp24t.StringResult;
            csGSheetIDG = csSheetID;

            //Get the data from the google sheet Column Schedule
            string strCellAddr = "ColSched!G" + shRowp24t.ToString() + ":T" + shRowp24t.ToString();
            SpreadsheetsResource.ValuesResource.GetRequest grColDet = DataGlobal.sheetsService.Spreadsheets.Values.Get(csSheetID, strCellAddr);
            ValueRange vrColDet = grColDet.Execute();
            IList<IList<object>> iioColDet = vrColDet.Values;

            string tieDia = Convert.ToString(iioColDet[0][7]);
            double longSd = Convert.ToDouble(iioColDet[0][1]);
            double dbConf = Convert.ToDouble(iioColDet[0][10]);
            string colSpc = Convert.ToString(iioColDet[0][9]);
            double numOfConf = Math.Ceiling(longSd/dbConf);
            string colCT = $"{tieDia}%%C - {Convert.ToInt16(numOfConf)} @ {dbConf}, REST @ \\P{colSpc} O.C. to C.L.";


            using (Transaction trColText = aCurDB.TransactionManager.StartTransaction())
            {
                BlockTable blktbl;
                blktbl = trColText.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord bltrec;
                bltrec = trColText.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                Entity textObj = trColText.GetObject(ssColText[0].ObjectId, OpenMode.ForWrite) as Entity;
                if (textObj is DBText)
                {
                    DBText dbText = (DBText)textObj;
                    dbText.TextString = colCT;
                }
                else if (textObj is MText)
                {
                    MText mText = (MText)textObj;
                    mText.Contents = colCT;
                }
                else {
                    actDoc.Editor.WriteMessage("\nSelected object is not a text. Quit...");
                    return;
                }

                trColText.Commit();
            }

        }
    }
}
