using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.Colors;

using Google.Apis.Sheets.v4;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using Google.Apis.Sheets.v4.Data;

using System.Threading;
using System.Xml.Linq;
using acColor = Autodesk.AutoCAD.Colors.Color;
using Autodesk.AutoCAD.DatabaseServices.Filters;
using Autodesk.AutoCAD.Windows.Data;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;
using WF = System.Windows.Forms;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.Interop.Common;
using WinApp = System.Windows.Forms;
/**
 * 06 Aug 2025 6:34am 3CPTT BAC
 * AI: It's generally not recommended to add WinForms to a class library project. 
 * Instead, create a separate WinForms project and reference your class library.
 * using EASI_CAD31.Forms; //<-- this one is removed due to cited reason above.
 */


namespace EASI_CAD31
{
    public class ConcCuttingList
    {
        static Document actDoc = Application.DocumentManager.MdiActiveDocument;
        static Database aCurDB = actDoc.Database;

        static CCData ccData = new CCData();

        //Default sheet address
        //static string shAddrG = "Sheet1!A2:B2";
        //Default BBS info
        static string bbsInfoG = "Tab=1,Row=19,Seg=A";


        // Default google sheet
        /* https://docs.google.com/spreadsheets/d/1xmapL82ZIP0CPuvVMQmPWKPLXpsyzOANs6x5gSyPdFQ/edit#gid=2098576256 */
        static string gshtIdG = DataGlobal.bbsGSheetId;
        static double beamHG = 600.0; // Drawing height
        static double barDiaG = 25.0;
        static double lapK1G = 48.5;
        static double bDepthG = 600.0;
        static double webDiaG = 10;
        static double stiDiaG = 10;
        static double barXDiG = 371;
        static double barYDiG = 471;
        static double txtHgtG = 250;
        static string infoStartG = "Left";
        static string selectBarTypeG = "Web";
        static double concCoverG = 40.0;
        static int gsRowG = 3;
        static int gsRowBBSG = 19;
        static string tabBBSG = "1";
        static string bbsBeamMarkG = "G2-1";
        static int clStartingRowG = 3;
        static int clEndingRowG = 12;
        static string stirrupDrawLineG = "Center";

        const int txtDimCol  = 10;
        const int txtMarkCol = 50;
        const int txtInfCol  = 193;
        const int beamDimenSourceColor = 1;
        const int beamMarkSourceColor = 3;
        const int beamRebarInfoColor = 200;

        const int webLinCol  = 20;
        const int countedWebLineColor = 21;
        const int botBarCol  = 130;
        const int botLinCol  = 53;
        const int botExtCol  = 140;
        const int endLinCol  = 54;
        const int topBarCol  = 4;
        const int topExtCol  = 120;
        const int topLinCol  = 52;
        const int stirTxCol  = 3;
        const int stirLnCol  = 251;
        //const int stirrupLineSectionColor = 51;
        const int stirrupSectionColor = 51;
        const int labelColor = 80; //3;
        const int labelCodeColor = 181;
        const int slabBarLineColor = 220;
        const int tableLineColor = 1;
        const int tableHeaderTextColor = 4;
        const int tableTextColor = 3;

        //Cutting List
        static string clGSheetIDG = "1WdD0AdxEHthbzz4g0CkAXNs_ywVOnjXYaOIUmXfEsOg";
        static string bbsGSheetIDG = "1xmapL82ZIP0CPuvVMQmPWKPLXpsyzOANs6x5gSyPdFQ";
        /**
         * Date added: 09 May 2024 11:22am
         * Added by: Bernardo A. Cabebe Jr.
         * Venue: 31H One Orchard
         * Status: Done - 09 May 2024 2:51pm
         */
        [CommandMethod("BBC3_BottomBarContinuous3")]
        public static void BottomBarContinuous3()
        {
            if (!ccData.isLicenseActive())
            {
                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }


            TypedValue[] tvFilterS8r = new TypedValue[2];
            tvFilterS8r[0] = new TypedValue((int)DxfCode.Start, "LINE");
            tvFilterS8r[1] = new TypedValue((int)DxfCode.Color, botLinCol);
            SelectionFilter sfBBLS8r = new SelectionFilter(tvFilterS8r);

            PromptSelectionOptions psoBBLS8r = new PromptSelectionOptions(); //Bottom Beam Line
            psoBBLS8r.MessageForAdding = "\nSelect bottom beam lines: ";

            PromptSelectionResult psrBBLS8r = actDoc.Editor.GetSelection(psoBBLS8r, sfBBLS8r);
            if (psrBBLS8r.Status != PromptStatus.OK) return;

            SelectionSet ssBBLS8r = psrBBLS8r.Value;


            //Get from user the rebar diameter and lap factor
            PromptStringOptions pso2BBLS8r = new PromptStringOptions("");
            string defVBBLS8r = barDiaG.ToString() + "," + lapK1G.ToString() + "," + bDepthG.ToString() + "," + beamHG.ToString();
            pso2BBLS8r.Message = "\nEnter bar dia, lap factor, beam depth D and drawing height H: ";
            pso2BBLS8r.DefaultValue = defVBBLS8r;
            pso2BBLS8r.AllowSpaces = false;

            PromptResult pr2BBLS8r = actDoc.Editor.GetString(pso2BBLS8r);
            if (pr2BBLS8r.Status != PromptStatus.OK) return;
            string strBBLS8r = pr2BBLS8r.StringResult;
            string[] arrStrS8r = strBBLS8r.Split(new char[] { ',' });
            double barDia = Convert.ToDouble(arrStrS8r[0]);
            double lapK1 = Convert.ToDouble(arrStrS8r[1]);
            double bDepth = Convert.ToDouble(arrStrS8r[2]);
            double dBeamH = Convert.ToDouble(arrStrS8r[3]); //Drawing beam height
            actDoc.Editor.WriteMessage("\nBeam depth: {0}; Drawing beam height: {1}", bDepth, dBeamH);
            barDiaG = barDia;
            lapK1G = lapK1;
            bDepthG = bDepth;
            beamHG = dBeamH;
            double offDS8r = dBeamH / 6;
            double lapLenS8r = barDia * lapK1;

            //Under development as of 09 May 2024
            using (Transaction trBBLS8r = aCurDB.TransactionManager.StartTransaction())
            {
                List<List<object>> liliBBL = new List<List<object>>();
                foreach (SelectedObject soBBLS8r in ssBBLS8r)
                {
                    Curve crBBLS8r = trBBLS8r.GetObject(soBBLS8r.ObjectId, OpenMode.ForRead) as Curve;
                    Point3d startPt = crBBLS8r.StartPoint;
                    Point2d sttPt2d = new Point2d(startPt.X, startPt.Y + offDS8r);
                    //actDoc.Editor.WriteMessage("\nStart point: {0}", startPt);
                    Point3d endPt = crBBLS8r.EndPoint;
                    Point2d endPt2d = new Point2d(endPt.X, endPt.Y + offDS8r);
                    //actDoc.Editor.WriteMessage("\nEnd point: {0}", endPt);
                    double parA = crBBLS8r.StartParam;
                    double parB = crBBLS8r.EndParam;
                    double dblLen = crBBLS8r.GetDistanceAtParameter(parB) - crBBLS8r.GetDistanceAtParameter(parA);
                    double dbl2H1 = bDepth * 2;
                    double dbl2H2 = lapLenS8r + dbl2H1;
                    Point3d pnt31 = crBBLS8r.GetPointAtDist(dbl2H1);
                    Point2d p22H1 = new Point2d(pnt31.X, pnt31.Y + offDS8r);

                    Point3d pnt32 = crBBLS8r.GetPointAtDist(dbl2H2);
                    Point2d p22H2 = new Point2d(pnt32.X, pnt32.Y + offDS8r);

                    Point3d midPt = Midpoint(startPt, endPt);

                    if(dblLen > dbl2H2)
                    {
                        liliBBL.Add(new List<object>() { midPt.X, sttPt2d, p22H1, p22H2, endPt2d });
                        // actDoc.Editor.WriteMessage("\nMidpoint: {0}", midPt.X);
                        // actDoc.Editor.WriteMessage("\n----------#-#-#----------");
                    }
                    else
                    {
                        // Date added: 22 May 2024 by BACJr
                        actDoc.Editor.WriteMessage("\nHighlighted object is not included because length is less than 2H");
                        crBBLS8r.Highlight();
                    }

                }

                List<List<object>> sliliBBL = new List<List<object>>();
                sliliBBL = liliBBL.OrderBy(a => a[0]).ToList();

                BlockTable blktbl;
                blktbl = trBBLS8r.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;

                BlockTableRecord bltrec;
                bltrec = trBBLS8r.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                for (int aa = 0; aa < sliliBBL.Count; aa++)
                {
                    double dblLenS8r = 0;
                    using (Polyline acPoly = new Polyline())
                    {
                        if (aa == 0)
                        {
                            acPoly.AddVertexAt(0, (Point2d)sliliBBL[aa][2], 0, 0, 0);
                            acPoly.AddVertexAt(1, (Point2d)sliliBBL[aa + 1][3], 0, 0, 0);
                        }
                        else if (aa == (sliliBBL.Count - 1))
                        {
                            acPoly.AddVertexAt(0, (Point2d)sliliBBL[aa][2], 0, 0, 0);
                            acPoly.AddVertexAt(1, (Point2d)sliliBBL[aa][4], 0, 0, 0);
                        }
                        else
                        {
                            acPoly.AddVertexAt(0, (Point2d)sliliBBL[aa][2], 0, 0, 0);
                            acPoly.AddVertexAt(1, (Point2d)sliliBBL[aa + 1][3], 0, 0, 0);
                        }

                        dblLenS8r = acPoly.Length;

                        acPoly.ColorIndex = botBarCol;
                        bltrec.AppendEntity(acPoly);
                        trBBLS8r.AddNewlyCreatedDBObject(acPoly, true);
                    }

                    actDoc.Editor.WriteMessage("\nBar {0} length: {1}", aa, dblLenS8r);
                }

                trBBLS8r.Commit();
            }
        }



        /**
         * Date added: 06 May 2024 11:12pm
         * Added by: Bernardo A. Cabebe Jr.
         * Venue: 3407 Cityland Pasong Tamo Tower
         * Status: Done - 08 May 2024 11:02am
         */
        [CommandMethod("BLS_BarLapSplice")]
        public static void BarLapSplice()
        {
            if (!ccData.isLicenseActive())
            {
                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }


            //Select entity 1 (Bar line 1)
            PromptEntityOptions peoBLS1Y9t = new PromptEntityOptions("Select near end of rebar line 1: ");
            peoBLS1Y9t.SetRejectMessage("\nSelect polyline only.");
            peoBLS1Y9t.AddAllowedClass(typeof(Polyline), true);

            PromptEntityResult perBLS1Y9t = actDoc.Editor.GetEntity(peoBLS1Y9t);
            if (perBLS1Y9t.Status != PromptStatus.OK) return;


            //Select entity 2 (Bar line 2)
            PromptEntityOptions peoBLS2Y9t = new PromptEntityOptions("Select near end of rebar line 2: ");
            peoBLS2Y9t.SetRejectMessage("\nSelect polyline only.");
            peoBLS2Y9t.AddAllowedClass(typeof(Polyline), true);

            PromptEntityResult perBLS2Y9t = actDoc.Editor.GetEntity(peoBLS2Y9t);
            if (perBLS2Y9t.Status != PromptStatus.OK) return;


            //Get from user the rebar diameter and lap factor
            PromptStringOptions psoBLSY9t = new PromptStringOptions("");
            string defVBLSY9t = barDiaG.ToString() + "," + lapK1G.ToString();
            psoBLSY9t.Message = "\nEnter bar dia and lap factor: ";
            psoBLSY9t.DefaultValue = defVBLSY9t;
            psoBLSY9t.AllowSpaces = false;

            PromptResult prBLSY9t = actDoc.Editor.GetString(psoBLSY9t);
            if (prBLSY9t.Status != PromptStatus.OK) return;
            string strBLSY9t = prBLSY9t.StringResult;
            string[] arrStrY9t = strBLSY9t.Split(new char[] { ',' });
            double barDia = Convert.ToDouble(arrStrY9t[0]);
            double lapK1 = Convert.ToDouble(arrStrY9t[1]);
            actDoc.Editor.WriteMessage("\nBar D: {0}; Lap factor: {1}", barDia, lapK1);
            barDiaG = barDia;
            lapK1G = lapK1;


            using (Transaction trBLSY9t = aCurDB.TransactionManager.StartTransaction())
            {
                //Get the end point of bar lines nearest to the associated pick point
                Point3d pickPt1 = perBLS1Y9t.PickedPoint;
                Point3d pickPt2 = perBLS2Y9t.PickedPoint;

                Polyline plBLS1Y9t = trBLSY9t.GetObject(perBLS1Y9t.ObjectId, OpenMode.ForRead) as Polyline;
                Curve cuBLS1Y9t = trBLSY9t.GetObject(perBLS1Y9t.ObjectId, OpenMode.ForRead) as Curve;

                Polyline plBLS2Y9t = trBLSY9t.GetObject(perBLS2Y9t.ObjectId, OpenMode.ForRead) as Polyline;
                Curve cuBLS2Y9t = trBLSY9t.GetObject(perBLS2Y9t.ObjectId, OpenMode.ForRead) as Curve;


                //Get the distance or length between the identified two end points
                Point3d nrstBLS1 = new Point3d();
                if (plBLS1Y9t.StartPoint.DistanceTo(pickPt1) < plBLS1Y9t.EndPoint.DistanceTo(pickPt1))
                {
                    nrstBLS1 = plBLS1Y9t.StartPoint;
                }
                else
                {
                    nrstBLS1 = plBLS1Y9t.EndPoint;
                }

                Point3d nrstBLS2 = new Point3d();
                if (plBLS2Y9t.StartPoint.DistanceTo(pickPt2) < plBLS2Y9t.EndPoint.DistanceTo(pickPt2))
                {
                    nrstBLS2 = plBLS2Y9t.StartPoint;
                }
                else
                {
                    nrstBLS2 = plBLS2Y9t.EndPoint;
                }


                /**
                 * Check if the distance or length satisfies the lap length. If the length is shorter than the 
                 * lap length, increase the distance or length; otherwise, do vise versa. The adjustment of identified
                 * end points shall be made with respect to its paired end point so that proper direction of 
                 * adjustment of end points can be performed.
                 */
                double actSpLenY9t = nrstBLS1.DistanceTo(nrstBLS2);
                double spliceLeY9t = lapK1 * barDia;
                double deltaLap = spliceLeY9t - actSpLenY9t;

                /**
                 * Redefine the vertex of bar line 1 and bar line 2
                 * https://forums.autodesk.com/t5/net/modifying-polyline-vertex/td-p/2534299
                 */
                plBLS1Y9t.UpgradeOpen();
                double indexBLS1 = 0.0;
                int vtxBLS1 = 0;
                if (plBLS1Y9t.StartPoint.DistanceTo(pickPt1) < plBLS1Y9t.EndPoint.DistanceTo(pickPt1))
                {
                    Point2d newPtBLS1 = new Point2d((nrstBLS1.X - deltaLap * 0.5), nrstBLS1.Y);
                    indexBLS1 = plBLS1Y9t.StartParam;
                    vtxBLS1 = (int)indexBLS1;
                    plBLS1Y9t.SetPointAt(vtxBLS1, newPtBLS1);
                }
                else
                {
                    Point2d newPtBLS1 = new Point2d((nrstBLS1.X - deltaLap * 0.5), nrstBLS1.Y);
                    indexBLS1 = plBLS1Y9t.EndParam;
                    vtxBLS1 = (int)indexBLS1;
                    plBLS1Y9t.SetPointAt(vtxBLS1, newPtBLS1);
                }

                plBLS2Y9t.UpgradeOpen();
                double indexBLS2 = 0.0;
                int vtxBLS2 = 0;
                if (plBLS2Y9t.StartPoint.DistanceTo(pickPt2) < plBLS2Y9t.EndPoint.DistanceTo(pickPt2))
                {
                    Point2d newPtBLS2 = new Point2d((nrstBLS2.X + deltaLap * 0.5), nrstBLS2.Y);
                    indexBLS2 = plBLS2Y9t.StartParam;
                    vtxBLS2 = (int)indexBLS2;
                    plBLS2Y9t.SetPointAt(vtxBLS2, newPtBLS2);
                }
                else
                {
                    Point2d newPtBLS2 = new Point2d((nrstBLS2.X + deltaLap * 0.5), nrstBLS2.Y);
                    indexBLS2 = plBLS2Y9t.EndParam;
                    vtxBLS2 = (int)indexBLS2;
                    plBLS2Y9t.SetPointAt(vtxBLS2, newPtBLS2);
                }


                //Lengths of line 1 and 2
                double lenLine1 = plBLS1Y9t.Length;
                actDoc.Editor.WriteMessage("\nLine 1 length: {0}", lenLine1);
                double lenLine2 = plBLS2Y9t.Length;
                actDoc.Editor.WriteMessage("\nLine 2 length: {0}", lenLine2);


                BlockTable blktbl;
                blktbl = trBLSY9t.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;

                BlockTableRecord bltrec;
                bltrec = trBLSY9t.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                // using (DBPoint acPoint = new DBPoint(pickPt1))
                // {
                //     acPoint.ColorIndex = 250;
                //     bltrec.AppendEntity(acPoint);
                //     trBLSY9t.AddNewlyCreatedDBObject(acPoint, true);
                // }

                trBLSY9t.Commit();

            }

        }

        /**
         * Date added: 06 May 2024
         * Added by: Bernardo A. Cabebe Jr.
         * Venue: 3407 Cityland Pasong Tamo Tower
         * Status: Done 06 May 2024 10:57pm
         *         Modified 08 May 2024 11:02am
         */
        [CommandMethod("TBE_TopBarEnd")]
        public static void TopBarEnd()
        {
            if (!ccData.isLicenseActive())
            {
                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            TypedValue[] tvFilterL2w = new TypedValue[6];
            tvFilterL2w[0] = new TypedValue((int)DxfCode.Start, "*LINE");
            tvFilterL2w[1] = new TypedValue((int)DxfCode.Operator, "<OR");
            tvFilterL2w[2] = new TypedValue((int)DxfCode.Color, topBarCol);
            tvFilterL2w[3] = new TypedValue((int)DxfCode.Color, endLinCol);
            tvFilterL2w[4] = new TypedValue((int)DxfCode.Color, topExtCol);
            tvFilterL2w[5] = new TypedValue((int)DxfCode.Operator, "OR>");
            SelectionFilter sfTBEL2w = new SelectionFilter(tvFilterL2w);

            PromptSelectionOptions psoTBEL2w = new PromptSelectionOptions(); //Top Beam Line
            psoTBEL2w.MessageForAdding = "\nSelect end line and bar line: ";

            PromptSelectionResult psrTBEL2w = actDoc.Editor.GetSelection(psoTBEL2w, sfTBEL2w);
            if (psrTBEL2w.Status != PromptStatus.OK) return;

            SelectionSet ssTBEL2w = psrTBEL2w.Value;

            if (ssTBEL2w.Count < 2)
            {
                actDoc.Editor.WriteMessage("\nInvalid selection.");
                return;
            }

            PromptDoubleOptions pdoTBEL2w = new PromptDoubleOptions("");
            pdoTBEL2w.DefaultValue = barDiaG;
            pdoTBEL2w.Message = "\nEnter bar diameter: ";
            pdoTBEL2w.AllowNegative = false;
            pdoTBEL2w.AllowZero = false;

            PromptDoubleResult pdrTBEL2w = actDoc.Editor.GetDouble(pdoTBEL2w);
            if (pdrTBEL2w.Status != PromptStatus.OK) return;
            double barDia = pdrTBEL2w.Value;
            barDiaG = barDia;

            using (Transaction trTBEL2w = aCurDB.TransactionManager.StartTransaction())
            {
                Curve crLine = null;
                Polyline plBar = null;
                List<Point3d> liPLVtx = new List<Point3d>();
                int barColor = 8;
                foreach (SelectedObject soTBEL2w in ssTBEL2w)
                {
                    Curve crTBEL2w = trTBEL2w.GetObject(soTBEL2w.ObjectId, OpenMode.ForRead) as Curve;
                    if (crTBEL2w.ColorIndex == topBarCol || crTBEL2w.ColorIndex == topExtCol) //if (crTBEL2w.ColorIndex == 4 || crTBEL2w.ColorIndex == 120)
                    {
                        plBar = trTBEL2w.GetObject(soTBEL2w.ObjectId, OpenMode.ForRead) as Polyline;
                        barColor = plBar.ColorIndex;
                        for (int aa = 0; aa < plBar.NumberOfVertices; aa++)
                        {
                            liPLVtx.Add(plBar.GetPoint3dAt(aa));
                        }
                    }

                    if (crTBEL2w.ColorIndex == endLinCol) //54)
                    {
                        crLine = crTBEL2w;
                    }
                }

                Point3d closestPt = crLine.GetClosestPointTo(plBar.StartPoint, true);
                if (closestPt.DistanceTo(liPLVtx[0]) < closestPt.DistanceTo(liPLVtx[liPLVtx.Count - 1]))
                {
                    Vector3d v3Vtx0CP = closestPt.GetVectorTo(liPLVtx[0]);
                    Vector3d nv3Tv0CP = v3Vtx0CP.GetNormal();
                    double offD = 0;
                    if (barDia > 0.95) //No bar diameter greater than 95mm
                    {
                        offD = 40 + 12 + 0.5 * barDia; //mm unit
                    }
                    else
                    {
                        offD = 0.040 + 0.012 + 0.5 * barDia; //m unit
                    }

                    nv3Tv0CP = nv3Tv0CP * offD;

                    Point3d p3Vtx0 = closestPt + nv3Tv0CP;
                    Point3d p3VtNS = p3Vtx0 + new Vector3d(0, -12 * barDia, 0); //New start vertext
                    liPLVtx[0] = p3Vtx0;
                    liPLVtx.Insert(0, p3VtNS);
                }
                else
                {
                    Vector3d v3VtxnCP = closestPt.GetVectorTo(liPLVtx[liPLVtx.Count - 1]);
                    Vector3d nv3TvnCP = v3VtxnCP.GetNormal();
                    double offD = 0;
                    if (barDia > 0.95) //No bar diameter greater than 95mm
                    {
                        offD = 40 + 12 + 0.5 * barDia; //mm unit
                    }
                    else
                    {
                        offD = 0.040 + 0.012 + 0.5 * barDia; //m unit
                    }

                    nv3TvnCP = nv3TvnCP * offD;

                    Point3d p3Vtxn = closestPt + nv3TvnCP;
                    Point3d p3VtNE = p3Vtxn + new Vector3d(0, -12 * barDia, 0); //New end vertext
                    liPLVtx[liPLVtx.Count - 1] = p3Vtxn;
                    liPLVtx.Add(p3VtNE);
                }

                plBar.UpgradeOpen();

                plBar.Erase();

                //Convert List<Point3d> to List<Point2d>
                List<Point2d> liVtx2d = new List<Point2d>();
                for (int bb = 0; bb < liPLVtx.Count; bb++)
                {
                    liVtx2d.Add(new Point2d(liPLVtx[bb].X, liPLVtx[bb].Y));
                }

                BlockTable blktbl;
                blktbl = trTBEL2w.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;

                BlockTableRecord bltrec;
                bltrec = trTBEL2w.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                double lineLen = 0;
                using (Polyline acPoly = new Polyline())
                {
                    for (int cc = 0; cc < liVtx2d.Count; cc++)
                    {
                        acPoly.AddVertexAt(0, liVtx2d[cc], 0, 0, 0);
                    }

                    acPoly.ColorIndex = barColor;
                    lineLen = acPoly.Length;
                    bltrec.AppendEntity(acPoly);
                    trTBEL2w.AddNewlyCreatedDBObject(acPoly, true);
                }

                actDoc.Editor.WriteMessage("\nLine length: {0}", lineLen);

                trTBEL2w.Commit();
            }

        }

        /**
         * Date added: 05 May 2024
         * Added by: Bernardo A. Cabebe Jr.
         * Venue: Marquinton Carwash
         * Status: Done - 06 May 2024
         */
        [CommandMethod("TBC3_TopBarContinuous3")]
        public static void TopBarContinuous3()
        {
            if (!ccData.isLicenseActive())
            {
                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            PromptDoubleOptions pdoH = new PromptDoubleOptions("");
            pdoH.Message = "\nEnter beam depth: ";
            pdoH.DefaultValue = beamHG;
            pdoH.AllowZero = false;
            pdoH.AllowNegative = false;

            PromptDoubleResult pdrH = actDoc.Editor.GetDouble(pdoH);
            if (pdrH.Status != PromptStatus.OK) return;
            double dblH = pdrH.Value;
            beamHG = dblH;
            double offD = dblH / 6;

            TypedValue[] tvFilterB3t = new TypedValue[2];
            tvFilterB3t[0] = new TypedValue((int)DxfCode.Start, "LINE");
            tvFilterB3t[1] = new TypedValue((int)DxfCode.Color, topLinCol);
            SelectionFilter sfTBLB3t = new SelectionFilter(tvFilterB3t);

            PromptSelectionOptions psoTBLB3t = new PromptSelectionOptions(); //Top Beam Line
            psoTBLB3t.MessageForAdding = "\nSelect top beam lines: ";

            PromptSelectionResult psrTBLB3t = actDoc.Editor.GetSelection(psoTBLB3t, sfTBLB3t);
            if (psrTBLB3t.Status != PromptStatus.OK) return;

            SelectionSet ssTBLB3t = psrTBLB3t.Value;

            if (ssTBLB3t.Count < 1)
            {
                actDoc.Editor.WriteMessage("\nSelect at least one top beam line.");
                return;
            }

            using (Transaction trTBLB3t = aCurDB.TransactionManager.StartTransaction())
            {
                List<List<object>> liliTBL = new List<List<object>>();
                foreach (SelectedObject soTBLB3t in ssTBLB3t)
                {
                    Curve crTBLB3t = trTBLB3t.GetObject(soTBLB3t.ObjectId, OpenMode.ForRead) as Curve;
                    Point3d startPt = crTBLB3t.StartPoint;
                    Point2d sttPt2d = new Point2d(startPt.X, startPt.Y - offD);
                    //actDoc.Editor.WriteMessage("\nStart point: {0}", startPt);
                    Point3d endPt = crTBLB3t.EndPoint;
                    Point2d endPt2d = new Point2d(endPt.X, endPt.Y - offD);
                    //actDoc.Editor.WriteMessage("\nEnd point: {0}", endPt);
                    double parA = crTBLB3t.StartParam;
                    double parB = crTBLB3t.EndParam;
                    double dblLen = crTBLB3t.GetDistanceAtParameter(parB) - crTBLB3t.GetDistanceAtParameter(parA);
                    double dblL31 = dblLen / 3;
                    double dblL32 = 2 * dblLen / 3;
                    Point3d pnt31 = crTBLB3t.GetPointAtDist(dblL31);
                    Point2d p2d31 = new Point2d(pnt31.X, pnt31.Y - offD);

                    Point3d pnt32 = crTBLB3t.GetPointAtDist(dblL32);
                    Point2d p2d32 = new Point2d(pnt32.X, pnt32.Y - offD);

                    Point3d midPt = Midpoint(startPt, endPt);

                    liliTBL.Add(new List<object>() { midPt.X, sttPt2d, p2d31, p2d32, endPt2d });
                    // actDoc.Editor.WriteMessage("\nMidpoint: {0}", midPt.X);
                    // actDoc.Editor.WriteMessage("\n----------#-#-#----------");
                }
                // for (int aa = 0; aa < liliTBL.Count; aa++)
                // {
                //     actDoc.Editor.WriteMessage("\nX{0}: {1}", aa, liliTBL[aa][0]);
                // }

                List<List<object>> sliliTBL = new List<List<object>>();
                sliliTBL = liliTBL.OrderBy(a => a[0]).ToList();
                // actDoc.Editor.WriteMessage("\n----------#-#-#----------");
                // for (int aa = 0; aa < liliTBL.Count; aa++)
                // {
                //     actDoc.Editor.WriteMessage("\nX{0}: {1}", aa, sliliTBL[aa][0]);
                // }

                BlockTable blktbl;
                blktbl = trTBLB3t.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;

                BlockTableRecord bltrec;
                bltrec = trTBLB3t.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                for (int aa = 0; aa < sliliTBL.Count; aa++)
                {

                    using (Polyline acPoly = new Polyline())
                    {
                        if (aa == 0)
                        {
                            if (sliliTBL.Count == 1)
                            {
                                acPoly.AddVertexAt(0, (Point2d)sliliTBL[aa][1], 0, 0, 0);
                                acPoly.AddVertexAt(1, (Point2d)sliliTBL[aa][4], 0, 0, 0);
                            }
                            else
                            {
                                acPoly.AddVertexAt(0, (Point2d)sliliTBL[aa][1], 0, 0, 0);
                                acPoly.AddVertexAt(1, (Point2d)sliliTBL[aa][3], 0, 0, 0);
                            }
                        }
                        else if (aa == (sliliTBL.Count - 1))
                        {
                            acPoly.AddVertexAt(0, (Point2d)sliliTBL[aa][2], 0, 0, 0);
                            acPoly.AddVertexAt(1, (Point2d)sliliTBL[aa][4], 0, 0, 0);
                        }
                        else
                        {
                            acPoly.AddVertexAt(0, (Point2d)sliliTBL[aa][2], 0, 0, 0);
                            acPoly.AddVertexAt(1, (Point2d)sliliTBL[aa + 1][3], 0, 0, 0);
                        }

                        acPoly.ColorIndex = topBarCol;
                        bltrec.AppendEntity(acPoly);
                        trTBLB3t.AddNewlyCreatedDBObject(acPoly, true);
                    }

                    if (aa == 0 && sliliTBL.Count > 1)
                    {
                        using (Polyline acPoly2 = new Polyline())
                        {
                            acPoly2.AddVertexAt(0, (Point2d)sliliTBL[aa][2], 0, 0, 0);
                            acPoly2.AddVertexAt(1, (Point2d)sliliTBL[aa + 1][3], 0, 0, 0);

                            acPoly2.ColorIndex = topBarCol; //4;
                            bltrec.AppendEntity(acPoly2);
                            trTBLB3t.AddNewlyCreatedDBObject(acPoly2, true);
                        }
                    }

                }

                trTBLB3t.Commit();
            }

        }

        private static Point3d Midpoint(Point3d Pt1, Point3d Pt2)
        {
            double mPtX = (Pt2.X + Pt1.X) / 2;
            double mPtY = (Pt2.Y + Pt1.Y) / 2;
            double mPtZ = (Pt2.Z + Pt1.Z) / 2;
            return new Point3d(mPtX, mPtY, mPtZ);
        }

        /**
         * Date added: 23 Mar 2024
         * Added by: Bernardo A. Cabebe Jr.
         * Venue: Dunkin Donut Marcos Highway
         * Status: Under development
         */
        [CommandMethod("CLS_CreateLabelSingleBar")]
        public static void CreateLabelSingleBar()
        {
            if (!ccData.isLicenseActive())
            {

                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            TypedValue[] tvFilterS6y = new TypedValue[6];
            tvFilterS6y[0] = new TypedValue((int)DxfCode.Start, "MTEXT");
            tvFilterS6y[1] = new TypedValue((int)DxfCode.Operator, "<OR");
            tvFilterS6y[2] = new TypedValue((int)DxfCode.Color, txtInfCol);
            tvFilterS6y[3] = new TypedValue((int)DxfCode.Color, txtMarkCol);
            tvFilterS6y[4] = new TypedValue((int)DxfCode.Color, txtDimCol);
            tvFilterS6y[5] = new TypedValue((int)DxfCode.Operator, "OR>");
            SelectionFilter sfCLLS6y = new SelectionFilter(tvFilterS6y);

            PromptSelectionOptions psoCLLS6y = new PromptSelectionOptions();
            psoCLLS6y.MessageForAdding = "\nSelect beam informations: ";

            PromptSelectionResult psrCLLS6y = actDoc.Editor.GetSelection(psoCLLS6y, sfCLLS6y);
            if (psrCLLS6y.Status != PromptStatus.OK) return;

            SelectionSet ssCLLS6y = psrCLLS6y.Value;
            if(ssCLLS6y.Count != 7)
            {
                actDoc.Editor.WriteMessage("\nInvalid number of selection.");
                return;
            }

            PromptKeywordOptions pkoCLLS6y = new PromptKeywordOptions("");
            pkoCLLS6y.Message = "\nEnter section: ";
            pkoCLLS6y.Keywords.Add("1");
            pkoCLLS6y.Keywords.Add("2");
            pkoCLLS6y.Keywords.Add("Other");
            pkoCLLS6y.AllowNone = false;

            PromptResult prCLLS6y = actDoc.Editor.GetKeywords(pkoCLLS6y);
            string beamSection = prCLLS6y.StringResult;

            Transaction trLL = aCurDB.TransactionManager.StartTransaction();
            using (trLL)
            {
                int e1Top = 0;
                int e1Bot = 0;
                int e1Dia = 0;
                int e2Top = 0;
                int e2Bot = 0;
                int e2Dia = 0;
                int msTop = 0;
                int msBot = 0;
                int msDia = 0;
                string stirrV = "";
                foreach (SelectedObject soMTx in ssCLLS6y)
                {
                    MText aaMTx = trLL.GetObject(soMTx.ObjectId, OpenMode.ForRead) as MText;
                    string mtxCont = aaMTx.Text;
                    actDoc.Editor.WriteMessage("\nText: {0}", mtxCont);

                    string patt1A = @"E1\:[0-9\-\/]+";
                    Regex regex1A = new Regex(patt1A);
                    MatchCollection match1A = regex1A.Matches(mtxCont);
                    
                    if(match1A.Count > 0)
                    {
                        actDoc.Editor.WriteMessage("\nE1: {0}", match1A[0].Value);
                        string pattE1Top = @"(?<=\-)\d+";
                        Regex regexE1Top = new Regex(pattE1Top);
                        Match matchE1Top = regexE1Top.Match(mtxCont);
                        e1Top = Convert.ToInt16(matchE1Top.Value);

                        string pattE1Bot = @"(?<=\/)\d+";
                        Regex regexE1Bot = new Regex(pattE1Bot);
                        Match matchE1Bot = regexE1Bot.Match(mtxCont);
                        e1Bot = Convert.ToInt16(matchE1Bot.Value);

                        string pattE1Dia = @"(?<=\:)\d+";
                        Regex regexE1Dia = new Regex(pattE1Dia);
                        Match matchE1Dia = regexE1Dia.Match(mtxCont);
                        e1Dia = Convert.ToInt16(matchE1Dia.Value);
                        actDoc.Editor.WriteMessage("\nEnd1 info: Dia={0}, Top={1}, Bot={2}", e1Dia, e1Top, e1Bot);
                    }

                    string patt2A = @"E2\:[0-9\-\/]+";
                    Regex regex2A = new Regex(patt2A);
                    MatchCollection match2A = regex2A.Matches(mtxCont);
                    
                    if (match2A.Count > 0)
                    {
                        actDoc.Editor.WriteMessage("\nE2: {0}", match2A[0].Value);
                        string pattE2Top = @"(?<=\-)\d+";
                        Regex regexE2Top = new Regex(pattE2Top);
                        Match matchE2Top = regexE2Top.Match(mtxCont);
                        e2Top = Convert.ToInt16(matchE2Top.Value);

                        string pattE2Bot = @"(?<=\/)\d+";
                        Regex regexE2Bot = new Regex(pattE2Bot);
                        Match matchE2Bot = regexE2Bot.Match(mtxCont);
                        e2Bot = Convert.ToInt16(matchE2Bot.Value);

                        string pattE2Dia = @"(?<=\:)\d+";
                        Regex regexE2Dia = new Regex(pattE2Dia);
                        Match matchE2Dia = regexE2Dia.Match(mtxCont);
                        e2Dia = Convert.ToInt16(matchE2Dia.Value);
                        actDoc.Editor.WriteMessage("\nEnd2 info: Dia={0}, Top={1}, Bot={2}", e2Dia, e2Top, e2Bot);
                    }

                    string patt3A = @"MS\:[0-9\-\/]+";
                    Regex regex3A = new Regex(patt3A);
                    MatchCollection match3A = regex3A.Matches(mtxCont);
                    
                    if (match3A.Count > 0)
                    {
                        actDoc.Editor.WriteMessage("\nMS: {0}", match3A[0].Value);
                        string pattMSTop = @"(?<=\-)\d+";
                        Regex regexMSTop = new Regex(pattMSTop);
                        Match matchMSTop = regexMSTop.Match(mtxCont);
                        msTop = Convert.ToInt16(matchMSTop.Value);

                        string pattMSBot = @"(?<=\/)\d+";
                        Regex regexMSBot = new Regex(pattMSBot);
                        Match matchMSBot = regexMSBot.Match(mtxCont);
                        msBot = Convert.ToInt16(matchMSBot.Value);

                        string pattMSDia = @"(?<=\:)\d+";
                        Regex regexMSDia = new Regex(pattMSDia);
                        Match matchMSDia = regexMSDia.Match(mtxCont);
                        msDia = Convert.ToInt16(matchMSDia.Value);
                        actDoc.Editor.WriteMessage("\nMidspan info: Dia={0}, Top={1}, Bot={2}", msDia, msTop, msBot);
                    }

                    string patt4A = @"Stirr\.\:[A-Z]+";
                    Regex regex4A = new Regex(patt4A);
                    MatchCollection match4A = regex4A.Matches(mtxCont);
                    
                    if (match4A.Count > 0)
                    {
                        actDoc.Editor.WriteMessage("\nMS: {0}", match4A[0].Value);
                        string pattStTop = @"(?<=\.\:)[A-Z]+";
                        Regex regexStTop = new Regex(pattStTop);
                        Match matchStTop = regexStTop.Match(mtxCont);
                        stirrV = matchStTop.Value;
                        actDoc.Editor.WriteMessage("\nStirrup: {0}", stirrV);
                    }
                }


                //Select entity 1 (Bar line 1)
                PromptEntityOptions peoCCLS6y = new PromptEntityOptions("\nSelect rebar line: ");
                peoCCLS6y.SetRejectMessage("\nSelect polyline only.");
                peoCCLS6y.AddAllowedClass(typeof(Polyline), true);

                PromptEntityResult perCCLS6y = actDoc.Editor.GetEntity(peoCCLS6y);
                if (perCCLS6y.Status != PromptStatus.OK) return;

                Point3d pickPt1 = perCCLS6y.PickedPoint;

                Curve cuBarLine = trLL.GetObject(perCCLS6y.ObjectId, OpenMode.ForRead) as Curve;
                Point3d p3dPt1 = cuBarLine.GetClosestPointTo(pickPt1,false);

                PromptPointOptions ppoPt2 = new PromptPointOptions("");
                ppoPt2.Message = "\nEnter point 2: ";
                ppoPt2.UseBasePoint = true;
                ppoPt2.BasePoint = pickPt1;
                PromptPointResult pprPt2 = actDoc.Editor.GetPoint(ppoPt2);
                if (pprPt2.Status != PromptStatus.OK) return;
                Point3d p3dPt2 = pprPt2.Value;

                PromptPointOptions ppoPt3 = new PromptPointOptions("");
                ppoPt3.Message = "\nEnter point 3: ";
                ppoPt3.UseBasePoint = true;
                ppoPt3.BasePoint = p3dPt2;
                PromptPointResult pprPt3 = actDoc.Editor.GetPoint(ppoPt3);
                if (pprPt3.Status != PromptStatus.OK) return;
                Point3d p3dPt3 = pprPt3.Value;

                PromptDoubleOptions pdoCLLS6y = new PromptDoubleOptions("");
                pdoCLLS6y.Message = "\nEnter text height: ";
                pdoCLLS6y.AllowNegative = false;
                pdoCLLS6y.AllowZero = false;
                pdoCLLS6y.DefaultValue = txtHgtG;
                pdoCLLS6y.AllowNone = true;

                PromptDoubleResult pdrCLLS6y = actDoc.Editor.GetDouble(pdoCLLS6y);
                if (pdrCLLS6y.Status  != PromptStatus.OK) return;

                double textHgt = pdrCLLS6y.Value;
                txtHgtG = textHgt;

                int barLineColor = cuBarLine.ColorIndex;
                double lineLength = cuBarLine.GetDistAtPoint(cuBarLine.EndPoint) - cuBarLine.GetDistAtPoint(cuBarLine.StartPoint);
                lineLength = Math.Round(lineLength, 2);
                if(lineLength > 100)
                {
                    lineLength = Math.Round(lineLength / 1000, 2);
                }
                string barDesc = "";

                List<int> liTopBarQty = new List<int>();
                liTopBarQty.Add(e1Top);
                liTopBarQty.Add(msTop);
                liTopBarQty.Add(e2Top);
                actDoc.Editor.WriteMessage("\nMin. top bar qty: {0}", liTopBarQty.Min());

                if(barLineColor == topBarCol && beamSection == "Other") //Continuous top bar
                {
                    barDesc = "CTB" + RebarId() + ":" + liTopBarQty.Min() + "-" + msDia + "D x " + lineLength.ToString();
                }

                if (barLineColor == topExtCol && beamSection == "1")
                {
                    barDesc = "ETB" + RebarId() + ":" + (e1Top-liTopBarQty.Min()) + "-" + e1Dia + "D x " + lineLength.ToString();
                }

                if (barLineColor == topExtCol && beamSection == "2")
                {
                    barDesc = "ETB" + RebarId() + ":" + (e2Top - liTopBarQty.Min()) + "-" + e2Dia + "D x " + lineLength.ToString();
                }

                if (barLineColor == botBarCol && beamSection == "1")
                {
                    barDesc = "CBB" + RebarId() + ":" + e1Bot + "-" + e1Dia + "D x " + lineLength.ToString();
                }

                if (barLineColor == botBarCol && beamSection == "2")
                {
                    barDesc = "CBB" + RebarId() + ":" + e2Bot + "-" + e2Dia + "D x " + lineLength.ToString();
                }

                if (barLineColor == botExtCol && beamSection == "Other")
                {
                    barDesc = "EBB" + RebarId() + ":" + (msBot - Math.Min(e1Bot, e2Bot)) + "-" + e2Dia + "D x " + lineLength.ToString();
                }

                MText mTxt = new MText();
                mTxt.Location = p3dPt3;
                // mTxt.Contents = "Some text in the default colour...\\P" +
                //                 "{\\C1;Something red}\\P" +
                //                 "{\\C2;Something yellow}\\P" +
                //                 "{\\C3;And} {\\C4;something} " +
                //                 "{\\C5;multi-}{\\C6;coloured}\\P";
                mTxt.Contents = barDesc; //"TEXT CONTENT 1 \\PTEXT CONTENT 2";

                if (p3dPt3.X > p3dPt2.X)
                {
                    mTxt.Attachment = AttachmentPoint.MiddleLeft;
                }
                else
                {
                    mTxt.Attachment = AttachmentPoint.MiddleRight;
                }
                mTxt.ColorIndex = labelColor; //3;

                double arrowSize = 0.75 * txtHgtG; //185;
                if (p3dPt3.DistanceTo(p3dPt2) > 50)
                {
                    mTxt.TextHeight = txtHgtG; //250;
                    arrowSize = 0.75 * txtHgtG; // 185;
                }
                else
                {
                    mTxt.TextHeight = txtHgtG / 1000; // 0.250;
                    arrowSize = 0.75 * txtHgtG / 1000; // 0.185;
                }

                Leader acLdr = new Leader();
                acLdr.AppendVertex(p3dPt1);
                acLdr.AppendVertex(p3dPt2);
                acLdr.AppendVertex(p3dPt3);
                acLdr.HasArrowHead = true;
                acLdr.Dimasz = arrowSize;
                acLdr.ColorIndex = 251;

                BlockTable btLL = (BlockTable)trLL.GetObject(aCurDB.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btrLL = (BlockTableRecord)trLL.GetObject(btLL[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                ObjectId oiLL = btrLL.AppendEntity(mTxt);
                trLL.AddNewlyCreatedDBObject(mTxt, true);

                btrLL.AppendEntity(acLdr);
                trLL.AddNewlyCreatedDBObject(acLdr, true);

                trLL.Commit();
            }
        }

        private static string RebarId()
        {
            string[] alphabet = new string[26];
            alphabet[0] = "A";
            alphabet[1] = "B";
            alphabet[2] = "C";
            alphabet[3] = "D";
            alphabet[4] = "E";
            alphabet[5] = "F";
            alphabet[6] = "G";
            alphabet[7] = "H";
            alphabet[8] = "I";
            alphabet[9] = "J";
            alphabet[10] = "K";
            alphabet[11] = "L";
            alphabet[12] = "M";
            alphabet[13] = "N";
            alphabet[14] = "O";
            alphabet[15] = "P";
            alphabet[16] = "Q";
            alphabet[17] = "R";
            alphabet[18] = "S";
            alphabet[19] = "T";
            alphabet[20] = "U";
            alphabet[21] = "V";
            alphabet[22] = "W";
            alphabet[23] = "X";
            alphabet[24] = "Y";
            alphabet[25] = "Z";

            Random rand = new Random();
            int randNum1 = rand.Next(0, 25);
            int randNum2 = rand.Next(0, 9);
            int randNum3 = rand.Next(0, 25);

            string barId = alphabet[randNum1].ToString() + randNum2.ToString() + alphabet[randNum3].ToString();

            return barId;
        }

        /**
         * Credit to source: https://forums.autodesk.com/t5/net/remove-colinear-vertices-polyline/td-p/10755687
         * Date added: 15 March 2024
         * Added by: Bernardo A. Cabebe Jr.
         * Venue: 3407 Cityland Pasong Tamo Tower, Chino Roces Ave., Makati City
         */
        [CommandMethod("RCV_RemoveColinearVerticesPolyline")]
        public static void RemoveColinearVerticesPolyline()
        {
            if (!ccData.isLicenseActive())
            {

                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            PromptSelectionOptions psoRefObj = new PromptSelectionOptions();
            psoRefObj.MessageForAdding = "\nSelect reference object: ";

            PromptSelectionResult psrRefObj = actDoc.Editor.GetSelection(psoRefObj);
            if (psrRefObj.Status != PromptStatus.OK) return;

            SelectionSet ssRefObj = psrRefObj.Value;

            if (ssRefObj.Count != 1)
            {
                actDoc.Editor.WriteMessage("\nSelect one reference object only.");
                return;
            }

            string layerName;
            int colIndex;
            string objType;
            using (Transaction trRefObj = aCurDB.TransactionManager.StartTransaction())
            {
                Entity enRefObj = trRefObj.GetObject(ssRefObj[0].ObjectId, OpenMode.ForRead) as Entity;
                layerName = enRefObj.Layer;
                colIndex = enRefObj.ColorIndex;
                objType = enRefObj.GetType().Name.ToUpper();

                actDoc.Editor.WriteMessage("\nObject layer name: {0}, color: {1} and type: {2}", layerName, colIndex, objType);

                trRefObj.Commit();
                trRefObj.Dispose();
            }

            TypedValue[] tvFilter = new TypedValue[3];
            if (objType == "POLYLINE")
            {
                tvFilter[0] = new TypedValue((int)DxfCode.Start, "*" + objType + "*");
            }
            else
            {
                actDoc.Editor.WriteMessage("\nInvalid object type: {0}", objType);
                return;
            }

            tvFilter[1] = new TypedValue((int)DxfCode.LayerName, layerName);
            tvFilter[2] = new TypedValue((int)DxfCode.Color, colIndex);

            SelectionFilter sfObjs = new SelectionFilter(tvFilter);
            PromptSelectionResult psrObjs = actDoc.Editor.GetSelection(sfObjs);

            if (psrObjs.Status != PromptStatus.OK) return;

            SelectionSet ssObjs = psrObjs.Value;

            using (Transaction trObjs = actDoc.TransactionManager.StartTransaction())
            {

                foreach (SelectedObject soObjs in ssObjs)
                {
                    Polyline pline = (Polyline)trObjs.GetObject(soObjs.ObjectId, OpenMode.ForRead);

                    List<int> verticesToRemove = new List<int>();

                    for (int i = 0; i < pline.NumberOfVertices - 1; i++)
                    {
                        SegmentType st1 = pline.GetSegmentType(i);
                        SegmentType st2 = pline.GetSegmentType(i + 1);
                        if (st1 == SegmentType.Line && st1 == st2)
                        {
                            LineSegment2d ls2d1 = pline.GetLineSegment2dAt(i);
                            LineSegment2d ls2d2 = pline.GetLineSegment2dAt(i + 1);

                            if (ls2d1.IsColinearTo(ls2d2)) verticesToRemove.Add(i + 1);
                        }
                    }

                    verticesToRemove.Reverse();
                    pline.UpgradeOpen();
                    for (int j = 0; j < verticesToRemove.Count; j++)
                        pline.RemoveVertexAt(verticesToRemove[j]);

                }

                trObjs.Commit();
            }
        }

        /**
         * Credit to source: forums.augi.com/showthread.php?746037-Select-polyline-segment
         */
        [CommandMethod("BBS_BarBendingSchedule")]
        public static void BarBendingSchedule()
        {
            if (!ccData.isLicenseActive())
            {
                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            PromptNestedEntityOptions pneoM3t = new PromptNestedEntityOptions("\nSelect a polyline segment: ");
            pneoM3t.AllowNone = false;
            PromptNestedEntityResult pnerM3t = actDoc.Editor.GetNestedEntity(pneoM3t);
            if (pnerM3t.Status != PromptStatus.OK) return;

            //Enter BBS information
            PromptStringOptions psoBBSInf = new PromptStringOptions("");
            psoBBSInf.DefaultValue = bbsInfoG;
            psoBBSInf.Message = "\nEnter BSS Info: ";
            psoBBSInf.AllowSpaces = false;
            PromptResult prBBSInfo = actDoc.Editor.GetString(psoBBSInf);

            string bbsInfo = prBBSInfo.StringResult;
            bbsInfoG = bbsInfo;
            actDoc.Editor.WriteMessage("\nBBS Information: {0}", bbsInfoG);
            string[] arrBBSInfo = bbsInfo.Split(',');
            if (arrBBSInfo.Length != 3)
            {
                actDoc.Editor.WriteMessage("\nInvalid BBS info.");
                return;
            }

            PromptStringOptions psoGSheetID = new PromptStringOptions("");
            psoGSheetID.DefaultValue = gshtIdG;
            psoGSheetID.Message = "\nEnter Google Spreadsheet ID: ";
            psoGSheetID.AllowSpaces = false;
            PromptResult prGSheetID = actDoc.Editor.GetString(psoGSheetID);

            string spreadSheetID = prGSheetID.StringResult;
            gshtIdG = spreadSheetID;
            actDoc.Editor.WriteMessage("\nSpreadsheet ID: {0}", gshtIdG);


            // If the selected entity is polyline
            if (pnerM3t.ObjectId.ObjectClass.Name != "AcDbPolyline")
            {
                actDoc.Editor.WriteMessage("\nError U9R: Not AcDbPolyline");
                return;
            }

            double pLLen = 0.0;
            double length = 0.0;
            // Start transaction to open the selected polyline
            using (Transaction trM3t = aCurDB.TransactionManager.StartTransaction())
            {
                // Transform the picked point from current UCS to WCS
                Point3d wcsPickedPoint = pnerM3t.PickedPoint.TransformBy(actDoc.Editor.CurrentUserCoordinateSystem);

                // Open the polyline
                Polyline pLine = (Polyline)trM3t.GetObject(pnerM3t.ObjectId, OpenMode.ForRead);

                // Get the length of line
                pLLen = pLine.Length;

                // Get the closest point to picked point on the polyline.
                //If the polyline is nested, it's needed to transform the pick point using the
                //transformation matrix that is applied to the polyline by its containers.
                Point3d pointOnPline = pnerM3t.GetContainers().Length == 0 ?
                    pLine.GetClosestPointTo(wcsPickedPoint, false) : //not nested
                    pLine.GetClosestPointTo(wcsPickedPoint.TransformBy(pnerM3t.Transform.Inverse()), false);//nested polyline

                //Get the selected segment index.
                int segmentIndex = (int)pLine.GetParameterAtPoint(pointOnPline);
                actDoc.Editor.WriteMessage("\nSegment index: {0}", segmentIndex);

                Point3d p3dPt1 = pLine.GetPoint3dAt(segmentIndex);

                int segmentIndex2 = segmentIndex + 1;
                Point3d p3dPt2 = pLine.GetPoint3dAt(segmentIndex2);

                length = p3dPt1.DistanceTo(p3dPt2);
                actDoc.Editor.WriteMessage("\nSegment length: {0}", length);

                trM3t.Commit();

            }

            IDictionary<string, string> idGSCol = new Dictionary<string, string>();
            idGSCol["A"] = "L";
            idGSCol["B"] = "M";
            idGSCol["C"] = "N";
            idGSCol["D"] = "O";
            idGSCol["E"] = "P";
            idGSCol["F"] = "Q";
            idGSCol["G"] = "R";
            idGSCol["H"] = "S";
            idGSCol["I"] = "T";
            idGSCol["J"] = "U";
            idGSCol["K"] = "V";
            idGSCol["L"] = "W";
            idGSCol["M"] = "X";
            idGSCol["N"] = "Y";
            idGSCol["O"] = "Z";
            idGSCol["P"] = "AA";
            idGSCol["Q"] = "AB";

            object[,] arr2ObjM3t = new object[1, 1];
            arr2ObjM3t[0, 0] = length;

            List<IList<object>> lioValuesM3t = new List<IList<object>>();
            for (int aa = 0; aa < arr2ObjM3t.GetLength(0); aa++)
            {
                List<object> loValueM3t = new List<object>();
                for (int bb = 0; bb < arr2ObjM3t.GetLength(1); bb++)
                {
                    loValueM3t.Add(arr2ObjM3t[aa, bb]);
                }
                lioValuesM3t.Add(loValueM3t);
            }

            string[] arrTab = arrBBSInfo[0].Split('=');
            string strTab = arrTab[1];
            string[] arrRow = arrBBSInfo[1].Split('=');
            string strRow = arrRow[1];
            string[] arrCol = arrBBSInfo[2].Split('=');
            string strCol = idGSCol[arrCol[1]];

            string strShtRng = strTab + '!' + strCol + strRow;
            actDoc.Editor.WriteMessage("\nRange: {0}", strShtRng);

            ValueRange vrUpdateCell = new ValueRange();
            vrUpdateCell.MajorDimension = "ROWS";
            vrUpdateCell.Values = lioValuesM3t;
            SpreadsheetsResource.ValuesResource.UpdateRequest urUpdateCell = DataGlobal.sheetsService.Spreadsheets.Values.Update(vrUpdateCell, spreadSheetID, strShtRng);
            urUpdateCell.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            UpdateValuesResponse exeUpdateCell = urUpdateCell.Execute();
            actDoc.Editor.WriteMessage("\nResult: {0}", exeUpdateCell);


            object[,] arr2ObjM4t = new object[1, 1];
            arr2ObjM4t[0, 0] = pLLen / 1000; //Convert to meter as shown in GSheet BBS

            List<IList<object>> lioValuesM4t = new List<IList<object>>();
            for (int cc = 0; cc < arr2ObjM3t.GetLength(0); cc++)
            {
                List<object> loValueM4t = new List<object>();
                for (int dd = 0; dd < arr2ObjM4t.GetLength(1); dd++)
                {
                    loValueM4t.Add(arr2ObjM4t[cc, dd]);
                }
                lioValuesM4t.Add(loValueM4t);
            }

            string strShtRngM4t = strTab + "!H" + strRow;
            actDoc.Editor.WriteMessage("\nRange: {0}", strShtRngM4t);

            ValueRange vrUpdateCellM4t = new ValueRange();
            vrUpdateCellM4t.MajorDimension = "ROWS";
            vrUpdateCellM4t.Values = lioValuesM4t;
            SpreadsheetsResource.ValuesResource.UpdateRequest urUpdateCellM4t = DataGlobal.sheetsService.Spreadsheets.Values.Update(vrUpdateCellM4t, spreadSheetID, strShtRngM4t);
            urUpdateCellM4t.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            UpdateValuesResponse exeUpdateCellM4t = urUpdateCellM4t.Execute();
            actDoc.Editor.WriteMessage("\nResult: {0}", exeUpdateCellM4t);

            return;
        }


        /**
         * Date added: 09 May 2024 2:51pm
         * Added by: Bernardo A. Cabebe Jr.
         * Venue: Converge Reliance
         * Status: Under development
         * BBEE - Bottom Bar End Extend
         */
        [CommandMethod("BBEA_BottomBarEndAdd")]
        public static void BottomBarEndAdd()
        {
            if (!ccData.isLicenseActive())
            {
                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            TypedValue[] tvFilterD0e = new TypedValue[5];
            tvFilterD0e[0] = new TypedValue((int)DxfCode.Start, "*LINE");
            tvFilterD0e[1] = new TypedValue((int)DxfCode.Operator, "<OR");
            tvFilterD0e[2] = new TypedValue((int)DxfCode.Color, botBarCol);
            tvFilterD0e[3] = new TypedValue((int)DxfCode.Color, endLinCol);
            tvFilterD0e[4] = new TypedValue((int)DxfCode.Operator, "OR>");
            SelectionFilter sfBBED0e = new SelectionFilter(tvFilterD0e);

            PromptSelectionOptions psoBBED0e = new PromptSelectionOptions(); //Bottom Beam Line
            psoBBED0e.MessageForAdding = "\nSelect end line and beam line: ";

            PromptSelectionResult psrBBED0e = actDoc.Editor.GetSelection(psoBBED0e, sfBBED0e);
            if (psrBBED0e.Status != PromptStatus.OK) return;

            SelectionSet ssBBED0e = psrBBED0e.Value;
            if (ssBBED0e.Count != 2) return;

            //Get from user the rebar diameter and lap factor
            PromptStringOptions psoBLSD0e = new PromptStringOptions("");
            string defVBLSD0e = barDiaG.ToString() + "," + lapK1G.ToString();
            psoBLSD0e.Message = "\nEnter bar dia and lap factor: ";
            psoBLSD0e.DefaultValue = defVBLSD0e;
            psoBLSD0e.AllowSpaces = false;

            PromptResult prBLSD0e = actDoc.Editor.GetString(psoBLSD0e);
            if (prBLSD0e.Status != PromptStatus.OK) return;
            string strBLSD0e = prBLSD0e.StringResult;
            string[] arrStrD0e = strBLSD0e.Split(new char[] { ',' });
            double barDia = Convert.ToDouble(arrStrD0e[0]);
            double lapK1 = Convert.ToDouble(arrStrD0e[1]);
            actDoc.Editor.WriteMessage("\nBar D: {0}; Lap factor: {1}", barDia, lapK1);
            barDiaG = barDia;
            lapK1G = lapK1;
            double lapLenD03 = barDia * lapK1;

            using (Transaction trBBED0e = actDoc.TransactionManager.StartTransaction())
            {
                Point3d barEndP1 = new Point3d();
                Point3d barEndP2 = new Point3d();
                Curve crLine = null;
                Polyline plBar = null;
                foreach (SelectedObject soBBED0e in ssBBED0e)
                {
                    Curve crBBED0e = trBBED0e.GetObject(soBBED0e.ObjectId, OpenMode.ForRead) as Curve;

                    if (crBBED0e.ColorIndex == botBarCol) //130)
                    {
                        barEndP1 = crBBED0e.StartPoint;
                        barEndP2 = crBBED0e.EndPoint;
                        plBar = trBBED0e.GetObject(soBBED0e.ObjectId, OpenMode.ForRead) as Polyline;
                    }

                    if (crBBED0e.ColorIndex == endLinCol) //54)
                    {
                        crLine = crBBED0e;
                    }
                }

                Point3d closestPt = crLine.GetClosestPointTo(barEndP1, true);
                Point2d closestP2d = new Point2d(closestPt.X, closestPt.Y);
                Point3d nearPt = new Point3d();//Closest end point of bar line to end line
                Point3d farPnt = new Point3d();
                //Get the closest end point of rebar line to end line
                if (closestPt.DistanceTo(barEndP1) < closestPt.DistanceTo(barEndP2))
                {
                    nearPt = barEndP1;
                    farPnt = barEndP2;
                }
                else
                {
                    nearPt = barEndP2;
                    farPnt = barEndP1;
                }

                //Define point 3 of bar line
                Vector3d v3NFPnt = nearPt.GetVectorTo(farPnt);
                Vector3d nv3NFPt = v3NFPnt.GetNormal();
                nv3NFPt = nv3NFPt * lapLenD03;
                Point3d barP3 = nearPt + nv3NFPt;
                Point2d p2BarP3 = new Point2d(barP3.X, barP3.Y);

                //Define point 2 of bar line
                Vector3d v3CNPnt = closestPt.GetVectorTo(nearPt);
                Vector3d nv3CNPt = v3CNPnt.GetNormal();
                double dblCov = 0;
                if (lapLenD03 > 100)
                {
                    dblCov = 40 + 12;
                }
                else
                {
                    dblCov = 0.040 + 0.012;
                }
                nv3CNPt = nv3CNPt * dblCov;
                Point3d barP2 = closestPt + nv3CNPt;
                Point2d p2BarP2 = new Point2d(barP2.X, barP2.Y);

                //Define point 1 of bar line
                Point3d p3BarP0 = barP2 + new Vector3d(0, 100, 0);//100 is applicable whether mm or m units
                Vector3d v3BarP2P0 = barP2.GetVectorTo(p3BarP0);
                Vector3d nv3P2P0 = v3BarP2P0.GetNormal();
                nv3P2P0 = nv3P2P0 * 12 * barDia;
                Point3d p3barP1 = barP2 + nv3P2P0;
                Point2d p2BarP1 = new Point2d(p3barP1.X, p3barP1.Y);


                BlockTable blktbl;
                blktbl = trBBED0e.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;

                BlockTableRecord bltrec;
                bltrec = trBBED0e.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                double lineLen = 0;
                using (Polyline acPoly = new Polyline())
                {
                    acPoly.AddVertexAt(0, p2BarP1, 0, 0, 0);
                    acPoly.AddVertexAt(1, p2BarP2, 0, 0, 0);
                    acPoly.AddVertexAt(2, p2BarP3, 0, 0, 0);

                    acPoly.ColorIndex = botBarCol; //130;
                    lineLen = acPoly.Length;
                    bltrec.AppendEntity(acPoly);
                    trBBED0e.AddNewlyCreatedDBObject(acPoly, true);
                }

                actDoc.Editor.WriteMessage("\nBar length: {0}", lineLen);
                trBBED0e.Commit();
            }

        }

        /**
         * Date added: 10 May 2023 5:4pm
         * Added by: Bernardo A. Cabebe Jr.
         * Venue: 27I Asp Twr
         * Status: Done - 11 May 2024
         */
        [CommandMethod("BBEE_BottomBarEndExtend")]
        public static void BottomBarEndExtend()
        {
            if (!ccData.isLicenseActive())
            {
                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            TypedValue[] tvFilterF9r = new TypedValue[5];
            tvFilterF9r[0] = new TypedValue((int)DxfCode.Start, "*LINE");
            tvFilterF9r[1] = new TypedValue((int)DxfCode.Operator, "<OR");
            tvFilterF9r[2] = new TypedValue((int)DxfCode.Color, botBarCol); //130);
            tvFilterF9r[3] = new TypedValue((int)DxfCode.Color, endLinCol); //54);
            tvFilterF9r[4] = new TypedValue((int)DxfCode.Operator, "OR>");
            SelectionFilter sfBBEF9r = new SelectionFilter(tvFilterF9r);

            PromptSelectionOptions psoBBEF9r = new PromptSelectionOptions(); //Bottom Beam Line
            psoBBEF9r.MessageForAdding = "\nSelect end line and beam line: ";

            PromptSelectionResult psrBBEF9r = actDoc.Editor.GetSelection(psoBBEF9r, sfBBEF9r);
            if (psrBBEF9r.Status != PromptStatus.OK) return;
            SelectionSet ssBBEF9r = psrBBEF9r.Value;


            //Get from user the rebar diameter and lap factor
            PromptStringOptions psoBLSF9r = new PromptStringOptions("");
            string defVBLSF9r = barDiaG.ToString() + "," + lapK1G.ToString();
            psoBLSF9r.Message = "\nEnter bar dia and lap factor: ";
            psoBLSF9r.DefaultValue = defVBLSF9r;
            psoBLSF9r.AllowSpaces = false;

            PromptResult prBLSF9r = actDoc.Editor.GetString(psoBLSF9r);
            if (prBLSF9r.Status != PromptStatus.OK) return;
            string strBLSF9r = prBLSF9r.StringResult;
            string[] arrStrF9r = strBLSF9r.Split(new char[] { ',' });
            double barDia = Convert.ToDouble(arrStrF9r[0]);
            double lapK1 = Convert.ToDouble(arrStrF9r[1]);
            actDoc.Editor.WriteMessage("\nBar D: {0}; Lap factor: {1}", barDia, lapK1);
            barDiaG = barDia;
            lapK1G = lapK1;
            double lapLenF9r = barDia * lapK1;

            using (Transaction trBBEF9r = actDoc.TransactionManager.StartTransaction())
            {
                Point3d barEndP1 = new Point3d();
                Point3d barEndP2 = new Point3d();
                Curve crLine = null;
                Polyline plBar = new Polyline();
                foreach (SelectedObject soBBEF9r in ssBBEF9r)
                {
                    Curve crBBEF9r = trBBEF9r.GetObject(soBBEF9r.ObjectId, OpenMode.ForRead) as Curve;

                    if (crBBEF9r.ColorIndex ==  botBarCol) // 130)
                    {
                        barEndP1 = crBBEF9r.StartPoint;
                        barEndP2 = crBBEF9r.EndPoint;
                        plBar = trBBEF9r.GetObject(soBBEF9r.ObjectId, OpenMode.ForRead) as Polyline;
                    }

                    if (crBBEF9r.ColorIndex == endLinCol) //54)
                    {
                        crLine = crBBEF9r;
                    }
                }

                Point3d closestPt = crLine.GetClosestPointTo(barEndP1, true);
                Point2d closestP2d = new Point2d(closestPt.X, closestPt.Y);
                Point3d nearPt = new Point3d();//Closest end point of bar line to end line
                Point3d farPnt = new Point3d();
                double nearParam = 0;
                int nearVtx = 0;
                //Get the closest end point of rebar line to end line
                if (closestPt.DistanceTo(barEndP1) < closestPt.DistanceTo(barEndP2))
                {
                    nearPt = barEndP1;
                    farPnt = barEndP2;
                    nearParam = plBar.StartParam;
                    nearVtx = (int)nearParam;
                }
                else
                {
                    nearPt = barEndP2;
                    farPnt = barEndP1;
                    nearParam = plBar.EndParam;
                    nearVtx = (int)nearParam;
                }

                //Define point 3 of bar line
                Point3d barP3 = farPnt;
                Point2d p2BarP3 = new Point2d(barP3.X, barP3.Y);

                //Define point 2 of bar line
                Vector3d v3CNPnt = closestPt.GetVectorTo(nearPt);
                Vector3d nv3CNPt = v3CNPnt.GetNormal();
                double dblCov = 0;
                if (lapLenF9r > 100)
                {
                    dblCov = 40 + 12; //mm unit
                }
                else
                {
                    dblCov = 0.040 + 0.012; //m unit
                }
                nv3CNPt = nv3CNPt * dblCov;
                Point3d barP2 = closestPt + nv3CNPt;
                Point2d p2BarP2 = new Point2d(barP2.X, barP2.Y);


                //Define point 1 of bar line
                Point3d p3BarP0 = barP2 + new Vector3d(0, 100, 0);//100 is applicable whether mm or m units
                Vector3d v3BarP2P0 = barP2.GetVectorTo(p3BarP0);
                Vector3d nv3P2P0 = v3BarP2P0.GetNormal();
                nv3P2P0 = nv3P2P0 * 12 * barDia;
                Point3d p3barP1 = barP2 + nv3P2P0;
                Point2d p2BarP1 = new Point2d(p3barP1.X, p3barP1.Y);


                BlockTable blktbl;
                blktbl = trBBEF9r.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;

                BlockTableRecord bltrec;
                bltrec = trBBEF9r.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                double lineLen = 0;
                using (Polyline acPoly = new Polyline())
                {
                    acPoly.AddVertexAt(0, p2BarP1, 0, 0, 0);
                    acPoly.AddVertexAt(1, p2BarP2, 0, 0, 0);
                    acPoly.AddVertexAt(2, p2BarP3, 0, 0, 0);

                    acPoly.ColorIndex = botBarCol; //130;
                    lineLen = acPoly.Length;
                    bltrec.AppendEntity(acPoly);
                    trBBEF9r.AddNewlyCreatedDBObject(acPoly, true);
                }

                plBar.UpgradeOpen();
                plBar.Erase();

                actDoc.Editor.WriteMessage("\nBar length: {0}", lineLen);
                trBBEF9r.Commit();
            }
        }


        /**
         * Date added: 12 May 2024
         * Added by: Bernardo A. Cabebe Jr
         * Venue: 31H One Orchard
         * Status: Under development
         */
        [CommandMethod("DSS_DrawStirrupSimpleSpan")]
        public static void DrawStirrupSimpleSpan()
        {
            if (!ccData.isLicenseActive())
            {
                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            TypedValue[] tvFilterB1q = new TypedValue[2];
            tvFilterB1q[0] = new TypedValue((int)DxfCode.Start, "Text");
            tvFilterB1q[1] = new TypedValue((int)DxfCode.Color, stirTxCol); //3);
            SelectionFilter sfDSSB1q = new SelectionFilter(tvFilterB1q);

            PromptSelectionOptions psoDSSB1q = new PromptSelectionOptions();
            psoDSSB1q.MessageForAdding = "\nSelect stirrup type: ";

            PromptSelectionResult psrDSSB1q = actDoc.Editor.GetSelection(psoDSSB1q, sfDSSB1q);
            if (psrDSSB1q.Status != PromptStatus.OK) return;

            SelectionSet ssDSSB1q = psrDSSB1q.Value;
            if (ssDSSB1q.Count > 1)
            {
                actDoc.Editor.WriteMessage("\nSelect only one stirrup type: ");
                return;
            }

            TypedValue[] tvFilterB2q = new TypedValue[2];
            tvFilterB2q[0] = new TypedValue((int)DxfCode.Start, "LINE");
            tvFilterB2q[1] = new TypedValue((int)(DxfCode.Color), botLinCol); //53);
            SelectionFilter sfDSSB2q = new SelectionFilter(tvFilterB2q);

            PromptSelectionOptions psoDSSB2q = new PromptSelectionOptions();
            psoDSSB2q.MessageForAdding = "\nSelect bottom bar line: ";

            PromptSelectionResult psrDSSB2q = actDoc.Editor.GetSelection(psoDSSB2q, sfDSSB2q);
            if (psrDSSB2q.Status != PromptStatus.OK) return;

            SelectionSet ssDSSB2q = psrDSSB2q.Value;
            if (ssDSSB2q.Count != 1)
            {
                actDoc.Editor.WriteMessage("\nInvalid number of selection. Exit.");
                return;
            }


            PromptDoubleOptions pdoDSSB3q = new PromptDoubleOptions("\nEnter beam drawing height: ");
            pdoDSSB3q.DefaultValue = beamHG;
            pdoDSSB3q.AllowZero = false;
            pdoDSSB3q.AllowNone = true;
            pdoDSSB3q.AllowNegative = false;

            PromptDoubleResult pdrDSSB3q = actDoc.Editor.GetDouble(pdoDSSB3q);
            if (pdrDSSB3q.Status != PromptStatus.OK) return;

            double bDwgHgt = pdrDSSB3q.Value; //beam drawing height
            beamHG = bDwgHgt;

            using (Transaction trDSSB1q = aCurDB.TransactionManager.StartTransaction())
            {
                DBText txDSSB1q = trDSSB1q.GetObject(ssDSSB1q[0].ObjectId, OpenMode.ForRead) as DBText;
                string txcDSSB1q = txDSSB1q.TextString;

                string patt1 = @"\d+(?=D)";
                Regex regexDb = new Regex(patt1);
                MatchCollection matchDb = regexDb.Matches(txcDSSB1q);
                actDoc.Editor.WriteMessage("\nStirrrup diameter: {0}", matchDb[0].Value);

                string patt2 = @"\d+(?=\@)";
                Regex regexQb = new Regex(patt2);
                MatchCollection matchQb = regexQb.Matches(txcDSSB1q);
                actDoc.Editor.WriteMessage("\n");
                for (int aa = 0; aa < matchQb.Count; aa++)
                {
                    actDoc.Editor.WriteMessage("N={0} ", matchQb[aa].Value);
                }

                string patt3 = @"(?<=\@)\d+";
                Regex regexSb = new Regex(patt3);
                MatchCollection matchSb = regexSb.Matches(txcDSSB1q);
                actDoc.Editor.WriteMessage("\n");
                for (int bb = 0; bb < matchSb.Count; bb++)
                {
                    actDoc.Editor.WriteMessage("S={0} ", matchSb[bb].Value);
                }

                string patt4 = @"\d+(?=\sO\.C\.)";
                Regex regexRe = new Regex(patt4);
                Match matchRe = regexRe.Match(txcDSSB1q);
                string strRe = "";
                if (matchRe.Success)
                {
                    strRe = matchRe.Value;
                    actDoc.Editor.WriteMessage("\nRest S-{0}", strRe);
                }
                else
                {
                    actDoc.Editor.WriteMessage("\nInvalid stirrup pattern. Exit...");
                    return;
                }

                Line lnDSSB2q = trDSSB1q.GetObject(ssDSSB2q[0].ObjectId, OpenMode.ForRead) as Line;
                double lnLen = lnDSSB2q.Length;
                Point3d p3EndP1 = lnDSSB2q.StartPoint;
                Point3d p3EndP2 = lnDSSB2q.EndPoint;
                Point3d p3MidPt = Midpoint(p3EndP1, p3EndP2);

                double offDis = 0;
                if (lnLen > 100)
                {
                    offDis = 52;
                }
                else
                {
                    offDis = 0.052;
                }
                Point3d p3StPt1 = p3EndP1 + new Vector3d(0, offDis, 0);
                Point3d p3StPt2 = p3StPt1 + new Vector3d(0, bDwgHgt - 2 * offDis, 0);
                Point3d p3S2Pt1 = p3EndP2 + new Vector3d(0, offDis, 0);
                Point3d p3S2Pt2 = p3S2Pt1 + new Vector3d(0, bDwgHgt - 2 * offDis, 0);
                Point3d p3StMP1 = p3MidPt + new Vector3d(0, offDis, 0);
                Point3d p3StMP2 = p3StMP1 + new Vector3d(0, bDwgHgt - 2 * offDis, 0);

                BlockTable blktbl;
                blktbl = trDSSB1q.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;

                BlockTableRecord bltrec;
                bltrec = trDSSB1q.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                Vector3d v3P1MP = p3EndP1.GetVectorTo(p3MidPt);
                Vector3d nv3P1M = v3P1MP.GetNormal();
                Vector3d v3P2MP = p3EndP2.GetVectorTo(p3MidPt);
                Vector3d nv3P2M = v3P2MP.GetNormal();

                double dblX = 0;
                for (int cc = 0; cc < matchQb.Count; cc++)
                {
                    for (int dd = 1; dd <= Convert.ToInt16(matchQb[cc].Value); dd++)
                    {
                        dblX += Convert.ToDouble(matchSb[cc].Value);
                        if (dblX <= (0.5 * lnLen))
                        {
                            p3StPt1 = p3StPt1 + nv3P1M * Convert.ToDouble(matchSb[cc].Value);
                            p3StPt2 = p3StPt2 + nv3P1M * Convert.ToDouble(matchSb[cc].Value);
                            using (Line lnStirr = new Line(p3StPt1, p3StPt2))
                            {
                                lnStirr.ColorIndex = stirLnCol;
                                bltrec.AppendEntity(lnStirr);
                                trDSSB1q.AddNewlyCreatedDBObject(lnStirr, true);
                            }

                            p3S2Pt1 = p3S2Pt1 + nv3P2M * Convert.ToDouble(matchSb[cc].Value);
                            p3S2Pt2 = p3S2Pt2 + nv3P2M * Convert.ToDouble(matchSb[cc].Value);
                            using (Line lnStir2 = new Line(p3S2Pt1, p3S2Pt2))
                            {
                                lnStir2.ColorIndex = stirLnCol;
                                bltrec.AppendEntity(lnStir2);
                                trDSSB1q.AddNewlyCreatedDBObject(lnStir2, true);
                            }
                        }
                    }
                }

                dblX += Convert.ToDouble(strRe);
                while (dblX <= (0.5 * lnLen))
                {
                    p3StPt1 = p3StPt1 + nv3P1M * Convert.ToDouble(strRe);
                    p3StPt2 = p3StPt2 + nv3P1M * Convert.ToDouble(strRe);
                    using (Line lnStirr = new Line(p3StPt1, p3StPt2))
                    {
                        lnStirr.ColorIndex = stirLnCol;
                        bltrec.AppendEntity(lnStirr);
                        trDSSB1q.AddNewlyCreatedDBObject(lnStirr, true);
                    }

                    p3S2Pt1 = p3S2Pt1 + nv3P2M * Convert.ToDouble(strRe);
                    p3S2Pt2 = p3S2Pt2 + nv3P2M * Convert.ToDouble(strRe);
                    using (Line lnStir2 = new Line(p3S2Pt1, p3S2Pt2))
                    {
                        lnStir2.ColorIndex = stirLnCol;
                        bltrec.AppendEntity(lnStir2);
                        trDSSB1q.AddNewlyCreatedDBObject(lnStir2, true);
                    }

                    dblX += Convert.ToDouble(strRe);
                }

                dblX -= Convert.ToDouble(strRe);
                actDoc.Editor.WriteMessage("\nX: {0}", dblX);

                double dblRe = Convert.ToDouble(strRe);
                if ((lnLen - (2 * dblX)) > dblRe)
                {
                    using (Line lnStir2 = new Line(p3StMP1, p3StMP2))
                    {
                        lnStir2.ColorIndex = stirLnCol;
                        bltrec.AppendEntity(lnStir2);
                        trDSSB1q.AddNewlyCreatedDBObject(lnStir2, true);
                    }
                }

                trDSSB1q.Commit();
            }

        }

        [CommandMethod("TBE_TopBarExtra")]
        public static void TopBarExtra()
        {
            if (!ccData.isLicenseActive())
            {
                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            TypedValue[] tvFilterS0p = new TypedValue[5];
            tvFilterS0p[0] = new TypedValue((int)DxfCode.Start, "*LINE");
            tvFilterS0p[1] = new TypedValue((int)DxfCode.Operator, "<OR");
            tvFilterS0p[2] = new TypedValue((int)DxfCode.Color, topLinCol); //52);
            tvFilterS0p[3] = new TypedValue((int)DxfCode.Color, endLinCol); //54);
            tvFilterS0p[4] = new TypedValue((int)DxfCode.Operator, "OR>");
            SelectionFilter sfTBES0p = new SelectionFilter(tvFilterS0p);

            PromptSelectionOptions psoTBES0p = new PromptSelectionOptions(); //Bottom Beam Line
            psoTBES0p.MessageForAdding = "\nSelect beam lines and/or end line: ";

            PromptSelectionResult psrTBES0p = actDoc.Editor.GetSelection(psoTBES0p, sfTBES0p);
            if (psrTBES0p.Status != PromptStatus.OK) return;

            SelectionSet ssTBES0p = psrTBES0p.Value;
            if (ssTBES0p.Count != 2) return;

            using (Transaction trTBES0p = aCurDB.TransactionManager.StartTransaction())
            {
                Curve cuObj1 = trTBES0p.GetObject(ssTBES0p[0].ObjectId, OpenMode.ForRead) as Curve;

                Curve cuObj2 = trTBES0p.GetObject(ssTBES0p[1].ObjectId, OpenMode.ForRead) as Curve;

                Point3d p3NrstPO1 = new Point3d();
                Vector3d nv3Obj1 = new Vector3d();
                if (cuObj2.StartPoint.DistanceTo(cuObj1.StartPoint) < cuObj2.StartPoint.DistanceTo(cuObj1.EndPoint))
                {
                    p3NrstPO1 = cuObj1.StartPoint;
                    nv3Obj1 = p3NrstPO1.GetVectorTo(cuObj1.EndPoint);
                    nv3Obj1 = nv3Obj1.GetNormal();
                }
                else
                {
                    p3NrstPO1 = cuObj1.EndPoint;
                    nv3Obj1 = p3NrstPO1.GetVectorTo(cuObj1.StartPoint);
                    nv3Obj1 = nv3Obj1.GetNormal();
                }

                Point3d p3NrstPO2 = new Point3d();
                Vector3d nv3Obj2 = new Vector3d();
                if (p3NrstPO1.DistanceTo(cuObj2.StartPoint) < p3NrstPO1.DistanceTo(cuObj2.EndPoint))
                {
                    p3NrstPO2 = cuObj2.StartPoint;
                    nv3Obj2 = p3NrstPO2.GetVectorTo(cuObj2.EndPoint);
                    nv3Obj2 = nv3Obj2.GetNormal();
                }
                else
                {
                    p3NrstPO2 = cuObj2.EndPoint;
                    nv3Obj2 = p3NrstPO2.GetVectorTo(cuObj2.StartPoint);
                    nv3Obj2 = nv3Obj2.GetNormal();
                }

                double lenObj1 = cuObj1.GetDistanceAtParameter(cuObj1.GetParameterAtPoint(cuObj1.EndPoint)) -
                                 cuObj1.GetDistanceAtParameter(cuObj1.GetParameterAtPoint(cuObj1.StartPoint));
                double ob1L3rd = lenObj1 / 3;
                nv3Obj1 = nv3Obj1 * ob1L3rd;

                double lenObj2 = cuObj2.GetDistanceAtParameter(cuObj2.GetParameterAtPoint(cuObj2.EndPoint)) -
                                 cuObj2.GetDistanceAtParameter(cuObj2.GetParameterAtPoint(cuObj2.StartPoint));
                double ob2L3rd = lenObj2 / 3;
                nv3Obj2 = nv3Obj2 * ob2L3rd;

                double offDisY = beamHG / 4;

                Point3d plnEnd1 = p3NrstPO1 + nv3Obj1;
                Point2d pln2dE1 = new Point2d(plnEnd1.X, plnEnd1.Y - offDisY);
                Point3d plnEnd2 = p3NrstPO2 + nv3Obj2;
                Point2d pln2dE2 = new Point2d(plnEnd2.X, plnEnd2.Y - offDisY);

                BlockTable blktbl;
                blktbl = trTBES0p.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;

                BlockTableRecord bltrec;
                bltrec = trTBES0p.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                double lineLen = 0;
                using (Polyline acPoly = new Polyline())
                {
                    acPoly.AddVertexAt(0, pln2dE1, 0, 0, 0);
                    acPoly.AddVertexAt(1, pln2dE2, 0, 0, 0);

                    acPoly.ColorIndex = topExtCol; //120;
                    lineLen = acPoly.Length;
                    bltrec.AppendEntity(acPoly);
                    trTBES0p.AddNewlyCreatedDBObject(acPoly, true);
                }

                actDoc.Editor.WriteMessage("\nRebar length: {0}", lineLen);

                trTBES0p.Commit();
            }
        }


        /**
         * Date added: 17 May 2024
         * Added by: Bernardo A. Cabebe Jr.
         * Venue: 3407 Cityland Pasong Tamo Tower
         * Status: Under development
         */
        [CommandMethod("BBE_BottomBarExtra")]
        public static void BottomBarExtra()
        {
            if (!ccData.isLicenseActive())
            {
                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            TypedValue[] tvFilterG5t = new TypedValue[2];
            tvFilterG5t[0] = new TypedValue((int)DxfCode.Start, "*LINE");
            tvFilterG5t[1] = new TypedValue((int)DxfCode.Color, botLinCol); //53);
            SelectionFilter sfBBEG5t = new SelectionFilter(tvFilterG5t);

            PromptSelectionOptions psoBBEG5t = new PromptSelectionOptions(); //Bottom Beam Line
            psoBBEG5t.MessageForAdding = "\nSelect bootm beam lines: ";

            PromptSelectionResult psrBBEG5t = actDoc.Editor.GetSelection(psoBBEG5t, sfBBEG5t);
            if (psrBBEG5t.Status != PromptStatus.OK) return;

            SelectionSet ssBBEG5t = psrBBEG5t.Value;
            if (ssBBEG5t.Count < 1) return;

            using (Transaction trBBEG5t = aCurDB.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject soBBEG5t in ssBBEG5t)
                {
                    Curve cObj = trBBEG5t.GetObject(soBBEG5t.ObjectId, OpenMode.ForRead) as Curve;
                    Point3d cObjP1 = cObj.StartPoint;
                    Point3d cObjP2 = cObj.EndPoint;
                    Point3d cObjMP = Midpoint(cObjP1, cObjP2);

                    double cuLen = cObj.GetDistanceAtParameter(cObj.GetParameterAtPoint(cObj.EndPoint)) -
                                   cObj.GetDistanceAtParameter(cObj.GetParameterAtPoint(cObj.StartPoint));

                    double cuDis = cuLen / 7;

                    Vector3d v3P1mPt = cObjP1.GetVectorTo(cObjMP);
                    Vector3d nv3P1mP = v3P1mPt.GetNormal();
                    nv3P1mP = nv3P1mP * cuDis;

                    Vector3d v3P2mPt = cObjP2.GetVectorTo(cObjMP);
                    Vector3d nv3P2mP = v3P2mPt.GetNormal();
                    nv3P2mP = nv3P2mP * cuDis;

                    double offDisY = beamHG / 4;
                    Point3d p3ExBe1 = cObjP1 + nv3P1mP;
                    Point2d p2ExBe1 = new Point2d(p3ExBe1.X, p3ExBe1.Y + offDisY);

                    Point3d p3ExBe2 = cObjP2 + nv3P2mP;
                    Point2d p2ExBe2 = new Point2d(p3ExBe2.X, p3ExBe2.Y + offDisY);

                    BlockTable blktbl;
                    blktbl = trBBEG5t.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;

                    BlockTableRecord bltrec;
                    bltrec = trBBEG5t.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    double lineLen = 0;
                    using (Polyline acPoly = new Polyline())
                    {
                        acPoly.AddVertexAt(0, p2ExBe1, 0, 0, 0);
                        acPoly.AddVertexAt(1, p2ExBe2, 0, 0, 0);

                        acPoly.ColorIndex = botExtCol; //140;
                        lineLen = acPoly.Length;
                        bltrec.AppendEntity(acPoly);
                        trBBEG5t.AddNewlyCreatedDBObject(acPoly, true);
                    }

                    actDoc.Editor.WriteMessage("\nRebar length: {0}", lineLen);


                }

                trBBEG5t.Commit();
            }

        }


        /**
         * Date added: 17 May 2024
         * Added by: Bernardo A. Cabebe Jr
         * Venue: 3407 Cityland Pasong Tamo Tower
         * Status: Under development
         */
        [CommandMethod("DSC_DrawStirrupCantileverSpan")]
        public static void DrawStirrupCantileverSpan()
        {
            if (!ccData.isLicenseActive())
            {
                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            TypedValue[] tvFilterU4s = new TypedValue[2];
            tvFilterU4s[0] = new TypedValue((int)DxfCode.Start, "Text");
            tvFilterU4s[1] = new TypedValue((int)DxfCode.Color, stirTxCol); //3);
            SelectionFilter sfDSCU4s = new SelectionFilter(tvFilterU4s);

            PromptSelectionOptions psoDSCU4s = new PromptSelectionOptions();
            psoDSCU4s.MessageForAdding = "\nSelect stirrup type: ";

            PromptSelectionResult psrDSCU4s = actDoc.Editor.GetSelection(psoDSCU4s, sfDSCU4s);
            if (psrDSCU4s.Status != PromptStatus.OK) return;

            SelectionSet ssDSCU4s = psrDSCU4s.Value;
            if (ssDSCU4s.Count > 1)
            {
                actDoc.Editor.WriteMessage("\nSelect only one stirrup type: ");
                return;
            }

            TypedValue[] tvFilterU3s = new TypedValue[5];
            tvFilterU3s[0] = new TypedValue((int)DxfCode.Start, "*LINE");
            tvFilterU3s[1] = new TypedValue((int)DxfCode.Operator, "<OR");
            tvFilterU3s[2] = new TypedValue((int)(DxfCode.Color), botLinCol); //53);
            tvFilterU3s[3] = new TypedValue((int)(DxfCode.Color), endLinCol); //54);
            tvFilterU3s[4] = new TypedValue((int)DxfCode.Operator, "OR>");
            SelectionFilter sfDSCU3s = new SelectionFilter(tvFilterU3s);

            PromptSelectionOptions psoDSCU3s = new PromptSelectionOptions();
            psoDSCU3s.MessageForAdding = "\nSelect bottom bar line and end lines: ";

            PromptSelectionResult psrDSCU3s = actDoc.Editor.GetSelection(psoDSCU3s, sfDSCU3s);
            if (psrDSCU3s.Status != PromptStatus.OK) return;

            SelectionSet ssDSCU3s = psrDSCU3s.Value;
            if (ssDSCU3s.Count != 2)
            {
                actDoc.Editor.WriteMessage("\nInvalid number of selection. Exit.");
                return;
            }


            PromptDoubleOptions pdoDSCU3s = new PromptDoubleOptions("\nEnter beam drawing height: ");
            pdoDSCU3s.DefaultValue = beamHG;
            pdoDSCU3s.AllowZero = false;
            pdoDSCU3s.AllowNone = true;
            pdoDSCU3s.AllowNegative = false;

            PromptDoubleResult pdrDSCU3s = actDoc.Editor.GetDouble(pdoDSCU3s);
            if (pdrDSCU3s.Status != PromptStatus.OK) return;

            double bDwgHgt = pdrDSCU3s.Value; //beam drawing height
            beamHG = bDwgHgt;

            using(Transaction trDSCU4s = aCurDB.TransactionManager.StartTransaction())
            {
                DBText txDSSB1q = trDSCU4s.GetObject(ssDSCU4s[0].ObjectId, OpenMode.ForRead) as DBText;
                string txcDSSB1q = txDSSB1q.TextString;

                string patt1 = @"\d+(?=D)";//Diameter
                Regex regexDb = new Regex(patt1);
                MatchCollection matchDb = regexDb.Matches(txcDSSB1q);
                actDoc.Editor.WriteMessage("\nStirrrup diameter: {0}", matchDb[0].Value);

                string patt2 = @"\d+(?=\@)";//Quantity
                Regex regexQb = new Regex(patt2);
                MatchCollection matchQb = regexQb.Matches(txcDSSB1q);
                actDoc.Editor.WriteMessage("\n");
                for (int aa = 0; aa < matchQb.Count; aa++)
                {
                    actDoc.Editor.WriteMessage("N={0} ", matchQb[aa].Value);
                }

                string patt3 = @"(?<=\@)\d+";//Spacing
                Regex regexSb = new Regex(patt3);
                MatchCollection matchSb = regexSb.Matches(txcDSSB1q);
                actDoc.Editor.WriteMessage("\n");
                for (int bb = 0; bb < matchSb.Count; bb++)
                {
                    actDoc.Editor.WriteMessage("S={0} ", matchSb[bb].Value);
                }

                string patt4 = @"\d+(?=\sO\.C\.)";//Rest spacing
                Regex regexRe = new Regex(patt4);
                Match matchRe = regexRe.Match(txcDSSB1q);
                string strRe = "";
                if (matchRe.Success)
                {
                    strRe = matchRe.Value;
                    actDoc.Editor.WriteMessage("\nRest S-{0}", strRe);
                }
                else
                {
                    actDoc.Editor.WriteMessage("\nInvalid stirrup pattern. Exit...");
                    return;
                }

                
                Line lbotLn = new Line();
                Line lendLn = new Line();
                foreach (SelectedObject soDSCU4s in ssDSCU3s)
                {
                    Line lniDSCU4s = trDSCU4s.GetObject(soDSCU4s.ObjectId, OpenMode.ForRead) as Line;
                    if (lniDSCU4s.ColorIndex == botLinCol) //53)
                    {
                        lbotLn = lniDSCU4s;
                    }
                    if (lniDSCU4s.ColorIndex == endLinCol) //54)
                    {
                        lendLn = lniDSCU4s;
                    }
                }

                Point3d p3BotLnP1 = lbotLn.StartPoint;
                Point3d p3BotLnP2 = lbotLn.EndPoint;
                Point3d p3EndLnPt = lendLn.StartPoint;// Any point, eithter start point or end point

                Point3d p3BotLnSP = new Point3d();// Start point
                Point3d p3BotLnEP = new Point3d();// End point
                if(p3EndLnPt.DistanceTo(p3BotLnP1) > p3EndLnPt.DistanceTo(p3BotLnP2))
                {
                    p3BotLnSP = p3BotLnP1;
                    p3BotLnEP = p3BotLnP2;
                }
                else
                {
                    p3BotLnSP = p3BotLnP2;
                    p3BotLnEP = p3BotLnP1;
                }

                double lnLen = lbotLn.Length;
                double offDis = 0;
                if (lnLen > 100)
                {
                    offDis = 52;
                }
                else
                {
                    offDis = 0.052;
                }
                Vector3d v3SEPt = p3BotLnSP.GetVectorTo(p3BotLnEP); //Vector from start point to end point
                Vector3d nv3SEP = v3SEPt.GetNormal();
                Point3d p3StPt1 = p3BotLnSP + new Vector3d(0, offDis, 0);
                Point3d p3StPt2 = p3StPt1 + new Vector3d(0, bDwgHgt - 2 * offDis, 0);

                BlockTable blktbl;
                blktbl = trDSCU4s.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;

                BlockTableRecord bltrec;
                bltrec = trDSCU4s.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                double dblX = 0;
                for (int cc = 0; cc < matchQb.Count; cc++)
                {
                    for (int dd = 1; dd <= Convert.ToInt16(matchQb[cc].Value); dd++)
                    {
                        dblX += Convert.ToDouble(matchSb[cc].Value);
                        if (dblX <= lnLen)
                        {
                            p3StPt1 = p3StPt1 + nv3SEP * Convert.ToDouble(matchSb[cc].Value);
                            p3StPt2 = p3StPt2 + nv3SEP * Convert.ToDouble(matchSb[cc].Value);
                            using (Line lnStirr = new Line(p3StPt1, p3StPt2))
                            {
                                lnStirr.ColorIndex = stirLnCol; //251;
                                bltrec.AppendEntity(lnStirr);
                                trDSCU4s.AddNewlyCreatedDBObject(lnStirr, true);
                            }
                        }
                    }
                }

                dblX += Convert.ToDouble(strRe);
                while (dblX <= lnLen)
                {
                    p3StPt1 = p3StPt1 + nv3SEP * Convert.ToDouble(strRe);
                    p3StPt2 = p3StPt2 + nv3SEP * Convert.ToDouble(strRe);
                    using (Line lnStirr = new Line(p3StPt1, p3StPt2))
                    {
                        lnStirr.ColorIndex = stirLnCol; //251;
                        bltrec.AppendEntity(lnStirr);
                        trDSCU4s.AddNewlyCreatedDBObject(lnStirr, true);
                    }

                    dblX += Convert.ToDouble(strRe);
                }

                trDSCU4s.Commit();
            }
        }


        /**
         * Date added: 18 May 2024
         * Added by: Bernardo A. Cabebe Jr.
         * Venue: 304 Portofino, East Ortigas Mansions
         * Status: Under development
         */
        [CommandMethod("WB_WebBars")]
        public static void WebBars()
        {
            if (!ccData.isLicenseActive())
            {
                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            TypedValue[] tvFilterD9i = new TypedValue[2];
            tvFilterD9i[0] = new TypedValue((int)DxfCode.Start, "LINE");
            tvFilterD9i[1] = new TypedValue((int)DxfCode.Color, endLinCol); //54);
            SelectionFilter sfWBD9i = new SelectionFilter(tvFilterD9i);

            PromptSelectionOptions psoWBD9i = new PromptSelectionOptions();
            psoWBD9i.MessageForAdding = "\nSelect end lines: ";

            PromptSelectionResult psrWBD9i = actDoc.Editor.GetSelection(psoWBD9i, sfWBD9i);
            if (psrWBD9i.Status != PromptStatus.OK) return;

            SelectionSet ssWBD9i = psrWBD9i.Value;
            if (ssWBD9i.Count > 2) return;

            PromptDoubleOptions pdoWBD9i = new PromptDoubleOptions("");
            pdoWBD9i.DefaultValue = webDiaG;
            pdoWBD9i.Message = "\nEnter web bar diameter: ";
            pdoWBD9i.AllowNegative = false;
            pdoWBD9i.AllowZero = false;

            PromptDoubleResult pdrWBD9i = actDoc.Editor.GetDouble(pdoWBD9i);
            if (pdrWBD9i.Status != PromptStatus.OK) return;
            double wbarDia = pdrWBD9i.Value;
            webDiaG = wbarDia;

            using (Transaction trWBD9i = aCurDB.TransactionManager.StartTransaction())
            {
                Curve cuObj1 = trWBD9i.GetObject(ssWBD9i[0].ObjectId, OpenMode.ForRead) as Curve;
                Curve cuObj2 = trWBD9i.GetObject(ssWBD9i[1].ObjectId, OpenMode.ForRead) as Curve;

                Point3d p3Obj1P1 = cuObj1.StartPoint;
                Point3d p3Obj2P2 = cuObj1.EndPoint;
                Point3d p3WBPnt1 = Midpoint(p3Obj1P1, p3Obj2P2);

                Point3d p3WBPnt2 = cuObj2.GetClosestPointTo(p3WBPnt1,true);

                Vector3d v3WBP12 = p3WBPnt1.GetVectorTo(p3WBPnt2);
                Vector3d nv3WB12 = v3WBP12.GetNormal();


                Vector3d v3WBP21 = p3WBPnt2.GetVectorTo(p3WBPnt1);
                Vector3d nv3WB21 = v3WBP21.GetNormal();

                double lapLen = 40 * wbarDia;
                double spnLen = p3WBPnt1.DistanceTo(p3WBPnt2);

                double stdBarL = 0;
                double offDisX = 0;
                if(spnLen > 100)
                {
                    stdBarL = 6000;
                    offDisX = 40;
                }
                else
                {
                    stdBarL = 6.0;
                    offDisX = 0.040;
                }

                Vector3d nv3OffXi = offDisX * nv3WB12;
                Vector3d nv3OffXj = offDisX * nv3WB21;
                Vector3d nv3stdBL = stdBarL * nv3WB12;
                Vector3d nv3wbLap = lapLen * nv3WB21;
                Point3d p3VarP1 = p3WBPnt1 + nv3OffXi;
                Point3d p3VarP2 = p3VarP1 + nv3stdBL;

                if (p3WBPnt1.DistanceTo(p3VarP2) > spnLen)
                {
                    p3VarP2 = p3WBPnt2 + nv3OffXj;
                }
                BlockTable blktbl;
                blktbl = trWBD9i.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord bltrec;
                bltrec = trWBD9i.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                double countf = spnLen / (stdBarL - lapLen);
                int count = Convert.ToInt16(countf);
                for (int aa = 0; aa <= count; aa++)
                {
                    Point2d p2VarP1 = new Point2d(p3VarP1.X, p3VarP1.Y);
                    Point2d p2VarP2 = new Point2d(p3VarP2.X, p3VarP2.Y);
                    using(Polyline plObj = new Polyline())
                    {
                        plObj.AddVertexAt(0, p2VarP1, 0, 0, 0);
                        plObj.AddVertexAt(1, p2VarP2, 0, 0, 0);
                        plObj.ColorIndex = webLinCol;
                        bltrec.AppendEntity(plObj);
                        trWBD9i.AddNewlyCreatedDBObject(plObj, true);
                    }

                    //using (Line lnWeb = new Line(p3VarP1, p3VarP2))
                    //{
                    //    lnWeb.ColorIndex = webLinCol;
                    //    bltrec.AppendEntity(lnWeb);
                    //    trWBD9i.AddNewlyCreatedDBObject(lnWeb, true);
                    //}

                    p3VarP1 = p3VarP2 + nv3wbLap;
                    p3VarP2 = p3VarP1 + nv3stdBL;
                    if (p3WBPnt1.DistanceTo(p3VarP2) > spnLen)
                    {
                        p3VarP2 = p3WBPnt2 + nv3OffXj;
                    }
                }

                trWBD9i.Commit();
            }
             
        }


        /**
         * Date added: 19 May 2024
         * Added by: Bernardo A. Cabebe Jr.
         * Venue: 304 Potofino East Ortigas Mansions
         * Status: Under development
         */
        [CommandMethod("SS2_StirrupSection2")]
        public static void StirrupSection2()
        {
            if (!ccData.isLicenseActive())
            {
                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            //Get from user the rebar diameter and lap factor
            PromptStringOptions psoSSV9i = new PromptStringOptions("");
            string defSSV9i = "DB="+barDiaG.ToString() + ";DS=" + stiDiaG.ToString() + ";BX=" + barXDiG.ToString() + ";BY=" + barYDiG.ToString();
            psoSSV9i.Message = "\nEnter values of DB, DS, BX, BY,: ";
            psoSSV9i.DefaultValue = defSSV9i;
            psoSSV9i.AllowSpaces = false;

            PromptResult prSSV9i = actDoc.Editor.GetString(psoSSV9i);
            if (prSSV9i.Status != PromptStatus.OK) return;

            string strResSSV9i = prSSV9i.StringResult;
            actDoc.Editor.WriteMessage("\nInput: {0}", strResSSV9i);

            string patt1 = @"(?<=DB\=)\d+";
            Regex regexDb = new Regex(patt1);
            MatchCollection matchDb = regexDb.Matches(strResSSV9i);
            double barDia = Convert.ToDouble(matchDb[0].Value);
            actDoc.Editor.WriteMessage("\nMain bar diameter: {0}", barDia);
            barDiaG = barDia;

            string patt2 = @"(?<=DS\=)\d+";
            Regex regexDs = new Regex(patt2);
            MatchCollection matchDs = regexDs.Matches(strResSSV9i);
            double stiDia = Convert.ToDouble(matchDs[0].Value);
            actDoc.Editor.WriteMessage("\nStirrup bar diameter: {0}", stiDia);
            stiDiaG = stiDia;

            string patt3 = @"(?<=BX\=)\d+";
            Regex regexBX = new Regex(patt3);
            MatchCollection matchBX = regexBX.Matches(strResSSV9i);
            double disBX = Convert.ToDouble(matchBX[0].Value);
            actDoc.Editor.WriteMessage("\nDistance X: {0}", disBX);
            barXDiG = disBX;

            string patt4 = @"(?<=BY\=)\d+";
            Regex regexBY = new Regex(patt4);
            MatchCollection matchBY = regexBY.Matches(strResSSV9i);
            double disBY = Convert.ToDouble(matchBY[0].Value);
            actDoc.Editor.WriteMessage("\nDistance Y: {0}", disBY);
            barYDiG = disBY;

            PromptPointOptions ppoSSV9i = new PromptPointOptions("\nEnter point: ");
            PromptPointResult pprSSV9i = actDoc.Editor.GetPoint(ppoSSV9i);

            Point3d p3Orig = pprSSV9i.Value;

            double theta = 45 * Math.PI/180; //radians
            double rSSV9i = 0.5 * barDia + 0.5 * stiDia;
            double cDis = 6 * stiDia;

            double x2 = p3Orig.X + rSSV9i * Math.Cos(theta);
            double y2 = p3Orig.Y + rSSV9i * Math.Sin(theta);
            Point2d p2p2 = new Point2d(x2, y2);
            double p2 = 0.1989;

            double x1 = cDis * Math.Sin(theta) + x2;
            double y1 = y2 - cDis*Math.Cos(theta);
            Point2d p2p1 = new Point2d(x1, y1);
            double p1 = 0;

            double x3 = p3Orig.X;
            double y3 = p3Orig.Y + rSSV9i;
            Point2d p2p3 = new Point2d(x3, y3);
            double p3 = 0.4142;

            double x4 = p3Orig.X - rSSV9i;
            double y4 = p3Orig.Y;
            Point2d p2p4 = new Point2d(x4, y4);
            double p4 = 0;

            double x5 = x4;
            double y5 = y4 - disBY;
            Point2d p2p5 = new Point2d(x5, y5);
            double p5 = 0.4142;

            double x6 = x5 + rSSV9i;
            double y6 = y5 - rSSV9i;
            Point2d p2p6 = new Point2d(x6, y6);
            double p6 = 0;

            double x7 = x6 + disBX;
            double y7 = y6;
            Point2d p2p7 = new Point2d(x7, y7);
            double p7 = 0.4142;

            double x8 = x7 + rSSV9i;
            double y8 = y7 + rSSV9i;
            Point2d p2p8 = new Point2d(x8, y8);
            double p8 = 0;

            double x9 = x8;
            double y9 = y8 + disBY;
            Point2d p2p9 = new Point2d(x9, y9);
            double p9 = 0.4142;

            double x10 = x9 - rSSV9i;
            double y10 = y9 + rSSV9i;
            Point2d p2p10 = new Point2d(x10, y10);
            double p10 = 0;

            double x11 = x10 - disBX;
            double y11 = y10;
            Point2d p2p11 = new Point2d(x11, y11);
            double p11 = 0.4142;

            double x12 = x11 - rSSV9i;
            double y12 = y11 - rSSV9i;
            Point2d p2p12 = new Point2d(x12, y12);
            double p12 = 0.1989;

            double x13 = p3Orig.X - rSSV9i * Math.Cos(theta);
            double y13 = p3Orig.Y - rSSV9i * Math.Sin(theta);
            Point2d p2p13 = new Point2d(x13, y13);
            double p13 = 0;

            double x14 = x13 + cDis*Math.Sin(theta);
            double y14 = y13 - cDis*Math.Cos(theta);
            Point2d p2p14 = new Point2d(x14, y14);
            double p14 = 0;

            using (Transaction trSSV9i = aCurDB.TransactionManager.StartTransaction())
            {
                BlockTable blktbl;
                blktbl = trSSV9i.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord bltrec;
                bltrec = trSSV9i.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                double lineLen = 0;
                using (Polyline acPoly = new Polyline())
                {
                    acPoly.AddVertexAt(0, p2p1, p1, 0, 0);
                    acPoly.AddVertexAt(1, p2p2, p2, 0, 0);
                    acPoly.AddVertexAt(2, p2p3, p3, 0, 0);
                    acPoly.AddVertexAt(3, p2p4, p4, 0, 0);
                    acPoly.AddVertexAt(4, p2p5, p5, 0, 0);
                    acPoly.AddVertexAt(5, p2p6, p6, 0, 0);
                    acPoly.AddVertexAt(6, p2p7, p7, 0, 0);
                    acPoly.AddVertexAt(7, p2p8, p8, 0, 0);
                    acPoly.AddVertexAt(8, p2p9, p9, 0, 0);
                    acPoly.AddVertexAt(9, p2p10, p10, 0, 0);
                    acPoly.AddVertexAt(10, p2p11, p11, 0, 0);
                    acPoly.AddVertexAt(11, p2p12, p12, 0, 0);
                    acPoly.AddVertexAt(12, p2p13, p13, 0, 0);
                    acPoly.AddVertexAt(13, p2p14, p14, 0, 0);

                    acPoly.ColorIndex = stirrupSectionColor; //2;
                    lineLen = acPoly.Length;
                    bltrec.AppendEntity(acPoly);
                    trSSV9i.AddNewlyCreatedDBObject(acPoly, true);
                }

                actDoc.Editor.WriteMessage("\nStirrup length: {0}", lineLen);
                trSSV9i.Commit();
            }
        }


        /**
         * Date added: 19 May 2024 6:19pm
         * Added by: Bernardo A. Cabebe Jr.
         * Venue: Marquinton Carwash
         * Status: Under development
         */
        [CommandMethod("CBTI_ConcreteBeamTextInfo")]
        public static void ConcreteBeamTextInfo()
        {
            if (!ccData.isLicenseActive())
            {
                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            TypedValue[] tvFilterW0p = new TypedValue[2];
            tvFilterW0p[0] = new TypedValue((int)DxfCode.Start, "LINE");
            tvFilterW0p[1] = new TypedValue((int)DxfCode.Color, botLinCol); //53);
            SelectionFilter sfCBTW0p = new SelectionFilter(tvFilterW0p);

            PromptSelectionOptions psoCBTW0p = new PromptSelectionOptions();
            psoCBTW0p.MessageForAdding = "\nSelect bottom lines of beam: ";

            PromptSelectionResult psrCBTW0p = actDoc.Editor.GetSelection(psoCBTW0p, sfCBTW0p);
            if (psrCBTW0p.Status != PromptStatus.OK) return;

            SelectionSet ssCBTW0p = psrCBTW0p.Value;
            if (ssCBTW0p.Count < 1) return;

            PromptKeywordOptions pkoCBTW0p = new PromptKeywordOptions("");
            pkoCBTW0p.Message = "\nStart: ";
            pkoCBTW0p.Keywords.Add("Left");
            pkoCBTW0p.Keywords.Add("Right");
            pkoCBTW0p.Keywords.Default = infoStartG;
            pkoCBTW0p.AllowNone = true;

            PromptResult prCBTW0p = actDoc.Editor.GetKeywords(pkoCBTW0p);
            if (prCBTW0p.Status != PromptStatus.OK) return;
            string infoStart = prCBTW0p.StringResult;
            infoStartG = infoStart;

            PromptDoubleOptions pdoCBTW0p = new PromptDoubleOptions("");
            pdoCBTW0p.Message = "\nEnter text height: ";
            pdoCBTW0p.AllowNegative = false;
            pdoCBTW0p.AllowZero = false;
            pdoCBTW0p.DefaultValue = txtHgtG;
            pdoCBTW0p.AllowNone = true;

            PromptDoubleResult pdrCBTW0p = actDoc.Editor.GetDouble(pdoCBTW0p);
            if (pdrCBTW0p.Status != PromptStatus.OK) return;

            double textHgt = pdrCBTW0p.Value;
            txtHgtG = textHgt;

            using (Transaction trCTBW0p = aCurDB.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject soCBT in ssCBTW0p)
                {
                    Curve cuObj = trCTBW0p.GetObject(soCBT.ObjectId, OpenMode.ForRead) as Curve;
                    Point3d p3ObjP1 = cuObj.StartPoint;
                    Point3d p3ObjP2 = cuObj.EndPoint;

                    Point3d p3LeftPt = new Point3d();
                    Point3d p3RightPt = new Point3d();
                    if (p3ObjP1.X < p3ObjP2.X)
                    {
                        p3LeftPt = p3ObjP1;
                        p3RightPt = p3ObjP2;
                    }
                    else
                    {
                        p3LeftPt = p3ObjP2;
                        p3RightPt = p3ObjP1;
                    }
                    // ---20 May 2024---

                    double objLen = cuObj.GetDistAtPoint(cuObj.EndPoint) - cuObj.GetDistAtPoint(cuObj.StartPoint);
                    double disL8 = objLen / 8;
                    double offDisY = 0.6 * textHgt; //150;
                    double txtHgt = textHgt;
                    if (objLen < 100)
                    {
                        offDisY = 0.6 * textHgt / 1000; // 0.15;
                        txtHgt = textHgt / 1000; // 0.25;
                    }

                    Vector3d v3LtoR = p3LeftPt.GetVectorTo(p3RightPt);
                    Vector3d nvLtoR = v3LtoR.GetNormal();
                    Vector3d nvLtR1 = disL8 * nvLtoR;
                    Point3d p3Tx1P = p3LeftPt + nvLtR1;
                    p3Tx1P = p3Tx1P + new Vector3d(0, -offDisY, 0);

                    Vector3d nvLtR2 = 7 * disL8 * nvLtoR;
                    Point3d p3Tx2P = p3LeftPt + nvLtR2;
                    p3Tx2P = p3Tx2P + new Vector3d(0, -offDisY, 0);

                    Vector3d nvLtR3 = 4 * disL8 * nvLtoR;
                    Point3d p3Tx3P = p3LeftPt + nvLtR3;
                    p3Tx3P = p3Tx3P + new Vector3d(0, -offDisY, 0);

                    Vector3d nvLtR4 = 2.5 * disL8 * nvLtoR;
                    Point3d p3Tx4P = p3LeftPt + nvLtR4;
                    p3Tx4P = p3Tx4P + new Vector3d(0, -offDisY, 0);

                    Vector3d nvLtR5 = 3.95 * disL8 * nvLtoR;
                    Point3d p3Tx5P = p3LeftPt + nvLtR5;
                    p3Tx5P = p3Tx5P + new Vector3d(0, -2.25 * offDisY, 0);

                    Vector3d nvLtR6 = 4.05 * disL8 * nvLtoR;
                    Point3d p3Tx6P = p3LeftPt + nvLtR6;
                    p3Tx6P = p3Tx6P + new Vector3d(0, -2.25 * offDisY, 0);

                    Vector3d nvLtR7 = 5.5 * disL8 * nvLtoR;
                    Point3d p3Tx7P = p3LeftPt + nvLtR7;
                    p3Tx7P = p3Tx7P + new Vector3d(0, -offDisY, 0);

                    BlockTable blktbl;
                    blktbl = trCTBW0p.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord bltrec;
                    bltrec = trCTBW0p.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    MText mTx1 = new MText();
                    mTx1.Location = p3Tx1P;
                    if(infoStart == "Left")
                    {
                        mTx1.Contents = "E1:*";
                    }
                    else
                    {
                        mTx1.Contents = "E2:*";
                    }
                    mTx1.Attachment = AttachmentPoint.TopLeft;
                    mTx1.TextHeight = 0.5*txtHgt;
                    mTx1.ColorIndex = txtInfCol;
                    bltrec.AppendEntity(mTx1);
                    trCTBW0p.AddNewlyCreatedDBObject(mTx1, true);

                    MText mTx2 = new MText();
                    mTx2.Location = p3Tx2P;
                    
                    if (infoStart == "Left")
                    {
                        mTx2.Contents = "E2:*";
                    }
                    else
                    {
                        mTx2.Contents = "E1:*";
                    }
                    mTx2.Attachment = AttachmentPoint.TopRight;
                    mTx2.TextHeight = 0.5*txtHgt;
                    mTx2.ColorIndex = txtInfCol;
                    bltrec.AppendEntity(mTx2);
                    trCTBW0p.AddNewlyCreatedDBObject(mTx2, true);

                    MText mTx3 = new MText();
                    mTx3.Location = p3Tx3P;
                    mTx3.Contents = "MS:*";
                    mTx3.Attachment = AttachmentPoint.TopCenter;
                    mTx3.TextHeight = 0.5*txtHgt;
                    mTx3.ColorIndex = txtInfCol;
                    bltrec.AppendEntity(mTx3);
                    trCTBW0p.AddNewlyCreatedDBObject(mTx3, true);

                    MText mTx4 = new MText();
                    mTx4.Location = p3Tx4P;
                    mTx4.Contents = "Stirr.:*";
                    mTx4.Attachment = AttachmentPoint.TopCenter;
                    mTx4.TextHeight = 0.5*txtHgt;
                    mTx4.ColorIndex = txtInfCol;
                    bltrec.AppendEntity(mTx4);
                    trCTBW0p.AddNewlyCreatedDBObject(mTx4, true);

                    MText mTx5 = new MText();
                    mTx5.Location = p3Tx5P;
                    mTx5.Contents = "B-*";
                    mTx5.Attachment = AttachmentPoint.TopRight;
                    mTx5.TextHeight = txtHgt;
                    mTx5.ColorIndex = txtMarkCol;
                    bltrec.AppendEntity(mTx5);
                    trCTBW0p.AddNewlyCreatedDBObject(mTx5, true);

                    MText mTx6 = new MText();
                    mTx6.Location = p3Tx6P;
                    mTx6.Contents = "(BxH)";
                    mTx6.Attachment = AttachmentPoint.TopLeft;
                    mTx6.TextHeight = 0.75*txtHgt;
                    mTx6.ColorIndex = txtDimCol;
                    bltrec.AppendEntity(mTx6);
                    trCTBW0p.AddNewlyCreatedDBObject(mTx6, true);

                    MText mTx7 = new MText();
                    mTx7.Location = p3Tx7P;
                    mTx7.Contents = "WebD:10";
                    mTx7.Attachment = AttachmentPoint.TopCenter;
                    mTx7.TextHeight = 0.5 * txtHgt;
                    mTx7.ColorIndex = txtInfCol;
                    bltrec.AppendEntity(mTx7);
                    trCTBW0p.AddNewlyCreatedDBObject(mTx7, true);
                }

                trCTBW0p.Commit();
            }
        }


        /**
         * Date added: 23 May 2024
         * Added by: Bernardo A. Cabebe Jr.
         * Venue: 3407 Citland Pasong Tamo Tower
         * Status: Under development
         */
        [CommandMethod("CBI_CopyBarInformation")]
        public static void CopyBarInformation()
        {
            if (!ccData.isLicenseActive())
            {
                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            TypedValue[] tvFilter1D5t = new TypedValue[10];
            tvFilter1D5t[0] = new TypedValue((int)DxfCode.Start, "TEXT");
            tvFilter1D5t[1] = new TypedValue((int)DxfCode.Operator, "<OR");
            tvFilter1D5t[2] = new TypedValue((int)DxfCode.LayerName, "S-BeamMarks");
            tvFilter1D5t[3] = new TypedValue((int)DxfCode.LayerName, "S-GBRebars");
            tvFilter1D5t[4] = new TypedValue((int)DxfCode.Operator, "OR>");
            tvFilter1D5t[5] = new TypedValue((int)DxfCode.Operator, "<OR");
            tvFilter1D5t[6] = new TypedValue((int)DxfCode.Color, beamRebarInfoColor);
            tvFilter1D5t[7] = new TypedValue((int)DxfCode.Color, beamDimenSourceColor);
            tvFilter1D5t[8] = new TypedValue((int)DxfCode.Color, beamMarkSourceColor);
            tvFilter1D5t[9] = new TypedValue((int)DxfCode.Operator, "OR>");
            SelectionFilter sfCBI1D5t = new SelectionFilter(tvFilter1D5t);

            PromptSelectionOptions psoCBI1D5t = new PromptSelectionOptions();
            psoCBI1D5t.MessageForAdding = "\nSelect source information: ";

            PromptSelectionResult psrCBI1D5t = actDoc.Editor.GetSelection(psoCBI1D5t, sfCBI1D5t);
            if (psrCBI1D5t.Status != PromptStatus.OK) return;

            SelectionSet ssDBI1D5t = psrCBI1D5t.Value;
            if(ssDBI1D5t.Count != 7)
            {
                actDoc.Editor.WriteMessage("\nInvalid number of selection. Exit...");
                return;
            }

            TypedValue[] tvFilter2D5t = new TypedValue[6];
            tvFilter2D5t[0] = new TypedValue((int)DxfCode.Start, "MTEXT");
            tvFilter2D5t[1] = new TypedValue((int)DxfCode.Operator, "<OR");
            tvFilter2D5t[2] = new TypedValue((int)DxfCode.Color, txtMarkCol);
            tvFilter2D5t[3] = new TypedValue((int)DxfCode.Color, txtDimCol);
            tvFilter2D5t[4] = new TypedValue((int)DxfCode.Color, txtInfCol);
            tvFilter2D5t[5] = new TypedValue((int)DxfCode.Operator, "OR>");
            SelectionFilter sfCBI2D5t = new SelectionFilter(tvFilter2D5t);

            PromptSelectionOptions psoCBI2D5t = new PromptSelectionOptions();
            psoCBI2D5t.MessageForAdding = "\nSelect destination: ";

            PromptSelectionResult psrCBI2D5t = actDoc.Editor.GetSelection(psoCBI2D5t, sfCBI2D5t);
            if (psrCBI2D5t.Status != PromptStatus.OK) return;

            SelectionSet ssDBI2D5t = psrCBI2D5t.Value;
            if (ssDBI1D5t.Count != 7)
            {
                actDoc.Editor.WriteMessage("\nInvalid number of selection. Exit...");
                return;
            }

            using (Transaction trCBI1D5t = aCurDB.TransactionManager.StartTransaction())
            {
                //Check if beam marks are matched
                string e1Info = "";
                string e2Info = "";
                string msInfo = "";
                string stInfo = "";
                string bhInfo = "";
                string bmInfo = "";
                foreach (SelectedObject soCBI1D5t in ssDBI1D5t)
                {
                    DBText soDBText1D5t = trCBI1D5t.GetObject(soCBI1D5t.ObjectId, OpenMode.ForRead) as DBText;

                    string txtValue1D5t = soDBText1D5t.TextString;
                    actDoc.Editor.WriteMessage("\nBeam source info: {0}", txtValue1D5t);

                    string patt1 = @"(?<=E1\:)\d+";
                    Regex regex1 = new Regex(patt1);
                    Match match1 = regex1.Match(txtValue1D5t);
                    if (match1.Success)
                    {
                        e1Info = txtValue1D5t;
                    }

                    string patt2 = @"(?<=E2\:)\d+";
                    Regex regex2 = new Regex(patt2);
                    Match match2 = regex2.Match(txtValue1D5t);
                    if (match2.Success)
                    {
                        e2Info = txtValue1D5t;
                    }

                    string patt3 = @"(?<=MS\:)\d+";
                    Regex regex3 = new Regex(patt3);
                    Match match3 = regex3.Match(txtValue1D5t);
                    if (match3.Success)
                    {
                        msInfo = txtValue1D5t;
                    }

                    string patt4 = @"\d+[xX]\d+";
                    Regex regex4 = new Regex(patt4);
                    Match match4 = regex4.Match(txtValue1D5t);
                    if (match4.Success)
                    {
                        bhInfo = txtValue1D5t;
                    }

                    string patt5 = @"^[BGC][A-Z0-9]+\-\d+";
                    Regex regex5 = new Regex(patt5);
                    Match match5 = regex5.Match(txtValue1D5t);
                    if (match5.Success)
                    {
                        bmInfo = txtValue1D5t;
                    }

                    string patt6 = @"[rR]\.\:[A-Z]+";
                    Regex regex6 = new Regex(patt6);
                    Match match6 = regex6.Match(txtValue1D5t);
                    if (match6.Success)
                    {
                        stInfo = txtValue1D5t;
                    }
                }

                foreach (SelectedObject soCBI2D5t in ssDBI2D5t)
                {
                    MText soMText2D5t = trCBI1D5t.GetObject(soCBI2D5t.ObjectId, OpenMode.ForRead) as MText;

                    string txtValue2D5t = soMText2D5t.Text;
                    actDoc.Editor.WriteMessage("\nBeam destination: {0}", txtValue2D5t);

                    string patt1 = @"(?<=E1\:).+";
                    Regex regex1 = new Regex(patt1);
                    Match match1 = regex1.Match(txtValue2D5t);
                    if (match1.Success)
                    {
                        soMText2D5t.UpgradeOpen();
                        soMText2D5t.Contents = e1Info;
                    }

                    string patt2 = @"(?<=E2\:).+";
                    Regex regex2 = new Regex(patt2);
                    Match match2 = regex2.Match(txtValue2D5t);
                    if (match2.Success)
                    {
                        soMText2D5t.UpgradeOpen();
                        soMText2D5t.Contents = e2Info;
                    }

                    string patt3 = @"(?<=MS\:).+";
                    Regex regex3 = new Regex(patt3);
                    Match match3 = regex3.Match(txtValue2D5t);
                    if (match3.Success)
                    {
                        soMText2D5t.UpgradeOpen();
                        soMText2D5t.Contents = msInfo;
                    }

                    string patt4 = @"B[xX]H";
                    Regex regex4 = new Regex(patt4);
                    Match match4 = regex4.Match(txtValue2D5t);
                    if (match4.Success)
                    {
                        soMText2D5t.UpgradeOpen();
                        soMText2D5t.Contents = bhInfo;
                    }

                    string patt5 = @"B\-\*";
                    Regex regex5 = new Regex(patt5);
                    Match match5 = regex5.Match(txtValue2D5t);
                    if (match5.Success)
                    {
                        soMText2D5t.UpgradeOpen();
                        soMText2D5t.Contents = bmInfo;
                    }

                    string patt6 = @"[rR]\.\:.+";
                    Regex regex6 = new Regex(patt6);
                    Match match6 = regex6.Match(txtValue2D5t);
                    if (match6.Success)
                    {
                        soMText2D5t.UpgradeOpen();
                        soMText2D5t.Contents = stInfo;
                    }
                }

                trCBI1D5t.Commit();
            }
        }


        /**
         * Date added: 24 Mar 2024
         * Added by: Bernardo A. Cabebe Jr.
         * Venue: 3407 Cityland Pasong Tamo Tower
         * Status: Under development
         */
        [CommandMethod("CLN_CreateLabelNBar")]
        public static void CreateLabelNBar()
        {
            if (!ccData.isLicenseActive())
            {

                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            //Select entity 1 (Bar line 1)
            actDoc.Editor.WriteMessage("\nAllowed rebar lines are web, slab, stirrup and stirrup section.");
            PromptEntityOptions peoCLN1G8i = new PromptEntityOptions("\nSelect rebar line: ");
            peoCLN1G8i.SetRejectMessage("\nSelect polyline or line only.");
            peoCLN1G8i.AddAllowedClass(typeof(Polyline), true);
            peoCLN1G8i.AddAllowedClass(typeof(Line), true);

            PromptEntityResult perCLN1G8i = actDoc.Editor.GetEntity(peoCLN1G8i);
            if (perCLN1G8i.Status != PromptStatus.OK) return;

            Point3d pickPt1 = perCLN1G8i.PickedPoint;

            PromptKeywordOptions pkwCLN1G8i = new PromptKeywordOptions("");
            pkwCLN1G8i.Message = "\nSelect bar type: ";
            pkwCLN1G8i.Keywords.Add("Web");
            pkwCLN1G8i.Keywords.Add("STirrup");
            pkwCLN1G8i.Keywords.Add("SLab");
            pkwCLN1G8i.Keywords.Default = selectBarTypeG;
            pkwCLN1G8i.AllowNone = true;

            PromptResult prCLN1G8i = actDoc.Editor.GetKeywords(pkwCLN1G8i);
            if (prCLN1G8i.Status != PromptStatus.OK) return;
            string barType1G8i = prCLN1G8i.StringResult;
            actDoc.Editor.WriteMessage("\nBar type: {0}", barType1G8i);
            selectBarTypeG = barType1G8i;

            TypedValue[] tvFilter1G8i;
            switch (barType1G8i)
            {
                case "Web":
                    tvFilter1G8i = new TypedValue[2];
                    tvFilter1G8i[0] = new TypedValue((int)DxfCode.Start, "LINE");
                    tvFilter1G8i[1] = new TypedValue((int)DxfCode.Color, webLinCol);
                    break;
                case "STirrup":
                    tvFilter1G8i = new TypedValue[2];
                    tvFilter1G8i[0] = new TypedValue((int)DxfCode.Start, "LINE");
                    tvFilter1G8i[1] = new TypedValue((int)DxfCode.Color, stirLnCol);
                    break;
                case "SLab":
                    tvFilter1G8i = new TypedValue[2];
                    tvFilter1G8i[0] = new TypedValue((int)DxfCode.Start, "LINE");
                    tvFilter1G8i[1] = new TypedValue((int)DxfCode.Color, slabBarLineColor);
                    break;
                default:
                    actDoc.Editor.WriteMessage("\nUnknown bar type. Exit...");
                    return;
            } 
            SelectionFilter sfCLN1G8i = new SelectionFilter(tvFilter1G8i);

            PromptSelectionOptions psoCLN1G8i = new PromptSelectionOptions();
            psoCLN1G8i.MessageForAdding = "\nSelect multiple bar line: ";

            PromptSelectionResult psrCLN1G8i = actDoc.Editor.GetSelection(psoCLN1G8i, sfCLN1G8i);
            if (psrCLN1G8i.Status != PromptStatus.OK) return;

            SelectionSet ssCLN1G8i = psrCLN1G8i.Value;
            if (ssCLN1G8i.Count < 1)
            {
                actDoc.Editor.WriteMessage("\nInvalid number of selection. Exit...");
                return;
            }

            using(Transaction trCLN1G8i = aCurDB.TransactionManager.StartTransaction())
            {
                Curve cuLine1G8i = trCLN1G8i.GetObject(perCLN1G8i.ObjectId, OpenMode.ForRead) as Curve;
                double cuLineLength = cuLine1G8i.GetDistAtPoint(cuLine1G8i.EndPoint) - cuLine1G8i.GetDistAtPoint(cuLine1G8i.StartPoint);
                int lineColor1G8i = cuLine1G8i.ColorIndex;

                switch (lineColor1G8i)
                {
                    case webLinCol:
                        actDoc.Editor.WriteMessage("\nWeb bar line.");
                        break;
                    case slabBarLineColor:
                        actDoc.Editor.WriteMessage("\nSlab bar line.");
                        break;
                    case stirrupSectionColor:
                        actDoc.Editor.WriteMessage("\nStirrup section bar line.");
                        break;
                    default:
                        actDoc.Editor.WriteMessage("\nInvalid bar line selection. Exit...");
                        return;
                }

                int qtyBar = 0;
                string barDesc1G8i = "-*-";
                PromptStringOptions psoCLN2G8i = new PromptStringOptions("");
                PromptResult prCLN2G8i;
                string barDia1G8i = "";
                double barLen2G8i = 0.0;
                switch (barType1G8i)
                {
                    case "Web":
                        foreach(SelectedObject soCLN1G8i in ssCLN1G8i)
                        {
                            Curve cuCLN2G8i = trCLN1G8i.GetObject(soCLN1G8i.ObjectId, OpenMode.ForRead) as Curve;
                            double cuObjLength1G8i = cuCLN2G8i.GetDistAtPoint(cuCLN2G8i.EndPoint) - cuCLN2G8i.GetDistAtPoint(cuCLN2G8i.StartPoint);
                            double tolerance1G8i = 25.0;
                            if(cuLineLength < 100)
                            {
                                tolerance1G8i = 0.025;
                            }
                            if(Math.Abs(cuLineLength - cuObjLength1G8i) <= tolerance1G8i)
                            {
                                qtyBar++;
                                cuCLN2G8i.UpgradeOpen();
                                cuCLN2G8i.ColorIndex = countedWebLineColor;
                            }
                        }
                        psoCLN2G8i.Message = "\nEnter web bar diameter: ";
                        psoCLN2G8i.DefaultValue = webDiaG.ToString();
                        psoCLN2G8i.AllowSpaces = false;

                        prCLN2G8i = actDoc.Editor.GetString(psoCLN2G8i);
                        if (prCLN2G8i.Status != PromptStatus.OK) return;
                        barDia1G8i = prCLN2G8i.StringResult;
                        webDiaG = Convert.ToInt16(barDia1G8i);

                        barLen2G8i = Math.Round(cuLineLength, 2);
                        if(cuLineLength > 100)
                        {
                            barLen2G8i = Math.Round(cuLineLength / 1000, 2);
                        }
                        barDesc1G8i = "WEB" + RebarId() + ": " + (qtyBar * 2) + " - " + barDia1G8i + "D x " + barLen2G8i + "m";
                        break;
                    case "STirrup":
                        psoCLN2G8i.Message = "\nEnter stirrup bar diameter: ";
                        psoCLN2G8i.DefaultValue = stiDiaG.ToString();
                        psoCLN2G8i.AllowSpaces = false;

                        prCLN2G8i = actDoc.Editor.GetString(psoCLN2G8i);
                        if (prCLN2G8i.Status != PromptStatus.OK) return;
                        barDia1G8i = prCLN2G8i.StringResult;
                        stiDiaG = Convert.ToInt16(barDia1G8i);

                        qtyBar = ssCLN1G8i.Count;
                        barLen2G8i = Math.Round(cuLineLength, 2);
                        if (cuLineLength > 100)
                        {
                            barLen2G8i = Math.Round(cuLineLength / 1000, 2);
                        }
                        barDesc1G8i = "STI" + RebarId() + ": " + qtyBar + " - " + stiDiaG + "D x " + barLen2G8i + "m";
                        break;
                    case "SLab":
                        //For development
                        actDoc.Editor.WriteMessage("\nUnder development. Exit...");
                        break;
                    default:
                        actDoc.Editor.WriteMessage("\nUnknown bar type. Exit...");
                        return;
                }


                Curve cuBarLine = trCLN1G8i.GetObject(perCLN1G8i.ObjectId, OpenMode.ForRead) as Curve;
                Point3d p3dPt1 = cuBarLine.GetClosestPointTo(pickPt1, false);

                PromptPointOptions ppoPt2 = new PromptPointOptions("");
                ppoPt2.Message = "\nEnter point 2: ";
                ppoPt2.UseBasePoint = true;
                ppoPt2.BasePoint = pickPt1;
                PromptPointResult pprPt2 = actDoc.Editor.GetPoint(ppoPt2);
                if (pprPt2.Status != PromptStatus.OK) return;
                Point3d p3dPt2 = pprPt2.Value;

                PromptPointOptions ppoPt3 = new PromptPointOptions("");
                ppoPt3.Message = "\nEnter point 3: ";
                ppoPt3.UseBasePoint = true;
                ppoPt3.BasePoint = p3dPt2;
                PromptPointResult pprPt3 = actDoc.Editor.GetPoint(ppoPt3);
                if (pprPt3.Status != PromptStatus.OK) return;
                Point3d p3dPt3 = pprPt3.Value;

                PromptDoubleOptions pdoCLLS6y = new PromptDoubleOptions("");
                pdoCLLS6y.Message = "\nEnter text height: ";
                pdoCLLS6y.AllowNegative = false;
                pdoCLLS6y.AllowZero = false;
                pdoCLLS6y.DefaultValue = txtHgtG;
                pdoCLLS6y.AllowNone = true;

                PromptDoubleResult pdrCLLS6y = actDoc.Editor.GetDouble(pdoCLLS6y);
                if (pdrCLLS6y.Status != PromptStatus.OK) return;

                double textHgt = pdrCLLS6y.Value;
                txtHgtG = textHgt;

                MText mTxt = new MText();
                mTxt.Location = p3dPt3;
                // mTxt.Contents = "Some text in the default colour...\\P" +
                //                 "{\\C1;Something red}\\P" +
                //                 "{\\C2;Something yellow}\\P" +
                //                 "{\\C3;And} {\\C4;something} " +
                //                 "{\\C5;multi-}{\\C6;coloured}\\P";
                mTxt.Contents = barDesc1G8i;// "TEXT CONTENT 1 \\PTEXT CONTENT 2";

                if (p3dPt3.X > p3dPt2.X)
                {
                    mTxt.Attachment = AttachmentPoint.MiddleLeft;
                }
                else
                {
                    mTxt.Attachment = AttachmentPoint.MiddleRight;
                }
                mTxt.ColorIndex = labelColor; //3;

                double arrowSize = 0.75 * txtHgtG; //185;
                if (p3dPt3.DistanceTo(p3dPt2) > 50)
                {
                    mTxt.TextHeight = txtHgtG; //250;
                    arrowSize = 0.75 * txtHgtG; // 185;
                }
                else
                {
                    mTxt.TextHeight = txtHgtG / 1000; // 0.250;
                    arrowSize = 0.75 * txtHgtG / 1000; // 0.185;
                }

                Leader acLdr = new Leader();
                acLdr.AppendVertex(p3dPt1);
                acLdr.AppendVertex(p3dPt2);
                acLdr.AppendVertex(p3dPt3);
                acLdr.HasArrowHead = true;
                acLdr.Dimasz = arrowSize;
                acLdr.ColorIndex = 251;

                BlockTable btLL = (BlockTable)trCLN1G8i.GetObject(aCurDB.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btrLL = (BlockTableRecord)trCLN1G8i.GetObject(btLL[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                ObjectId oiLL = btrLL.AppendEntity(mTxt);
                trCLN1G8i.AddNewlyCreatedDBObject(mTxt, true);

                btrLL.AppendEntity(acLdr);
                trCLN1G8i.AddNewlyCreatedDBObject(acLdr, true);

                trCLN1G8i.Commit();
            }
        }



        /**
         * Date added: 26 May 2024
         * Added by: Bernardo A. Cabebe Jr.
         * Venue: 3407 Cityland Pasong Tamo Tower
         * Status: Under development
         */
        [CommandMethod("SS1_StirrupSection1")]
        public static void StirrupSection1()
        {
            if (!ccData.isLicenseActive())
            {
                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            //Get the beam dimension
            TypedValue[] tvFilterV9j = new TypedValue[2];
            tvFilterV9j[0] = new TypedValue((int)DxfCode.Start, "MTEXT");
            tvFilterV9j[1] = new TypedValue((int)DxfCode.Color, txtDimCol);
            SelectionFilter sfSSV9j = new SelectionFilter(tvFilterV9j);

            PromptSelectionOptions psoSSV9j = new PromptSelectionOptions();
            psoSSV9j.MessageForAdding = "\nSelect beam dimension: ";

            PromptSelectionResult psrSSV9j = actDoc.Editor.GetSelection(psoSSV9j, sfSSV9j);
            if (psrSSV9j.Status != PromptStatus.OK) return;

            SelectionSet ssSSV9j = psrSSV9j.Value;
            if(ssSSV9j.Count != 1)
            {
                actDoc.Editor.WriteMessage("\nInvalid number of selection. Exit...");
                return;
            }


            PromptKeywordOptions pkoSSV9j = new PromptKeywordOptions("");
            pkoSSV9j.Message = "\nStirrup draw line: ";
            pkoSSV9j.Keywords.Add("Perimeter");
            pkoSSV9j.Keywords.Add("Center");
            pkoSSV9j.Keywords.Default = stirrupDrawLineG;
            pkoSSV9j.AllowNone = true;

            PromptResult pkrSSV9j = actDoc.Editor.GetKeywords(pkoSSV9j);
            if (pkrSSV9j.Status != PromptStatus.OK) return;
            string stirrupDrawLine = pkrSSV9j.StringResult;
            stirrupDrawLineG = stirrupDrawLine;


            double disBX = 0.0;
            double disBY = 0.0;
            using (Transaction trSSV9j = aCurDB.TransactionManager.StartTransaction())
            {
                MText mtxt1V9j = trSSV9j.GetObject(ssSSV9j[0].ObjectId, OpenMode.ForRead) as MText;
                string txCont1V9j = mtxt1V9j.Contents;

                string patt1V9j = @"(?<=\()[0-9]+";
                Regex regex1V9j = new Regex(patt1V9j);
                Match match1V9j = regex1V9j.Match(txCont1V9j);
                string width1V9j = "";
                if(match1V9j.Success)
                {
                    width1V9j = match1V9j.Value;
                }
                disBX = Convert.ToDouble(width1V9j);

                string patt2V9j = @"(?<=[xX])[0-9]+";
                Regex regex2V9j = new Regex(patt2V9j);
                Match match2V9j = regex2V9j.Match(txCont1V9j);
                string height1V9j = "";
                if (match2V9j.Success)
                {
                    height1V9j = match2V9j.Value;
                }
                disBY = Convert.ToDouble(height1V9j);
            }

            //Get from user the rebar diameter and lap factor
            PromptStringOptions psoSSV9i = new PromptStringOptions("");
            string defSSV9i = "DB=" + barDiaG.ToString() + ";DS=" + stiDiaG.ToString();
            psoSSV9i.Message = "\nEnter values of DB, DS: ";
            psoSSV9i.DefaultValue = defSSV9i;
            psoSSV9i.AllowSpaces = false;

            PromptResult prSSV9i = actDoc.Editor.GetString(psoSSV9i);
            if (prSSV9i.Status != PromptStatus.OK) return;

            string strResSSV9i = prSSV9i.StringResult;
            actDoc.Editor.WriteMessage("\nInput: {0}", strResSSV9i);

            string patt1 = @"(?<=DB\=)\d+";
            Regex regexDb = new Regex(patt1);
            MatchCollection matchDb = regexDb.Matches(strResSSV9i);
            double barDia = Convert.ToDouble(matchDb[0].Value);
            actDoc.Editor.WriteMessage("\nMain bar diameter: {0}", barDia);
            barDiaG = barDia;

            string patt2 = @"(?<=DS\=)\d+";
            Regex regexDs = new Regex(patt2);
            MatchCollection matchDs = regexDs.Matches(strResSSV9i);
            double stiDia = Convert.ToDouble(matchDs[0].Value);
            actDoc.Editor.WriteMessage("\nStirrup bar diameter: {0}", stiDia);
            stiDiaG = stiDia;

            PromptPointOptions ppoSSV9i = new PromptPointOptions("\nEnter point: ");
            PromptPointResult pprSSV9i = actDoc.Editor.GetPoint(ppoSSV9i);

            Point3d p3Orig = pprSSV9i.Value;


            disBX = disBX - 2 * concCoverG - 2 * stiDia - barDia;
            disBY = disBY - 2 * concCoverG - 2 * stiDia - barDia;

            double theta = 45 * Math.PI / 180; //radians
            double rSSV9i = 0;
            switch (stirrupDrawLine)
            {
                case "Perimeter":
                    rSSV9i = 0.5 * barDia + stiDia;
                    break;
                case "Center":
                    rSSV9i = 0.5 * barDia + 0.5 * stiDia;
                    break;
                default:
                    rSSV9i = 0;
                    break;
            }
            
            double cDis = 6 * stiDia;

            double x2 = p3Orig.X + rSSV9i * Math.Cos(theta);
            double y2 = p3Orig.Y + rSSV9i * Math.Sin(theta);
            Point2d p2p2 = new Point2d(x2, y2);
            double p2 = 0.1989;

            double x1 = cDis * Math.Sin(theta) + x2;
            double y1 = y2 - cDis * Math.Cos(theta);
            Point2d p2p1 = new Point2d(x1, y1);
            double p1 = 0;

            double x3 = p3Orig.X;
            double y3 = p3Orig.Y + rSSV9i;
            Point2d p2p3 = new Point2d(x3, y3);
            double p3 = 0.4142;

            double x4 = p3Orig.X - rSSV9i;
            double y4 = p3Orig.Y;
            Point2d p2p4 = new Point2d(x4, y4);
            double p4 = 0;

            double x5 = x4;
            double y5 = y4 - disBY;
            Point2d p2p5 = new Point2d(x5, y5);
            double p5 = 0.4142;

            double x6 = x5 + rSSV9i;
            double y6 = y5 - rSSV9i;
            Point2d p2p6 = new Point2d(x6, y6);
            double p6 = 0;

            double x7 = x6 + disBX;
            double y7 = y6;
            Point2d p2p7 = new Point2d(x7, y7);
            double p7 = 0.4142;

            double x8 = x7 + rSSV9i;
            double y8 = y7 + rSSV9i;
            Point2d p2p8 = new Point2d(x8, y8);
            double p8 = 0;

            double x9 = x8;
            double y9 = y8 + disBY;
            Point2d p2p9 = new Point2d(x9, y9);
            double p9 = 0.4142;

            double x10 = x9 - rSSV9i;
            double y10 = y9 + rSSV9i;
            Point2d p2p10 = new Point2d(x10, y10);
            double p10 = 0;

            double x11 = x10 - disBX;
            double y11 = y10;
            Point2d p2p11 = new Point2d(x11, y11);
            double p11 = 0.4142;

            double x12 = x11 - rSSV9i;
            double y12 = y11 - rSSV9i;
            Point2d p2p12 = new Point2d(x12, y12);
            double p12 = 0.1989;

            double x13 = p3Orig.X - rSSV9i * Math.Cos(theta);
            double y13 = p3Orig.Y - rSSV9i * Math.Sin(theta);
            Point2d p2p13 = new Point2d(x13, y13);
            double p13 = 0;

            double x14 = x13 + cDis * Math.Sin(theta);
            double y14 = y13 - cDis * Math.Cos(theta);
            Point2d p2p14 = new Point2d(x14, y14);
            double p14 = 0;

            using (Transaction trSSV9i = aCurDB.TransactionManager.StartTransaction())
            {
                BlockTable blktbl;
                blktbl = trSSV9i.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord bltrec;
                bltrec = trSSV9i.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                double lineLen = 0;
                using (Polyline acPoly = new Polyline())
                {
                    acPoly.AddVertexAt(0, p2p1, p1, 0, 0);
                    acPoly.AddVertexAt(1, p2p2, p2, 0, 0);
                    acPoly.AddVertexAt(2, p2p3, p3, 0, 0);
                    acPoly.AddVertexAt(3, p2p4, p4, 0, 0);
                    acPoly.AddVertexAt(4, p2p5, p5, 0, 0);
                    acPoly.AddVertexAt(5, p2p6, p6, 0, 0);
                    acPoly.AddVertexAt(6, p2p7, p7, 0, 0);
                    acPoly.AddVertexAt(7, p2p8, p8, 0, 0);
                    acPoly.AddVertexAt(8, p2p9, p9, 0, 0);
                    acPoly.AddVertexAt(9, p2p10, p10, 0, 0);
                    acPoly.AddVertexAt(10, p2p11, p11, 0, 0);
                    acPoly.AddVertexAt(11, p2p12, p12, 0, 0);
                    acPoly.AddVertexAt(12, p2p13, p13, 0, 0);
                    acPoly.AddVertexAt(13, p2p14, p14, 0, 0);

                    acPoly.ColorIndex = stirrupSectionColor; //2;
                    lineLen = acPoly.Length;
                    bltrec.AppendEntity(acPoly);
                    trSSV9i.AddNewlyCreatedDBObject(acPoly, true);
                }

                actDoc.Editor.WriteMessage("\nStirrup length: {0}", lineLen);
                trSSV9i.Commit();
            }
        }



        /**
         * Date added: 26 May 2024
         * Added by: Bernardo A. Cabebe Jr.
         * Venue: 3407 Cityland Pasong Tamo Tower
         * Status: Under development
         */
        [CommandMethod("CBS_ConcreteBeamSection")]
        public static void ConcreteBeamSection()
        {
            if (!ccData.isLicenseActive())
            {
                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            //Get the beam dimension
            TypedValue[] tvFilter1F0u = new TypedValue[5];
            tvFilter1F0u[0] = new TypedValue((int)DxfCode.Start, "MTEXT");
            tvFilter1F0u[1] = new TypedValue((int)DxfCode.Operator, "<OR");
            tvFilter1F0u[2] = new TypedValue((int)DxfCode.Color, txtDimCol);
            tvFilter1F0u[3] = new TypedValue((int)DxfCode.Color, txtInfCol);
            tvFilter1F0u[4] = new TypedValue((int)DxfCode.Operator, "OR>");
            SelectionFilter sfCBS1F0u = new SelectionFilter(tvFilter1F0u);

            PromptSelectionOptions psoCBS1F0u = new PromptSelectionOptions();
            psoCBS1F0u.MessageForAdding = "\nSelect beam dimension and rebar info: ";

            PromptSelectionResult psrCBS1F0u = actDoc.Editor.GetSelection(psoCBS1F0u, sfCBS1F0u);
            if (psrCBS1F0u.Status != PromptStatus.OK) return;

            SelectionSet ssDBS1F0u = psrCBS1F0u.Value;
            if(ssDBS1F0u.Count != 2)
            {
                actDoc.Editor.WriteMessage("\nInvalid number of selection. Exit...");
                return;
            }

            PromptStringOptions psoCBS2F0u = new PromptStringOptions("");
            string defVal2F0u = "\nCC=" + concCoverG.ToString() + ",SD=" + stiDiaG.ToString();
            psoCBS2F0u.Message = "\nEnter Concrete Cover CC and Stirrup Diameter SD: ";
            psoCBS2F0u.DefaultValue = defVal2F0u;
            psoCBS2F0u.AllowSpaces = false;

            PromptResult prCBS2F0u = actDoc.Editor.GetString(psoCBS2F0u);
            if (prCBS2F0u.Status != PromptStatus.OK) return;

            string inputRes2F0u = prCBS2F0u.StringResult;
            if(inputRes2F0u.Length < 11)
            {
                actDoc.Editor.WriteMessage("\nCheck input text. Exit...");
                return;
            }

            double concCov = 0.0;
            string pattCC = @"(?<=C\=)\d+";
            Regex regexCC = new Regex(pattCC);
            Match matchCC = regexCC.Match(inputRes2F0u);
            if (matchCC.Success)
            {
                concCov = Convert.ToDouble(matchCC.Value);
            }
            concCoverG = concCov;

            double strDia = 0.0;
            string pattSD = @"(?<=D\=)\d+";
            Regex regexSD = new Regex(pattSD);
            Match matchSD = regexSD.Match(inputRes2F0u);
            if(matchSD.Success)
            {
                strDia = Convert.ToDouble(matchSD.Value);
            }
            stiDiaG = strDia;

            using (Transaction trCBS1F0u = aCurDB.TransactionManager.StartTransaction())
            {
                double barDia1F0u = 0.0;
                int qtyTop1F0u = 0;
                int qtyBot1F0u = 0;
                double wid1F0u = 0.0;
                double hgt1F0u = 0.0;

                BlockTable blktbl;
                blktbl = trCBS1F0u.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord bltrec;
                bltrec = trCBS1F0u.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                foreach (SelectedObject soCBS1F0u in ssDBS1F0u)
                {
                    MText mtx1F0u = (MText)trCBS1F0u.GetObject(soCBS1F0u.ObjectId, OpenMode.ForRead);
                    int txtColor1F0u = mtx1F0u.ColorIndex;
                    string txtCont1F0u = mtx1F0u.Contents;
                    if(txtColor1F0u == txtInfCol)
                    {
                        string patt1 = @"(?<=\:)\d+";
                        Regex regex1F0u = new Regex(patt1);
                        Match match1F0u = regex1F0u.Match(txtCont1F0u);
                        if (match1F0u.Success)
                        {
                            barDia1F0u = Convert.ToDouble(match1F0u.Value);
                        }

                        string patt2 = @"(?<=\-)\d+";
                        Regex regex2F0u = new Regex(patt2);
                        Match match2F0u = regex2F0u.Match(txtCont1F0u);
                        if (match2F0u.Success)
                        {
                            qtyTop1F0u = Convert.ToInt16(match2F0u.Value);
                        }

                        string patt3 = @"(?<=\/)\d+";
                        Regex regex3F0u = new Regex(patt3);
                        Match match3F0u = regex3F0u.Match(txtCont1F0u);
                        if (match3F0u.Success)
                        {
                            qtyBot1F0u = Convert.ToInt16(match3F0u.Value);
                        }
                    }

                    if (txtColor1F0u == txtDimCol)
                    {
                        string patt4 = @"(?<=\()\d+";
                        Regex regex4F0u = new Regex(patt4);
                        Match match4F0u = regex4F0u.Match(txtCont1F0u);
                        if (match4F0u.Success)
                        {
                            wid1F0u = Convert.ToDouble(match4F0u.Value);
                        }

                        string patt5 = @"(?<=[xX])\d+";
                        Regex regex5F0u = new Regex(patt5);
                        Match match5F0u = regex5F0u.Match(txtCont1F0u);
                        if (match5F0u.Success)
                        {
                            hgt1F0u = Convert.ToDouble(match5F0u.Value);
                            actDoc.Editor.WriteMessage("\nBeam height: {0}", hgt1F0u);
                        }
                    }
                }

                PromptPointOptions ppoCBS1F0u = new PromptPointOptions("");
                ppoCBS1F0u.Message = "\nEnter insertion point: ";
                ppoCBS1F0u.AllowNone = false;

                PromptPointResult pprCBS1F0u = actDoc.Editor.GetPoint(ppoCBS1F0u);
                if (pprCBS1F0u.Status != PromptStatus.OK) return;
                Point3d p3Pt1 = pprCBS1F0u.Value;

                //A. TOP BARS
                double sA1F0u = wid1F0u - 2 * concCov - 2 * strDia - barDia1F0u; //CTC spacing of outer bars
                double sB1F0u = (sA1F0u / (qtyTop1F0u - 1)) - barDia1F0u; //Calculated clear spacing of bars
                actDoc.Editor.WriteMessage("\nTop bar actual clear spacing: {0}", sB1F0u);
                double sC1F0u = sA1F0u / (qtyTop1F0u - 1); //Center spacing of bars

                double clearSpacing = 20.0;
                if(sA1F0u > 100)
                {
                    clearSpacing = 20.0 / 1000;
                }

                Point3d p3Pt2 = p3Pt1 + new Vector3d(sA1F0u, 0, 0);
                Vector3d v3P1P2 = p3Pt1.GetVectorTo(p3Pt2);
                Vector3d nv3P1P2 = v3P1P2.GetNormal();

                if (sB1F0u >= clearSpacing)//A.1 Arrangement 1
                {
                    for (int aa = 0; aa < qtyTop1F0u; ++aa)
                    {
                        Point3d p3aaPt = p3Pt1 + aa * sC1F0u * nv3P1P2;
                        using (Circle aaCircle = new Circle())
                        {
                            aaCircle.Center = p3aaPt;
                            aaCircle.Radius = barDia1F0u / 2;
                            aaCircle.ColorIndex = 2;
                            bltrec.AppendEntity(aaCircle);
                            trCBS1F0u.AddNewlyCreatedDBObject(aaCircle, true);
                        }
                    }
                }
                else //A.2 Arrangement 2 - with bundled bars
                {
                    int bundle = Convert.ToInt16((qtyTop1F0u - 2) / 2);
                    int nonBundle = (qtyTop1F0u - 2) % 2;
                    int nSpacing = bundle + nonBundle + 1;
                    double iSpacing = (sA1F0u-(2*barDia1F0u)*bundle-barDia1F0u*nonBundle - barDia1F0u)/nSpacing;
                    actDoc.Editor.WriteMessage("\nBundle: {0}, Nonbundle: {1}, N-Spacing: {2}, ith-spacing: {3}, sA: {4}", 
                        bundle, nonBundle, nSpacing, iSpacing,sA1F0u);

                    int countA = bundle + nonBundle + 2;
                    Point3d p3bbPt = p3Pt1;
                    for (int bb=0; bb < countA; ++bb)
                    {
                        if(bb == 0)
                        {
                            using (Circle aaCircle = new Circle())
                            {
                                aaCircle.Center = p3bbPt;
                                aaCircle.Radius = barDia1F0u / 2;
                                aaCircle.ColorIndex = 2;
                                bltrec.AppendEntity(aaCircle);
                                trCBS1F0u.AddNewlyCreatedDBObject(aaCircle, true);
                            }
                        }
                        else if(bb == 1)
                        {
                            p3bbPt = p3bbPt + (1.5*barDia1F0u + iSpacing)*nv3P1P2;
                            Point3d p3bbPta = p3bbPt + (- 0.5 * barDia1F0u) * nv3P1P2;
                            using (Circle aaCircle = new Circle())
                            {
                                aaCircle.Center = p3bbPta;
                                aaCircle.Radius = barDia1F0u / 2;
                                aaCircle.ColorIndex = 2;
                                bltrec.AppendEntity(aaCircle);
                                trCBS1F0u.AddNewlyCreatedDBObject(aaCircle, true);
                            }
                            Point3d p3bbPtb = p3bbPt + (0.5 * barDia1F0u) * nv3P1P2;
                            using (Circle aaCircle = new Circle())
                            {
                                aaCircle.Center = p3bbPtb;
                                aaCircle.Radius = barDia1F0u / 2;
                                aaCircle.ColorIndex = 2;
                                bltrec.AppendEntity(aaCircle);
                                trCBS1F0u.AddNewlyCreatedDBObject(aaCircle, true);
                            }
                        }
                        else if(bb > 1 && bb <= (countA - 2))
                        {
                            if (nonBundle > 0 && bb == ((countA-1) / 2))
                            {
                                p3bbPt = p3bbPt + (1.5 * barDia1F0u + iSpacing) * nv3P1P2;
                                using (Circle aaCircle = new Circle())
                                {
                                    aaCircle.Center = p3bbPt;
                                    aaCircle.Radius = barDia1F0u / 2;
                                    aaCircle.ColorIndex = 2;
                                    bltrec.AppendEntity(aaCircle);
                                    trCBS1F0u.AddNewlyCreatedDBObject(aaCircle, true);
                                }
                            }
                            else
                            {
                                if (nonBundle > 0 && bb == (((countA - 1) / 2) + 1))
                                {
                                    p3bbPt = p3bbPt + (1.5 * barDia1F0u + iSpacing) * nv3P1P2;
                                }
                                else
                                {
                                    p3bbPt = p3bbPt + (2 * barDia1F0u + iSpacing) * nv3P1P2;
                                }
                                  
                                Point3d p3bbPtc = p3bbPt + (-0.5 * barDia1F0u) * nv3P1P2;
                                using (Circle aaCircle = new Circle())
                                {
                                    aaCircle.Center = p3bbPtc;
                                    aaCircle.Radius = barDia1F0u / 2;
                                    aaCircle.ColorIndex = 2;
                                    bltrec.AppendEntity(aaCircle);
                                    trCBS1F0u.AddNewlyCreatedDBObject(aaCircle, true);
                                }
                                Point3d p3bbPtd = p3bbPt + (0.5 * barDia1F0u) * nv3P1P2;
                                using (Circle aaCircle = new Circle())
                                {
                                    aaCircle.Center = p3bbPtd;
                                    aaCircle.Radius = barDia1F0u / 2;
                                    aaCircle.ColorIndex = 2;
                                    bltrec.AppendEntity(aaCircle);
                                    trCBS1F0u.AddNewlyCreatedDBObject(aaCircle, true);
                                }
                            }
                        }
                        else if(bb == (countA - 1))
                        {
                            p3bbPt = p3bbPt + (1.5 * barDia1F0u + iSpacing) * nv3P1P2;
                            using (Circle aaCircle = new Circle())
                            {
                                aaCircle.Center = p3bbPt;
                                aaCircle.Radius = barDia1F0u / 2;
                                aaCircle.ColorIndex = 2;
                                bltrec.AppendEntity(aaCircle);
                                trCBS1F0u.AddNewlyCreatedDBObject(aaCircle, true);
                            }
                        }
                    }
                }


                //A. BOTTOM BARS
                double sA2F0u = wid1F0u - 2 * concCov - 2 * strDia - barDia1F0u; //CTC spacing of outer bars
                double sB2F0u = (sA2F0u / (qtyBot1F0u - 1)) - barDia1F0u; //Calculated clear spacing of bars
                actDoc.Editor.WriteMessage("\nBottom bar actual clear spacing: {0}", sB2F0u);
                double sC2F0u = sA2F0u / (qtyBot1F0u - 1); //Center spacing of bars
                double sH2F0u = hgt1F0u - 2 * concCov - 2 * strDia - barDia1F0u; //CTC spacing of outer bars


                Point3d p3BPt1 = p3Pt1 + new Vector3d(0, -sH2F0u, 0);
                Point3d p3BPt2 = p3BPt1 + new Vector3d(sA2F0u, 0, 0);
                Vector3d v3BP1P2 = p3BPt1.GetVectorTo(p3BPt2);
                Vector3d nv3BP1P2 = v3BP1P2.GetNormal();

                if (sB2F0u >= clearSpacing)//A.1 Arrangement 1
                {
                    for (int aa = 0; aa < qtyBot1F0u; ++aa)
                    {
                        Point3d p3aaPt = p3BPt1 + aa * sC2F0u * nv3BP1P2;
                        using (Circle aaCircle = new Circle())
                        {
                            aaCircle.Center = p3aaPt;
                            aaCircle.Radius = barDia1F0u / 2;
                            aaCircle.ColorIndex = 2;
                            bltrec.AppendEntity(aaCircle);
                            trCBS1F0u.AddNewlyCreatedDBObject(aaCircle, true);
                        }
                    }
                }
                else //A.2 Arrangement 2 - with bundled bars
                {
                    int bundle = Convert.ToInt16((qtyBot1F0u - 2) / 2);
                    int nonBundle = (qtyBot1F0u - 2) % 2;
                    int nSpacing = bundle + nonBundle + 1;
                    double iSpacing = (sA2F0u - (2 * barDia1F0u) * bundle - barDia1F0u * nonBundle - barDia1F0u) / nSpacing;
                    actDoc.Editor.WriteMessage("\nBottom bars: Bundle: {0}, Nonbundle: {1}, N-Spacing: {2}, ith-spacing: {3}, sA: {4}",
                        bundle, nonBundle, nSpacing, iSpacing, sA1F0u);

                    int countA = bundle + nonBundle + 2;
                    Point3d p3bbPt = p3BPt1;
                    for (int bb = 0; bb < countA; ++bb)
                    {
                        if (bb == 0)
                        {
                            using (Circle aaCircle = new Circle())
                            {
                                aaCircle.Center = p3bbPt;
                                aaCircle.Radius = barDia1F0u / 2;
                                aaCircle.ColorIndex = 2;
                                bltrec.AppendEntity(aaCircle);
                                trCBS1F0u.AddNewlyCreatedDBObject(aaCircle, true);
                            }
                        }
                        else if (bb == 1)
                        {
                            p3bbPt = p3bbPt + (1.5 * barDia1F0u + iSpacing) * nv3P1P2;
                            Point3d p3bbPta = p3bbPt + (-0.5 * barDia1F0u) * nv3P1P2;
                            using (Circle aaCircle = new Circle())
                            {
                                aaCircle.Center = p3bbPta;
                                aaCircle.Radius = barDia1F0u / 2;
                                aaCircle.ColorIndex = 2;
                                bltrec.AppendEntity(aaCircle);
                                trCBS1F0u.AddNewlyCreatedDBObject(aaCircle, true);
                            }
                            Point3d p3bbPtb = p3bbPt + (0.5 * barDia1F0u) * nv3P1P2;
                            using (Circle aaCircle = new Circle())
                            {
                                aaCircle.Center = p3bbPtb;
                                aaCircle.Radius = barDia1F0u / 2;
                                aaCircle.ColorIndex = 2;
                                bltrec.AppendEntity(aaCircle);
                                trCBS1F0u.AddNewlyCreatedDBObject(aaCircle, true);
                            }
                        }
                        else if (bb > 1 && bb <= (countA - 2))
                        {
                            if (nonBundle > 0 && bb == ((countA - 1) / 2))
                            {
                                p3bbPt = p3bbPt + (1.5 * barDia1F0u + iSpacing) * nv3P1P2;
                                using (Circle aaCircle = new Circle())
                                {
                                    aaCircle.Center = p3bbPt;
                                    aaCircle.Radius = barDia1F0u / 2;
                                    aaCircle.ColorIndex = 2;
                                    bltrec.AppendEntity(aaCircle);
                                    trCBS1F0u.AddNewlyCreatedDBObject(aaCircle, true);
                                }
                            }
                            else
                            {
                                if (nonBundle > 0 && bb == (((countA - 1) / 2) + 1))
                                {
                                    p3bbPt = p3bbPt + (1.5 * barDia1F0u + iSpacing) * nv3P1P2;
                                }
                                else
                                {
                                    p3bbPt = p3bbPt + (2 * barDia1F0u + iSpacing) * nv3P1P2;
                                }

                                Point3d p3bbPtc = p3bbPt + (-0.5 * barDia1F0u) * nv3P1P2;
                                using (Circle aaCircle = new Circle())
                                {
                                    aaCircle.Center = p3bbPtc;
                                    aaCircle.Radius = barDia1F0u / 2;
                                    aaCircle.ColorIndex = 2;
                                    bltrec.AppendEntity(aaCircle);
                                    trCBS1F0u.AddNewlyCreatedDBObject(aaCircle, true);
                                }
                                Point3d p3bbPtd = p3bbPt + (0.5 * barDia1F0u) * nv3P1P2;
                                using (Circle aaCircle = new Circle())
                                {
                                    aaCircle.Center = p3bbPtd;
                                    aaCircle.Radius = barDia1F0u / 2;
                                    aaCircle.ColorIndex = 2;
                                    bltrec.AppendEntity(aaCircle);
                                    trCBS1F0u.AddNewlyCreatedDBObject(aaCircle, true);
                                }
                            }
                        }
                        else if (bb == (countA - 1))
                        {
                            p3bbPt = p3bbPt + (1.5 * barDia1F0u + iSpacing) * nv3P1P2;
                            using (Circle aaCircle = new Circle())
                            {
                                aaCircle.Center = p3bbPt;
                                aaCircle.Radius = barDia1F0u / 2;
                                aaCircle.ColorIndex = 2;
                                bltrec.AppendEntity(aaCircle);
                                trCBS1F0u.AddNewlyCreatedDBObject(aaCircle, true);
                            }
                        }
                    }
                }


                Point2d p2RecPt1 = new Point2d(p3Pt1.X,p3Pt1.Y) + new Vector2d(-(0.5 * barDia1F0u + strDia + concCov), (0.5 * barDia1F0u + strDia + concCov));
                Point2d p2RecPt2 = p2RecPt1 + new Vector2d(wid1F0u, 0);
                Point2d p2RecPt3 = p2RecPt2 + new Vector2d(0, -hgt1F0u);
                Point2d p2RecPt4 = p2RecPt3 + new Vector2d(-wid1F0u, 0);

                using(Polyline plObj = new Polyline())
                {
                    plObj.AddVertexAt(0, p2RecPt1, 0, 0, 0);
                    plObj.AddVertexAt(1, p2RecPt2, 0, 0, 0);
                    plObj.AddVertexAt(2, p2RecPt3, 0, 0, 0);
                    plObj.AddVertexAt(3, p2RecPt4, 0, 0, 0);
                    plObj.AddVertexAt(4, p2RecPt1, 0, 0, 0);
                    plObj.ColorIndex = 1;

                    bltrec.AppendEntity(plObj);
                    trCBS1F0u.AddNewlyCreatedDBObject(plObj, true);
                }

                trCBS1F0u.Commit();
            }
        }



        /**
         * Date added: 27 May 2024
         * Added by: Bernardo A. Cabebe Jr.
         * Venue: 3407 Cityland Pasong Tamo Tower
         * Status: Under development
         */
        [CommandMethod("TCG_TabulateCuttingListGoogleSheet")]
        public static void TabulateCuttingListGoogleSheet()
        {
            if (!ccData.isLicenseActive())
            {
                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            //Get the leaders
            TypedValue[] tvFilterG3w = new TypedValue[2];
            tvFilterG3w[0] = new TypedValue((int)DxfCode.Start, "LEADER");
            tvFilterG3w[1] = new TypedValue((int)DxfCode.Color, 251);
            SelectionFilter sfTCGG3w = new SelectionFilter(tvFilterG3w);

            PromptSelectionOptions psoTCGG3w = new PromptSelectionOptions();
            psoTCGG3w.MessageForAdding = "\nSelect leader: ";

            PromptSelectionResult psrTCGG3w = actDoc.Editor.GetSelection(psoTCGG3w, sfTCGG3w);
            if (psrTCGG3w.Status != PromptStatus.OK) return;

            SelectionSet ssTCGG3w = psrTCGG3w.Value;


            PromptStringOptions psoTCGG4y = new PromptStringOptions("");
            psoTCGG4y.Message = "\nEnter Beam Mark:";
            psoTCGG4y.DefaultValue = bbsBeamMarkG;

            PromptResult prTCGG4y = actDoc.Editor.GetString(psoTCGG4y);
            if (prTCGG4y.Status != PromptStatus.OK) return;
            string bbsBeamMark = prTCGG4y.StringResult;
            bbsBeamMarkG = bbsBeamMark;


            PromptIntegerOptions pioTCGG3w = new PromptIntegerOptions("");
            pioTCGG3w.Message = "\nEnter Cutting List GS row: ";
            pioTCGG3w.DefaultValue = gsRowG;
            pioTCGG3w.AllowNone = true;
            pioTCGG3w.AllowNegative = false;
            pioTCGG3w.AllowZero = false;

            PromptIntegerResult pirTCGG3w = actDoc.Editor.GetInteger(pioTCGG3w);
            if (pirTCGG3w.Status != PromptStatus.OK) return;
            int gsRowG3w = pirTCGG3w.Value;
            gsRowG = gsRowG3w;

            PromptIntegerOptions pioTCGG4w = new PromptIntegerOptions("");
            pioTCGG4w.Message = "\nEnter BBS GS row: ";
            pioTCGG4w.DefaultValue = gsRowBBSG;
            pioTCGG4w.AllowNone = true;
            pioTCGG4w.AllowNegative = false;
            pioTCGG4w.AllowZero = false;

            PromptIntegerResult pirTCGG4w = actDoc.Editor.GetInteger(pioTCGG4w);
            if (pirTCGG4w.Status != PromptStatus.OK) return;
            int gsRowBBSG4w = pirTCGG4w.Value;
            gsRowBBSG = gsRowBBSG4w;

            PromptStringOptions psoBBSTab = new PromptStringOptions("");
            psoBBSTab.DefaultValue = tabBBSG;
            psoBBSTab.Message = "\nEnter BBS Tab name: ";
            psoBBSTab.AllowSpaces = false;

            PromptResult psrBBSTab = actDoc.Editor.GetString(psoBBSTab);
            if (psrBBSTab.Status != PromptStatus.OK) return;
            string bbsTab = psrBBSTab.StringResult;
            tabBBSG = bbsTab;

            actDoc.Editor.WriteMessage("\nCopy google sheet template from: https://docs.google.com/spreadsheets/d/1WdD0AdxEHthbzz4g0CkAXNs_ywVOnjXYaOIUmXfEsOg/view#gid=0");
            PromptStringOptions psoTCGG4w = new PromptStringOptions("");
            psoTCGG4w.Message = "\nEnter Google Sheet ID:";
            psoTCGG4w.DefaultValue = clGSheetIDG;

            PromptResult prTCGG4w = actDoc.Editor.GetString(psoTCGG4w);
            if(prTCGG4w.Status != PromptStatus.OK) return;
            string clGSheetID = prTCGG4w.StringResult;
            clGSheetIDG = clGSheetID;


            actDoc.Editor.WriteMessage("\nCopy google sheet BBS template from: https://docs.google.com/spreadsheets/d/1xmapL82ZIP0CPuvVMQmPWKPLXpsyzOANs6x5gSyPdFQ/edit#gid=1657268530");
            PromptStringOptions psoTCGG4x = new PromptStringOptions("");
            psoTCGG4x.Message = "\nEnter BBS Google Sheet ID:";
            psoTCGG4x.DefaultValue = bbsGSheetIDG;

            PromptResult prTCGG4x = actDoc.Editor.GetString(psoTCGG4x);
            if (prTCGG4x.Status != PromptStatus.OK) return;
            string bbsGSheetID = prTCGG4x.StringResult;
            bbsGSheetIDG = bbsGSheetID;



            TypedValue[] tvFilterG4w = new TypedValue[9];
            tvFilterG4w[0] = new TypedValue((int)DxfCode.Start, "LWPOLYLINE");
            tvFilterG4w[1] = new TypedValue((int)DxfCode.Operator, "<OR");
            tvFilterG4w[2] = new TypedValue((int)DxfCode.Color, botBarCol);
            tvFilterG4w[3] = new TypedValue((int)DxfCode.Color, botExtCol);
            tvFilterG4w[4] = new TypedValue((int)DxfCode.Color, countedWebLineColor);
            tvFilterG4w[5] = new TypedValue((int)DxfCode.Color, topExtCol);
            tvFilterG4w[6] = new TypedValue((int)DxfCode.Color, topBarCol);
            tvFilterG4w[7] = new TypedValue((int)DxfCode.Color, stirrupSectionColor);
            tvFilterG4w[8] = new TypedValue((int)DxfCode.Operator, "OR>");
            SelectionFilter sfTCGG4w = new SelectionFilter(tvFilterG4w);

            TypedValue[] tvFilterG5w = new TypedValue[5];
            tvFilterG5w[0] = new TypedValue((int)DxfCode.Start, "MTEXT");
            tvFilterG5w[1] = new TypedValue((int)DxfCode.Operator, "<OR");
            tvFilterG5w[2] = new TypedValue((int)DxfCode.Color, labelColor);
            tvFilterG5w[3] = new TypedValue((int)DxfCode.Color, labelCodeColor);
            tvFilterG5w[4] = new TypedValue((int)DxfCode.Operator, "OR>");
            SelectionFilter sfTCGG5w = new SelectionFilter(tvFilterG5w);

            using (Transaction trTCGG3w = aCurDB.TransactionManager.StartTransaction())
            {
                //Get all rebars
                List<KeyValuePair<string,ObjectId>> liStirrG3w  = new List<KeyValuePair<string,ObjectId>>();  //Get all stirrups and store in a list
                List<KeyValuePair<string,ObjectId>> liContBBG3w = new List<KeyValuePair<string,ObjectId>>(); //Get all Continuous bars and store in a list
                List<KeyValuePair<string,ObjectId>> liExtrBBG3w = new List<KeyValuePair<string,ObjectId>>(); //Get all Extra Bottom bars and store in a list
                List<KeyValuePair<string,ObjectId>> liWebBarG3w = new List<KeyValuePair<string,ObjectId>>(); //Get all web bars and store in a list
                List<KeyValuePair<string,ObjectId>> liContTBG3w = new List<KeyValuePair<string,ObjectId>>(); //Get all continous top bars and store in a list
                List<KeyValuePair<string,ObjectId>> liExtrTBG3w = new List<KeyValuePair<string,ObjectId>>(); //Get all extra top bars and store in a list
                foreach(SelectedObject soTCGG3w in ssTCGG3w)
                {
                    Leader soLeader = trTCGG3w.GetObject(soTCGG3w.ObjectId, OpenMode.ForRead) as Leader;
                    Entity enLeader = trTCGG3w.GetObject(soTCGG3w.ObjectId, OpenMode.ForRead) as Entity;
                    string dxfName = enLeader.GetRXClass().DxfName;

                    Point3d p3VtxP1 = soLeader.VertexAt(0);
                    Point3d p3VtxP2 = soLeader.VertexAt(1);
                    Point3d p3VtxP3 = soLeader.VertexAt(2);

                    Vector3d v3P21 = p3VtxP2.GetVectorTo(p3VtxP1);
                    Vector3d nv3P21 = v3P21.GetNormal();
                    Vector3d v3P23 = p3VtxP2.GetVectorTo(p3VtxP3);
                    Vector3d nv3P23 = v3P23.GetNormal();

                    double distY6u = p3VtxP1.DistanceTo(p3VtxP2);
                    double len1Y6u = 10;
                    double len2Y6u = 2000;
                    if (distY6u < 50)
                    {
                        len1Y6u = 0.1;
                        len2Y6u = 2.0;
                    }
                    Point3d p3VtxP1A = p3VtxP1 - len1Y6u * nv3P21;
                    Point3d p3VtxP1B = p3VtxP1 + len1Y6u * nv3P21;
                    Point3d p3VtxP2A = p3VtxP3;
                    Point3d p3VtxP2B = p3VtxP3 + len2Y6u * nv3P23;

                    Point3dCollection p3cFence1 = new Point3dCollection();
                    p3cFence1.Add(p3VtxP1A);
                    p3cFence1.Add(p3VtxP1B);

                    Point3dCollection p3cFence2 = new Point3dCollection();
                    p3cFence2.Add(p3VtxP2A);
                    p3cFence2.Add(p3VtxP2B);

                    

                    PromptSelectionResult psrMtx = actDoc.Editor.SelectFence(p3cFence2, sfTCGG5w);      
                    //PromptSelectionResult psrMtx = actDoc.Editor.SelectCrossingWindow(p3VtxP2A, p3VtxP2B);
                    SelectionSet ssMTx = psrMtx.Value;
                    actDoc.Editor.WriteMessage("\nLabels: {0}", ssMTx.Count);
                    if(ssMTx.Count < 2)
                    {
                        actDoc.Editor.WriteMessage("\nInvalid number of labels. Count {0}. Exit...", ssMTx.Count);
                        enLeader.Highlight();
                        return;
                    }
                    
                    string txtCont = "";
                    string txtCont2 = "";
                    if (ssMTx.Count > 0)
                    {
                        MText mtxObj = trTCGG3w.GetObject(ssMTx[0].ObjectId, OpenMode.ForRead) as MText;
                        int mtxObjCol = mtxObj.ColorIndex;

                        MText mtxObj2 = trTCGG3w.GetObject(ssMTx[1].ObjectId, OpenMode.ForRead) as MText;
                        int mtxObj2Col = mtxObj2.ColorIndex;
                        if(mtxObjCol == 80 && mtxObj2Col == 181)
                        {
                            txtCont = mtxObj.Contents;
                            txtCont2 = mtxObj2.Contents;
                        }
                        else
                        {
                            txtCont = mtxObj2.Contents;
                            txtCont2 = mtxObj.Contents;
                        }
                        actDoc.Editor.WriteMessage("\nBar info: {0}", txtCont);
                        actDoc.Editor.WriteMessage("\nBar code: {0}", txtCont2);
                    }
                    else
                    {
                        actDoc.Editor.WriteMessage("\nEntity {0} cannot select text information.");
                        enLeader.Highlight();
                        txtCont = "Empty";
                    }


                    PromptSelectionResult psrBar = actDoc.Editor.SelectFence(p3cFence1, sfTCGG4w);
                    //PromptSelectionResult psrBar = actDoc.Editor.SelectCrossingWindow(p3VtxP1A, p3VtxP1B);
                    SelectionSet ssBar = psrBar.Value;
                    if(ssBar.Count < 1)
                    {
                        actDoc.Editor.WriteMessage("\nInvalid number of bars. Count {0}. Exit...", ssBar.Count);
                        enLeader.Highlight();
                        return;
                    }
                    Polyline plBar = new Polyline();
                    double barLen = 0;
                    string strBarMessage = "";
                    ObjectId oiPolBar = new ObjectId();
                    if (ssBar.Count > 0)
                    {
                        plBar = trTCGG3w.GetObject(ssBar[0].ObjectId, OpenMode.ForRead) as Polyline;
                        oiPolBar = ssBar[0].ObjectId;
                        barLen = Math.Round(plBar.Length, 2);
                        strBarMessage = "Bar length";
                    }
                    else
                    {
                        strBarMessage = "Entity cannot select bar line.";
                        enLeader.Highlight();
                        barLen = 0;
                    }
                    actDoc.Editor.WriteMessage("\n{0}: {1}",strBarMessage, barLen);


                    //*****
                    string mtxContG3w = txtCont;
                    //actDoc.Editor.WriteMessage("\nCL info: {0}", mtxContG3w);

                    string pattStirr = @"^STI.*";
                    Regex regExStirr = new Regex(pattStirr);
                    Match matchStirr = regExStirr.Match(mtxContG3w);
                    if(matchStirr.Success)
                    {
                        liStirrG3w.Add(new KeyValuePair<string, ObjectId>(matchStirr.Value + ";;" + txtCont2, oiPolBar));
                    }

                    string pattContBB = @"^CBB.*";
                    Regex regExContBB = new Regex(pattContBB);
                    Match matchContBB = regExContBB.Match(mtxContG3w);
                    if(matchContBB.Success)
                    {
                        liContBBG3w.Add(new KeyValuePair<string, ObjectId>(matchContBB.Value + ";;" + txtCont2,oiPolBar));
                    }

                    string pattExtrBB = @"^EBB.*";
                    Regex regExExtrBB = new Regex(pattExtrBB);
                    Match matchExtrBB = regExExtrBB.Match(mtxContG3w);
                    if(matchExtrBB.Success)
                    {
                        liExtrBBG3w.Add(new KeyValuePair<string, ObjectId>(matchExtrBB.Value + ";;" + txtCont2,oiPolBar));
                    }

                    string pattWebBar = @"^WEB.*";
                    Regex regExWebBar = new Regex(pattWebBar);
                    Match matchWebBar = regExWebBar.Match(mtxContG3w);
                    if(matchWebBar.Success)
                    {
                        liWebBarG3w.Add(new KeyValuePair<string, ObjectId>(matchWebBar.Value + ";;" + txtCont2, oiPolBar));
                    }

                    string pattContTB = @"^CTB.*";
                    Regex regExContTB = new Regex(pattContTB);
                    Match matchContTB = regExContTB.Match(mtxContG3w);
                    if(matchContTB.Success)
                    {
                        liContTBG3w.Add(new KeyValuePair<string, ObjectId>(matchContTB.Value + ";;" + txtCont2, oiPolBar));
                    }

                    string pattExtrTB = @"^ETB.*";
                    Regex regExExtrTB = new Regex(pattExtrTB);
                    Match matchExtrTB = regExExtrTB.Match(mtxContG3w);
                    if (matchExtrTB.Success)
                    {
                        liExtrTBG3w.Add(new KeyValuePair<string, ObjectId>(matchExtrTB.Value + ";;" + txtCont2, oiPolBar));
                    }
                }

                

                liStirrG3w.OrderBy(ii => ii); //Ascending
                //liStirrG3w.OrderByDescending(ii => ii); //Descending
                liContBBG3w.Sort((pair1, pair2) => pair1.Key.CompareTo(pair2.Key)); //.OrderBy(ii => ii);
                liExtrBBG3w.Sort((pair1, pair2) => pair1.Key.CompareTo(pair2.Key)); //.OrderBy(ii => ii);
                liWebBarG3w.Sort((pair1, pair2) => pair1.Key.CompareTo(pair2.Key)); //.OrderBy(ii => ii);
                liContTBG3w.Sort((pair1, pair2) => pair1.Key.CompareTo(pair2.Key)); //.OrderBy(ii => ii);
                liExtrTBG3w.Sort((pair1, pair2) => pair1.Key.CompareTo(pair2.Key)); //.OrderBy(ii => ii);

                int liSize = liStirrG3w.Count + liContBBG3w.Count + liExtrBBG3w.Count + liWebBarG3w.Count + liContTBG3w.Count + liExtrTBG3w.Count;
                List<KeyValuePair<string, ObjectId>> liCutBars = new List<KeyValuePair<string, ObjectId>>(liSize);
                liCutBars.AddRange(liStirrG3w);
                liCutBars.AddRange(liContBBG3w);
                liCutBars.AddRange(liExtrBBG3w);
                liCutBars.AddRange(liWebBarG3w);
                liCutBars.AddRange(liContTBG3w);
                liCutBars.AddRange(liExtrTBG3w);

                BarBendingScheduleFu(liCutBars, bbsGSheetID,gsRowBBSG4w,bbsTab,bbsBeamMark);

                int gsRowCounter = 0;
                IList<IList<object>> iioCutBars = new List<IList<object>>();
                for(int aa=0; aa<liCutBars.Count; aa++)
                {
                    actDoc.Editor.WriteMessage("\nBar info: {0}, ID: {1}", liCutBars[aa].Key, liCutBars[aa].Value);
                    IList<object> ioCutBars = new List<object>();

                    ioCutBars.Add("=ROW()-2");

                    string pattBarID = @"[A-Z0-9]+(?=\:)";
                    Regex regExBarID = new Regex(pattBarID);
                    Match matchBarID = regExBarID.Match(liCutBars[aa].Key);
                    if (matchBarID.Success)
                    {
                        ioCutBars.Add(matchBarID.Value);
                        //actDoc.Editor.WriteMessage("\nBar ID: {0}", matchBarID.Value);
                    }


                    //Description
                    string strDesc = "";
                    string patDesc = @"^[A-Z]{3}";
                    Regex regExDes = new Regex(patDesc);
                    Match matchDes = regExDes.Match(liCutBars[aa].Key);
                    if (matchDes.Success)
                    {
                        switch (matchDes.Value)
                        {
                            case "STI":
                                strDesc = "STIRRUP";
                                ioCutBars.Add(strDesc);
                                break;
                            case "CBB":
                                strDesc = "CONTINUOUS BOT. BAR";
                                ioCutBars.Add(strDesc);
                                break;
                            case "EBB":
                                strDesc = "EXTRA BOT. BAR";
                                ioCutBars.Add(strDesc);
                                break;
                            case "WEB":
                                strDesc = "WEB BAR";
                                ioCutBars.Add(strDesc);
                                break;
                            case "CTB":
                                strDesc = "CONTINUOUS TOP BAR";
                                ioCutBars.Add(strDesc);
                                break;
                            case "ETB":
                                strDesc = "EXTRA TOP BAR";
                                ioCutBars.Add(strDesc);
                                break;
                            default:
                                strDesc = "-*-";
                                ioCutBars.Add(strDesc);
                                break;
                        }
                        //actDoc.Editor.WriteMessage("\nDescription: {0}", strDesc);
                    }

                    //Bar diameter
                    string pattDiam = @"(?<=\-)\s*\d+";
                    Regex regExDiam = new Regex(pattDiam);
                    Match matchDiam = regExDiam.Match(liCutBars[aa].Key);
                    if (matchDiam.Success)
                    {
                        string strValG3w = matchDiam.Value;
                        ioCutBars.Add(strValG3w);
                        //actDoc.Editor.WriteMessage("\nBar diameter: {0}", strValG3w);
                    }

                    //Bar quantity
                    string pattQty = @"(?<=:)\s*\d+";
                    Regex regExQty = new Regex(pattQty);
                    Match matchQty = regExQty.Match(liCutBars[aa].Key);
                    if (matchQty.Success)
                    {
                        string strValG3w = matchQty.Value;
                        ioCutBars.Add(strValG3w);
                        //actDoc.Editor.WriteMessage("\nBar quantity: {0}", strValG3w);
                    }

                    //Bar length
                    string pattLen = @"(?<=[xX])\s*[0-9\.]+";
                    Regex regExLen = new Regex(pattLen);
                    Match matchLen = regExLen.Match(liCutBars[aa].Key);
                    if (matchLen.Success)
                    {
                        string strValG3w = matchLen.Value;
                        ioCutBars.Add(strValG3w);
                        //actDoc.Editor.WriteMessage("\nBar length: {0}", strValG3w);
                    }

                    ioCutBars.Add("=if(INDIRECT(\"F\"&ROW())>=6,index(CLData!$F$2:$F$6,match(0,ARRAYFORMULA((INDIRECT(\"F\"&ROW())>CLData!$F$2:$F$6)*1),0)),index(CLData!$A$2:$A$20,match(0,ARRAYFORMULA((INDIRECT(\"F\"&ROW())>CLData!$A$2:$A$20)*1),0)))"); //SEG. L
                    ioCutBars.Add("=round(index(CLData!$G$10:$G$17,match(1,ARRAYFORMULA((INDIRECT(\"D\"&ROW())=CLData!$F$10:$F$17)*1),0))*INDIRECT(\"F\"&ROW()),2)"); //WT. KGS
                    ioCutBars.Add("=round(index(CLData!$G$10:$G$17,match(1,ARRAYFORMULA((INDIRECT(\"D\"&ROW())=CLData!$F$10:$F$17)*1),0))*(INDIRECT(\"G\"&ROW())-INDIRECT(\"F\"&ROW())),2)"); //EXC. WT.
                    ioCutBars.Add(bbsBeamMark); //SEG. L

                    iioCutBars.Add(ioCutBars);
                    gsRowCounter++;
                }


                if (gsRowG3w < 3)
                {
                    gsRowG3w = 3;
                }
                //actDoc.Editor.WriteMessage("\nList: {0}", iioCutBars.Count);
                //string tempShtId = "1WdD0AdxEHthbzz4g0CkAXNs_ywVOnjXYaOIUmXfEsOg";
                string strRangeAddrG3w = "CLTable!A" + gsRowG3w.ToString();
                ValueRange vrRngValG3w = new ValueRange();
                vrRngValG3w.MajorDimension = "ROWS";
                vrRngValG3w.Values = iioCutBars;
                SpreadsheetsResource.ValuesResource.UpdateRequest urRngValG3w = DataGlobal.sheetsService.Spreadsheets
                    .Values.Update(vrRngValG3w, clGSheetID, strRangeAddrG3w);
                urRngValG3w.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                UpdateValuesResponse uvrRngValG3w = urRngValG3w.Execute();
                //actDoc.Editor.WriteMessage("\nGSheet response: {0}", uvrRngValG3w);

                gsRowG3w = gsRowG3w + gsRowCounter;
                gsRowG = gsRowG3w;
                trTCGG3w.Commit();
            }
        }

        private static void BarBendingScheduleFu(List<KeyValuePair<string, ObjectId>> barInfOId, string gSheetID, int ithRow, string tabName, string beamName)
        {
            string[] keyVal =
            {
                "IdA1",
                "CodeB2",
                "GradeC3",
                "DiaD4",
                "DiaE5",
                "WtF6",
                "UnitG7",
                "LenH8",
                "MemI9",
                "QtyJ10",
                "QtyTK11",
                "AL12",
                "BM13",
                "CN14",
                "DO15",
                "EP16",
                "FQ17",
                "GR18",
                "HS19",
                "IT20",
                "JU21",
                "KV22",
                "LW23",
                "MX24",
                "NY25",
                "OZ26",
                "PAA27",
                "QAB28",
                "RemarkAC29"
            };

            IList<IList<object>> iioGSValue = new List<IList<object>>();
            int bbsRowCounter = 0;
            for (int aa = 0; aa < barInfOId.Count; aa++) {
                List<KeyValuePair<string, object>> lkvBBS = new List<KeyValuePair<string, object>>();
                object GetValueByKey(string key)
                {
                    var kvp = lkvBBS.Find(pair => pair.Key == key);
                    return kvp.Equals(default(KeyValuePair<string, object>)) ? null : kvp.Value;
                }

                string barInfo = barInfOId[aa].Key;
                ObjectId objectId = barInfOId[aa].Value;
                //actDoc.Editor.WriteMessage("\nKey: {0}, Value: {1}",barInfo, objectId);
                using (Transaction trBBSL3s = aCurDB.TransactionManager.StartTransaction())
                {
                    string pattId = @"^[A-Z0-9]+(?=\:)";
                    Regex regexID = new Regex(pattId);
                    Match matchID = regexID.Match(barInfo);
                    if (matchID.Success)
                    {
                        lkvBBS.Add(new KeyValuePair<string, object>("IdA1", matchID.Value)); //The Key is composed of description (Id) and column name/number pair (A1)
                    }

                    string pattCode = @"(?<=\;\;).+";
                    Regex regexCode = new Regex(pattCode);
                    Match matchCode = regexCode.Match(barInfo);
                    string code = "-*-";
                    if (matchCode.Success)
                    {
                        lkvBBS.Add(new KeyValuePair<string, object>("CodeB2", matchCode.Value)); //The Key is composed of description (Code) and column name/number pair (B2)
                        code = matchCode.Value;
                    }

                    lkvBBS.Add(new KeyValuePair<string, object>("GradeC3", 60));

                    string pattDia = @"(?<=\-)\s?\d+";
                    Regex regexDia = new Regex(pattDia);
                    Match matchDia = regexDia.Match(barInfo);
                    if (matchDia.Success)
                    {
                        lkvBBS.Add(new KeyValuePair<string, object>("DiaD4", matchDia.Value));
                    }
                    
                    lkvBBS.Add(new KeyValuePair<string, object>("DiaE5", ""));
                    lkvBBS.Add(new KeyValuePair<string, object>("WtF6", "=IFERROR(indirect(\"G\" & ROW())*indirect(\"H\" & ROW())*indirect(\"K\" & ROW()),\"-\")"));
                    lkvBBS.Add(new KeyValuePair<string, object>("UnitG7", "=IFERROR(VLOOKUP(indirect(\"D\" & ROW()),source_data!$A$1:$C$15,2,TRUE),\"\")"));

                    Polyline plObj = trBBSL3s.GetObject(objectId, OpenMode.ForRead) as Polyline;
                    double plObjLen = plObj.Length;
                    if(plObjLen > 100)
                    {
                        plObjLen = plObjLen / 1000;
                    }
                    lkvBBS.Add(new KeyValuePair<string, object>("LenH8", Math.Round(plObjLen,2)));

                    lkvBBS.Add(new KeyValuePair<string, object>("MemI9", 1));

                    string pattQty = @"\d+\s?(?=\-)";
                    Regex regexQty = new Regex(pattQty);
                    Match matchQty = regexQty.Match(barInfo);
                    if (matchQty.Success)
                    {
                        lkvBBS.Add(new KeyValuePair<string, object>("QtyJ10", matchQty.Value));
                    }

                    lkvBBS.Add(new KeyValuePair<string, object>("QtyTK11", "=indirect(\"I\" & ROW())*indirect(\"J\" & ROW())"));

                    switch (code)
                    {
                        case "CODE:T1": //Stirrup
                            for (int bb = plObj.NumberOfVertices - 1; bb >= 0; bb--)
                            {
                                switch (bb)
                                {
                                    case 12:
                                        lkvBBS.Add(new KeyValuePair<string, object>("AL12", plObj.GetLineSegmentAt(bb).Length));
                                        break;
                                    //case 9:
                                    //    lkvBBS.Add(new KeyValuePair<string, object>("BM13", plObj.GetLineSegmentAt(bb).Length));
                                    //    break;
                                    case 8:
                                        Point3d p3dVtx8 = plObj.GetPoint3dAt(bb);
                                        Point3d p3dVtx11 = plObj.GetPoint3dAt(bb + 3);
                                        double dist811 = p3dVtx8.DistanceTo(p3dVtx11);
                                        lkvBBS.Add(new KeyValuePair<string, object>("BM13", dist811));
                                        break;
                                    //case 7:
                                    //    lkvBBS.Add(new KeyValuePair<string, object>("CN14", plObj.GetLineSegmentAt(bb).Length));
                                    //    break;
                                    case 6:
                                        Point3d p3dVtx6 = plObj.GetPoint3dAt(bb);
                                        Point3d p3dVtx9 = plObj.GetPoint3dAt(bb + 3);
                                        double dist69 = p3dVtx6.DistanceTo(p3dVtx9);
                                        lkvBBS.Add(new KeyValuePair<string, object>("CN14", dist69));
                                        break;
                                    //case 5:
                                    //    lkvBBS.Add(new KeyValuePair<string, object>("DO15", plObj.GetLineSegmentAt(bb).Length));
                                    //    break;
                                    case 4:
                                        Point3d p3dVtx4 = plObj.GetPoint3dAt(bb);
                                        Point3d p3dVtx7 = plObj.GetPoint3dAt(bb + 3);
                                        double dist47 = p3dVtx4.DistanceTo(p3dVtx7);
                                        lkvBBS.Add(new KeyValuePair<string, object>("DO15",dist47));
                                        break;
                                    //case 3:
                                    //    lkvBBS.Add(new KeyValuePair<string, object>("EP16", plObj.GetLineSegmentAt(bb).Length));
                                    //    break;
                                    case 2:
                                        Point3d p3dVtx2 = plObj.GetPoint3dAt(bb);
                                        Point3d p3dVtx5 = plObj.GetPoint3dAt(bb + 3);
                                        double dist25 = p3dVtx2.DistanceTo(p3dVtx5);
                                        lkvBBS.Add(new KeyValuePair<string, object>("EP16", dist25));
                                        break;
                                    case 0:
                                        lkvBBS.Add(new KeyValuePair<string, object>("FQ17", plObj.GetLineSegmentAt(bb).Length));
                                        break;
                                    default:
                                        //None
                                        break;
                                }
                            }
                            break;
                        case "CODE:101":
                            double segL1 = plObj.GetLineSegmentAt(0).Length;
                            double segL2 = plObj.GetLineSegmentAt(1).Length;
                            if (segL1 < segL2)
                            {
                                lkvBBS.Add(new KeyValuePair<string, object>("AL12", segL1));
                                lkvBBS.Add(new KeyValuePair<string, object>("BM13", segL2));
                            }
                            else
                            {
                                lkvBBS.Add(new KeyValuePair<string, object>("AL12", segL2));
                                lkvBBS.Add(new KeyValuePair<string, object>("BM13", segL1));
                            }
                            break;
                        case "CODE:000":
                            double segLen = plObj.GetLineSegmentAt(0).Length;
                            lkvBBS.Add(new KeyValuePair<string, object>("AL12", segLen));
                            break;
                        case "CODE:201":
                            for (int bb = 0; bb < plObj.NumberOfVertices - 1; bb++)
                            {
                                switch (bb)
                                {
                                    case 0:
                                        lkvBBS.Add(new KeyValuePair<string, object>("AL12", plObj.GetLineSegmentAt(bb).Length));
                                        break;
                                    case 1:
                                        lkvBBS.Add(new KeyValuePair<string, object>("BM13", plObj.GetLineSegmentAt(bb).Length));
                                        break;
                                    case 2:
                                        lkvBBS.Add(new KeyValuePair<string, object>("CN14", plObj.GetLineSegmentAt(bb).Length));
                                        break;
                                    default:
                                        //None
                                        break;
                                }
                            }
                            break;
                        case "CODE:202":
                            for (int bb = 0; bb < plObj.NumberOfVertices - 1; bb++)
                            {
                                switch (bb)
                                {
                                    case 0:
                                        lkvBBS.Add(new KeyValuePair<string, object>("AL12", plObj.GetLineSegmentAt(bb).Length));
                                        break;
                                    case 1:
                                        lkvBBS.Add(new KeyValuePair<string, object>("BM13", plObj.GetLineSegmentAt(bb).Length));
                                        break;
                                    case 2:
                                        lkvBBS.Add(new KeyValuePair<string, object>("CN14", plObj.GetLineSegmentAt(bb).Length));
                                        break;
                                    default:
                                        //None
                                        break;
                                }
                            }
                            break;
                        case "CODE:310":
                            List<Point3d> lp3iVtx = new List<Point3d>();
                            for(int bb = 0; bb < plObj.NumberOfVertices; bb++)
                            {
                                lp3iVtx.Add(plObj.GetPoint3dAt(bb));
                            }
                            int nVtxG3w = plObj.NumberOfVertices;
                            actDoc.Editor.WriteMessage("\nN vtx: {0}, List length: {1}", nVtxG3w, lp3iVtx.Count);
                            Vector3d v3AP21 = lp3iVtx[1].GetVectorTo(lp3iVtx[0]);
                            Vector3d v3AP23 = lp3iVtx[1].GetVectorTo(lp3iVtx[2]);
                            double angleA = v3AP21.GetAngleTo(v3AP23);
                            //actDoc.Editor.WriteMessage("\nAngle at A: {0}", angleA);
                            if(angleA > 1.4708 && angleA < 1.6708)
                            {
                                for (int cc = 0; cc < plObj.NumberOfVertices - 1; cc++)
                                {
                                    switch (cc)
                                    {
                                        case 0:
                                            lkvBBS.Add(new KeyValuePair<string, object>("AL12", plObj.GetLineSegmentAt(cc).Length));
                                            break;                                                                     
                                        case 1:                                                                        
                                            lkvBBS.Add(new KeyValuePair<string, object>("BM13", plObj.GetLineSegmentAt(cc).Length));
                                            break;                                                                     
                                        case 2:                                                                        
                                            lkvBBS.Add(new KeyValuePair<string, object>("CN14", plObj.GetLineSegmentAt(cc).Length));
                                            break;                                                                     
                                        case 3:                                                                        
                                            lkvBBS.Add(new KeyValuePair<string, object>("DO15", plObj.GetLineSegmentAt(cc).Length));
                                            break;
                                        default:
                                            //None
                                            break;
                                    }
                                }
                                Vector3d v3KP21 = lp3iVtx[2].GetVectorTo(lp3iVtx[1]);
                                Vector3d v3KP23 = lp3iVtx[2].GetVectorTo(lp3iVtx[3]);
                                double angleK = v3KP21.GetAngleTo(v3KP23);
                                angleK = angleK * 180 / Math.PI;
                                lkvBBS.Add(new KeyValuePair<string, object>("EP16", ""));
                                lkvBBS.Add(new KeyValuePair<string, object>("FQ17", ""));
                                lkvBBS.Add(new KeyValuePair<string, object>("GR18", ""));
                                lkvBBS.Add(new KeyValuePair<string, object>("HS19", ""));
                                lkvBBS.Add(new KeyValuePair<string, object>("IT20", ""));
                                lkvBBS.Add(new KeyValuePair<string, object>("JU21", ""));
                                lkvBBS.Add(new KeyValuePair<string, object>("KV22", Math.Round(angleK, 2)));
                            }

                            Vector3d v3BP21 = lp3iVtx[nVtxG3w - 2].GetVectorTo(lp3iVtx[nVtxG3w - 1]);
                            Vector3d v3BP23 = lp3iVtx[nVtxG3w - 2].GetVectorTo(lp3iVtx[nVtxG3w - 3]);
                            double angleB = v3BP21.GetAngleTo(v3BP23);
                            //actDoc.Editor.WriteMessage("\nAngle at B: {0}", angleB);
                            if (angleB > 1.4708 && angleB < 1.6708)
                            {
                                for (int cc = 0; cc < plObj.NumberOfVertices - 1; cc++)
                                {
                                    switch (cc)
                                    {
                                        case 3:
                                            lkvBBS.Add(new KeyValuePair<string, object>("AL12", plObj.GetLineSegmentAt(cc).Length));
                                            break;
                                        case 2:
                                            lkvBBS.Add(new KeyValuePair<string, object>("BM13", plObj.GetLineSegmentAt(cc).Length));
                                            break;
                                        case 1:
                                            lkvBBS.Add(new KeyValuePair<string, object>("CN14", plObj.GetLineSegmentAt(cc).Length));
                                            break;
                                        case 0:
                                            lkvBBS.Add(new KeyValuePair<string, object>("DO15", plObj.GetLineSegmentAt(cc).Length));
                                            break;
                                        default:
                                            //None
                                            break;
                                    }
                                }
                                Vector3d v3KP21 = lp3iVtx[2].GetVectorTo(lp3iVtx[1]);
                                Vector3d v3KP23 = lp3iVtx[2].GetVectorTo(lp3iVtx[3]);
                                double angleK = v3KP21.GetAngleTo(v3KP23);
                                angleK = angleK * 180 / Math.PI;
                                lkvBBS.Add(new KeyValuePair<string, object>("EP16", ""));
                                lkvBBS.Add(new KeyValuePair<string, object>("FQ17", ""));
                                lkvBBS.Add(new KeyValuePair<string, object>("GR18", ""));
                                lkvBBS.Add(new KeyValuePair<string, object>("HS19", ""));
                                lkvBBS.Add(new KeyValuePair<string, object>("IT20", ""));
                                lkvBBS.Add(new KeyValuePair<string, object>("JU21", ""));
                                lkvBBS.Add(new KeyValuePair<string, object>("KV22",Math.Round(angleK,2)));
                            }
                            break;
                        default:
                            //None
                            break;
                    }
                }

                IList<object> ioGSValue = new List<object>();
                for (int cc = 0; cc < keyVal.Length; cc++)
                {
                    if (GetValueByKey(keyVal[cc]) != null)
                    {
                        ioGSValue.Add(GetValueByKey(keyVal[cc]));
                    }
                    else if(keyVal[cc] == "RemarkAC29")
                    {
                        ioGSValue.Add(beamName);
                    }
                    else
                    {
                        ioGSValue.Add("");
                    }
                }
                iioGSValue.Add(ioGSValue);

                bbsRowCounter++;
            }
            
            string strRangeAddrG3w = tabName + "!A" + ithRow.ToString();
            ValueRange vrRngValG3w = new ValueRange();
            vrRngValG3w.MajorDimension = "ROWS";
            vrRngValG3w.Values = iioGSValue;
            SpreadsheetsResource.ValuesResource.UpdateRequest urRngValG3w = DataGlobal.sheetsService.Spreadsheets
                .Values.Update(vrRngValG3w, gSheetID, strRangeAddrG3w);
            urRngValG3w.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            UpdateValuesResponse uvrRngValG3w = urRngValG3w.Execute();
            //Write to Google Sheet

            gsRowBBSG = ithRow + bbsRowCounter;
        }


        /**
         * Date added: 31 May 2024
         * Added by: Bernardo A. Cabebe Jr.
         * Venue: 31H Cityland Pasong Tamo Tower
         * Status: Under development
         */
        [CommandMethod("TCC_TabulateCuttingListCAD")]
        public static void TabulateCuttingListCAD()
        {
            if (!ccData.isLicenseActive())
            {
                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            PromptPointOptions ppoUpperLeftPt = new PromptPointOptions("");
            ppoUpperLeftPt.Message = "\nEnter upper left corner point: ";
            ppoUpperLeftPt.AllowNone = false;

            PromptPointResult pprUpperLeftPt = actDoc.Editor.GetPoint(ppoUpperLeftPt);
            if (pprUpperLeftPt.Status != PromptStatus.OK) return;
            Point3d p3UpperLeftPt = pprUpperLeftPt.Value;


            PromptPointOptions ppoLowerRightPt = new PromptPointOptions("");
            ppoLowerRightPt.Message = "\nEnter lower left corner point: ";
            ppoLowerRightPt.UseBasePoint = true;
            ppoLowerRightPt.BasePoint = p3UpperLeftPt;

            PromptPointResult pprLowerRightPt = actDoc.Editor.GetPoint(ppoLowerRightPt);
            if(pprLowerRightPt.Status != PromptStatus.OK) return;
            Point3d p3LowerRightPt = pprLowerRightPt.Value;


            PromptIntegerOptions pioStartingRow = new PromptIntegerOptions("");
            pioStartingRow.Message = "\nEnter staring row: ";
            pioStartingRow.AllowNone = true;
            pioStartingRow.DefaultValue = clStartingRowG;
            pioStartingRow.AllowZero = false;
            pioStartingRow.AllowNegative = false;
            pioStartingRow.LowerLimit = 3;

            PromptIntegerResult pirStartingRow = actDoc.Editor.GetInteger(pioStartingRow);
            if (pirStartingRow.Status != PromptStatus.OK) return;
            int startingRow = pirStartingRow.Value;
            clStartingRowG = startingRow;


            PromptIntegerOptions pioEndingRow = new PromptIntegerOptions("");
            pioEndingRow.Message = "\nEnter staring row: ";
            pioEndingRow.AllowNone = true;
            pioEndingRow.DefaultValue = clEndingRowG;
            pioEndingRow.AllowZero = false;
            pioEndingRow.AllowNegative = false;
            pioEndingRow.LowerLimit = 3;

            PromptIntegerResult pirEndingRow = actDoc.Editor.GetInteger(pioEndingRow);
            if (pirEndingRow.Status != PromptStatus.OK) return;
            int endingRow = pirEndingRow.Value;
            clEndingRowG = endingRow;


            PromptDoubleOptions pdoTextHgt = new PromptDoubleOptions("");
            pdoTextHgt.Message = "\nEnter text height: ";
            pdoTextHgt.UseDefaultValue = true;
            pdoTextHgt.DefaultValue = txtHgtG;
            pdoTextHgt.AllowNone = true;
            pdoTextHgt.AllowNegative = false;
            pdoTextHgt.AllowZero = false;

            PromptDoubleResult pdrTextHgt = actDoc.Editor.GetDouble(pdoTextHgt);
            if(pdrTextHgt.Status != PromptStatus.OK) return;
            double txtHgt = pdrTextHgt.Value;
            txtHgtG = txtHgt;

            PromptStringOptions psoTCCH0c = new PromptStringOptions("");
            psoTCCH0c.Message = "\nEnter Google Sheet ID:";
            psoTCCH0c.DefaultValue = clGSheetIDG;

            PromptResult prTCCH0c = actDoc.Editor.GetString(psoTCCH0c);
            if (prTCCH0c.Status != PromptStatus.OK) return;
            string clGSheetID = prTCCH0c.StringResult;
            clGSheetIDG = clGSheetID;


            double xDis = p3LowerRightPt.X - p3UpperLeftPt.X;
            double yDis = p3UpperLeftPt.Y - p3LowerRightPt.Y;
            Point3d p3Pt1 = pprUpperLeftPt.Value;
            Point3d p3Pt2 = p3Pt1 + new Vector3d(xDis, 0, 0);
            Point3d p3Pt3 = p3Pt1 + new Vector3d(0,-yDis, 0);
            Point3d p3Pt4 = p3LowerRightPt;
            Vector3d v3P12 = p3Pt1.GetVectorTo(p3Pt2);
            Vector3d nv3P12 = v3P12.GetNormal();
            Vector3d v3P13 = p3Pt1.GetVectorTo(p3Pt3);
            Vector3d nv3P13 = v3P13.GetNormal();

            GoogleSheetsV4 gsv4 = new GoogleSheetsV4();
            IList<IList<object>> iioGSHeaderV = new List<IList<object>>();
            string gsHdrRange = "CLTable!A1:J2";
            iioGSHeaderV = gsv4.GetRange(gsHdrRange, clGSheetID);
            for(int aa=0; aa<iioGSHeaderV.Count; aa++)
            {
                for(int bb=0; bb < iioGSHeaderV[0].Count; bb++)
                {
                    actDoc.Editor.WriteMessage("\nValue {0}{1}: {2}", aa, bb, iioGSHeaderV[aa][bb]);
                }
            }

            IList<IList<object>> iioGSValue = new List<IList<object>>();
            string gsRange = "CLTable!A" + startingRow.ToString() + ":J" + endingRow.ToString();
            iioGSValue = gsv4.GetRange(gsRange, clGSheetID);


            using(Transaction trTCCH0c = aCurDB.TransactionManager.StartTransaction())
            {
                BlockTable blktbl;
                blktbl = trTCCH0c.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord bltrec;
                bltrec = trTCCH0c.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                //Header vertical lines
                Point3d p3ccVLineP1 = p3Pt1;
                Point3d p3ccVLineP2 = p3Pt1 + 3 * txtHgt * nv3P13;
                using (Line acLine = new Line(p3ccVLineP1, p3ccVLineP2))
                {
                    acLine.ColorIndex = tableLineColor;
                    bltrec.AppendEntity(acLine);
                    trTCCH0c.AddNewlyCreatedDBObject(acLine, true);
                }
                for (int cc = 0; cc < 10; cc++)
                {
                    double scaleH0c = Convert.ToDouble(iioGSHeaderV[0][cc]) / 100;
                    p3ccVLineP1 = p3ccVLineP1 + scaleH0c * xDis * nv3P12;
                    p3ccVLineP2 = p3ccVLineP1 + 3 * txtHgt * nv3P13;
                    using (Line acLine = new Line(p3ccVLineP1, p3ccVLineP2))
                    {
                        acLine.ColorIndex = tableLineColor;
                        bltrec.AppendEntity(acLine);
                        trTCCH0c.AddNewlyCreatedDBObject(acLine, true);
                    }
                }

                //Header horizontal lines
                Point3d p3HdrTopHLineP1 = p3Pt1;
                Point3d p3HdrTopHLineP2 = p3Pt1 + xDis * nv3P12;
                using (Line acLine = new Line(p3HdrTopHLineP1, p3HdrTopHLineP2))
                {
                    acLine.ColorIndex = tableLineColor;
                    bltrec.AppendEntity(acLine);
                    trTCCH0c.AddNewlyCreatedDBObject(acLine, true);
                }

                //Header text
                Point3d p3ddTxtPt = p3Pt1 + 1.5 * txtHgt * nv3P13; ;
                for (int dd = 0; dd < 10; dd++)
                {

                    double scaleIH0c = (dd == 0) ? 0 : Convert.ToDouble(iioGSHeaderV[0][dd - 1]) / 100;
                    double scaleJH0c = Convert.ToDouble(iioGSHeaderV[0][dd]) / 100;
                    p3ddTxtPt = p3ddTxtPt + (0.5*scaleIH0c + 0.5*scaleJH0c) * xDis * nv3P12;
                    string txtVal = (string)iioGSHeaderV[1][dd];
                    using (DBText acText = new DBText())
                    {
                        acText.Position = p3ddTxtPt;
                        acText.Height = txtHgt;
                        acText.TextString = txtVal;
                        acText.HorizontalMode = TextHorizontalMode.TextCenter;
                        acText.VerticalMode = TextVerticalMode.TextVerticalMid;
                        acText.AlignmentPoint = p3ddTxtPt;
                        acText.ColorIndex = tableHeaderTextColor;
                        bltrec.AppendEntity(acText);
                        trTCCH0c.AddNewlyCreatedDBObject(acText, true);
                    }
                }


                //Horizontal lines of Content texts 
                Point3d p3ContXLnP1 = p3Pt1 + 3 * txtHgt * nv3P13;
                Point3d p3ContXLnP2 = p3ContXLnP1 + xDis * nv3P12;
                using (Line acLine = new Line(p3ContXLnP1, p3ContXLnP2))
                {
                    acLine.ColorIndex = tableLineColor;
                    bltrec.AppendEntity(acLine);
                    trTCCH0c.AddNewlyCreatedDBObject(acLine, true);
                }
                for (int ff=0; ff<iioGSValue.Count; ff++)
                {
                    p3ContXLnP1 = p3ContXLnP1 + 3 * txtHgt * nv3P13;
                    p3ContXLnP2 = p3ContXLnP2 + 3 * txtHgt * nv3P13;
                    using (Line acLine = new Line(p3ContXLnP1, p3ContXLnP2))
                    {
                        acLine.ColorIndex = tableLineColor;
                        bltrec.AppendEntity(acLine);
                        trTCCH0c.AddNewlyCreatedDBObject(acLine, true);
                    }
                }

                //Vertical lines of Content texts 
                Point3d p3ContYLnP1 = p3Pt1 + 3 * txtHgt * nv3P13;
                double contYDis = p3ContYLnP1.DistanceTo(p3ContXLnP1);
                Point3d p3ContYLnP2 = p3ContYLnP1 + contYDis * nv3P13;
                using (Line acLine = new Line(p3ContYLnP1, p3ContYLnP2))
                {
                    acLine.ColorIndex = tableLineColor;
                    bltrec.AppendEntity(acLine);
                    trTCCH0c.AddNewlyCreatedDBObject(acLine, true);
                }
                for (int ee = 0; ee < 10; ee++)
                {
                    double scaleH0c = Convert.ToDouble(iioGSHeaderV[0][ee]) / 100;
                    p3ContYLnP1 = p3ContYLnP1 + scaleH0c * xDis * nv3P12;
                    p3ContYLnP2 = p3ContYLnP1 + contYDis * nv3P13;
                    using (Line acLine = new Line(p3ContYLnP1, p3ContYLnP2))
                    {
                        acLine.ColorIndex = tableLineColor;
                        bltrec.AppendEntity(acLine);
                        trTCCH0c.AddNewlyCreatedDBObject(acLine, true);
                    }
                }

                //Content texts
                Point3d p3ContTxtPt = p3Pt1 + 1.5 * txtHgt * nv3P13;
                //p3ContTxtPt = p3ContTxtPt + 1.5 * txtHgt * nv3P13; //Center between two lines
                for (int gg=0; gg<iioGSValue.Count; gg++)
                {
                    p3ContTxtPt = p3ContTxtPt + (3 * txtHgt) * nv3P13;//Vertical position of text
                    Point3d p3ContTxtP2 = p3ContTxtPt;
                    for (int hh=0; hh < iioGSValue[gg].Count; hh++)
                    {
                        double scaleIH0c = (hh == 0) ? 0 : Convert.ToDouble(iioGSHeaderV[0][hh - 1]) / 100;
                        double scaleJH0c = Convert.ToDouble(iioGSHeaderV[0][hh]) / 100;
                        p3ContTxtP2 = p3ContTxtP2 + (0.5 * scaleIH0c + 0.5 * scaleJH0c) * xDis * nv3P12; //Horizontal position of text
                        string hhTxt = (string)iioGSValue[gg][hh];
                        using (DBText acText = new DBText())
                        {
                            acText.Position = p3ContTxtP2;
                            acText.Height = txtHgt;
                            acText.TextString = hhTxt;
                            acText.HorizontalMode = TextHorizontalMode.TextCenter;
                            acText.VerticalMode = TextVerticalMode.TextVerticalMid;
                            acText.AlignmentPoint = p3ContTxtP2;
                            acText.ColorIndex = tableTextColor;
                            bltrec.AppendEntity(acText);
                            trTCCH0c.AddNewlyCreatedDBObject(acText, true);
                        }
                    }
                }
                trTCCH0c.Commit();
            }
        }
    }
}
