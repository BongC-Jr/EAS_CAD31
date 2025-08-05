using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Interop.Common;
using Autodesk.AutoCAD.Runtime;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Autodesk.AutoCAD.ApplicationServices;

using acColor = Autodesk.AutoCAD.Colors.Color;

namespace EASI_CAD31
{
    public class ConcreteSlabDesign
    {
        static Document actDoc = Application.DocumentManager.MdiActiveDocument;
        static Database aCurDB = actDoc.Database;

        static CCData ccData = new CCData();

        static SEASTools seasTools = new SEASTools();

        /**
         * Date added: 17 Jan. 2024
         * Addede by: Engr Bernardo Cabebe Jr.
         * Venue: 3407 Cityland Pasong Tamo Tower
         */
        [CommandMethod("CSDA_AssignPolyline")]
        public static void ConcreteSlabDesignAssignPolyline()
        {
            if (!ccData.isLicenseActive())
            {

                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            PromptSelectionOptions psoSPline = new PromptSelectionOptions(); //Slab polyline
            psoSPline.MessageForAdding = "\nSelect slab polyline(s): ";

            PromptSelectionResult psrSPline = actDoc.Editor.GetSelection(psoSPline);
            if (psrSPline.Status != PromptStatus.OK) return;

            SelectionSet ssSPline = psrSPline.Value;

            using (Transaction trSPline = aCurDB.TransactionManager.StartTransaction())
            {
                //Create layer
                LayerTable lyrTbl = (LayerTable)trSPline.GetObject(aCurDB.LayerTableId, OpenMode.ForRead);
                string lyrName = "SEAS-ConcSlabDesign";
                if (lyrTbl.Has(lyrName) == false)
                {
                    using (LayerTableRecord lyrTblRec = new LayerTableRecord())
                    {
                        // Assign the layer the ACI color 57 and a name
                        lyrTblRec.Color = acColor.FromColorIndex(ColorMethod.ByAci, 57);
                        lyrTblRec.Name = lyrName;
                        lyrTblRec.IsPlottable = false;

                        // Upgrade the Layer table for write
                        trSPline.GetObject(aCurDB.LayerTableId, OpenMode.ForWrite);

                        // Append the new layer to the Layer table and the transaction
                        lyrTbl.Add(lyrTblRec);
                        trSPline.AddNewlyCreatedDBObject(lyrTblRec, true);
                    }
                }

                foreach (SelectedObject soSPline in ssSPline)
                {
                    Entity entSPline = (Entity)trSPline.GetObject(soSPline.ObjectId, OpenMode.ForWrite);
                    entSPline.Layer = lyrName;
                    entSPline.Color = acColor.FromColorIndex(ColorMethod.ByAci, 57);
                }
                trSPline.Commit();
                trSPline.Dispose();
            }
        }

        /**
         * Date added: 17 Jan. 2024
         * Addede by: Engr Bernardo Cabebe Jr.
         * Venue: 3407 Cityland Pasong Tamo Tower
         */
        [CommandMethod("CSDP_Perform")]
        public static void ConcreteSlabDesignPerform()
        {
            if (!ccData.isLicenseActive())
            {

                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            TypedValue[] tvFilterU6t = new TypedValue[2];
            tvFilterU6t[0] = new TypedValue((int)DxfCode.Start, "*POLYLINE*");
            tvFilterU6t[1] = new TypedValue((int)DxfCode.Color, 57);
            SelectionFilter sfSPline = new SelectionFilter(tvFilterU6t);


            PromptSelectionOptions psoRefObj = new PromptSelectionOptions();
            psoRefObj.MessageForAdding = "\nSelect slab polyline: ";

            PromptSelectionResult psrRefObj = actDoc.Editor.GetSelection(psoRefObj, sfSPline);
            if (psrRefObj.Status != PromptStatus.OK) return;

            SelectionSet ssRefObj = psrRefObj.Value;

            if (ssRefObj.Count != 1)
            {
                actDoc.Editor.WriteMessage("\nSelect one reference object only.");
                return;
            }

            //Edge 1
            Point3dCollection p3dcSelFence1 = new Point3dCollection();
            PromptPointOptions ppoSelFence11 = new PromptPointOptions("");
            ppoSelFence11.Message = "\nEdge 1 select point 1: ";
            ppoSelFence11.UseBasePoint = false;
            PromptPointResult pprSelFence11 = actDoc.Editor.GetPoint(ppoSelFence11);
            if (pprSelFence11.Status != PromptStatus.OK) return;
            p3dcSelFence1.Add(pprSelFence11.Value);

            PromptPointOptions ppoSelFence12 = new PromptPointOptions("");
            ppoSelFence12.Message = "\nEdge 1 select point 2: ";
            ppoSelFence12.UseBasePoint = true;
            ppoSelFence12.BasePoint = p3dcSelFence1[0];
            PromptPointResult pprSelFence12 = actDoc.Editor.GetPoint(ppoSelFence12);
            if (pprSelFence12.Status != PromptStatus.OK) return;
            p3dcSelFence1.Add(pprSelFence12.Value);

            PromptSelectionResult psrSelFence1 = actDoc.Editor.SelectFence(p3dcSelFence1, sfSPline);
            SelectionSet ssSelFence1 = psrSelFence1.Value;
            actDoc.Editor.WriteMessage("\nEdge 1 objects: {0}", ssSelFence1.Count);


            //Edge 2
            Point3dCollection p3dcSelFence2 = new Point3dCollection();
            PromptPointOptions ppoSelFence21 = new PromptPointOptions("");
            ppoSelFence21.Message = "\nEdge 2 select point 1: ";
            ppoSelFence21.UseBasePoint = false;
            PromptPointResult pprSelFence21 = actDoc.Editor.GetPoint(ppoSelFence21);
            if (pprSelFence21.Status != PromptStatus.OK) return;
            p3dcSelFence2.Add(pprSelFence21.Value);

            PromptPointOptions ppoSelFence22 = new PromptPointOptions("");
            ppoSelFence22.Message = "\nEdge 1 select point 2: ";
            ppoSelFence22.UseBasePoint = true;
            ppoSelFence22.BasePoint = p3dcSelFence2[0];
            PromptPointResult pprSelFence22 = actDoc.Editor.GetPoint(ppoSelFence22);
            if (pprSelFence22.Status != PromptStatus.OK) return;
            p3dcSelFence2.Add(pprSelFence22.Value);

            PromptSelectionResult psrSelFence2 = actDoc.Editor.SelectFence(p3dcSelFence2, sfSPline);
            SelectionSet ssSelFence2 = psrSelFence2.Value;
            actDoc.Editor.WriteMessage("\nEdge 2 objects: {0}", ssSelFence2.Count);


            //Edge 3
            Point3dCollection p3dcSelFence3 = new Point3dCollection();
            PromptPointOptions ppoSelFence31 = new PromptPointOptions("");
            ppoSelFence31.Message = "\nEdge 3 select point 1: ";
            ppoSelFence31.UseBasePoint = false;
            PromptPointResult pprSelFence31 = actDoc.Editor.GetPoint(ppoSelFence31);
            if (pprSelFence31.Status != PromptStatus.OK) return;
            p3dcSelFence3.Add(pprSelFence31.Value);

            PromptPointOptions ppoSelFence32 = new PromptPointOptions("");
            ppoSelFence32.Message = "\nEdge 3 select point 2: ";
            ppoSelFence32.UseBasePoint = true;
            ppoSelFence32.BasePoint = p3dcSelFence3[0];
            PromptPointResult pprSelFence32 = actDoc.Editor.GetPoint(ppoSelFence32);
            if (pprSelFence32.Status != PromptStatus.OK) return;
            p3dcSelFence3.Add(pprSelFence32.Value);

            PromptSelectionResult psrSelFence3 = actDoc.Editor.SelectFence(p3dcSelFence3, sfSPline);
            SelectionSet ssSelFence3 = psrSelFence3.Value;
            actDoc.Editor.WriteMessage("\nEdge 3 objects: {0}", ssSelFence3.Count);


            //Edge 4
            Point3dCollection p3dcSelFence4 = new Point3dCollection();
            PromptPointOptions ppoSelFence41 = new PromptPointOptions("");
            ppoSelFence41.Message = "\nEdge 4 select point 1: ";
            ppoSelFence41.UseBasePoint = false;
            PromptPointResult pprSelFence41 = actDoc.Editor.GetPoint(ppoSelFence41);
            if (pprSelFence41.Status != PromptStatus.OK) return;
            p3dcSelFence4.Add(pprSelFence41.Value);

            PromptPointOptions ppoSelFence42 = new PromptPointOptions("");
            ppoSelFence42.Message = "\nEdge 4 select point 2: ";
            ppoSelFence42.UseBasePoint = true;
            ppoSelFence42.BasePoint = p3dcSelFence4[0];
            PromptPointResult pprSelFence42 = actDoc.Editor.GetPoint(ppoSelFence42);
            if (pprSelFence42.Status != PromptStatus.OK) return;
            p3dcSelFence4.Add(pprSelFence42.Value);

            PromptSelectionResult psrSelFence4 = actDoc.Editor.SelectFence(p3dcSelFence4, sfSPline);
            SelectionSet ssSelFence4 = psrSelFence4.Value;
            actDoc.Editor.WriteMessage("\nEdge 4 objects: {0}", ssSelFence4.Count);


            //Get slab design data
            PromptStringOptions psoSDD = new PromptStringOptions("");
            psoSDD.Message = "\nEnter design data: ";
            psoSDD.DefaultValue = ccData.getSlabDesignData();
            psoSDD.AllowSpaces = false;
            PromptResult prSDD = actDoc.Editor.GetString(psoSDD);
            if (prSDD.Status != PromptStatus.OK) return;
            string strSDD = prSDD.StringResult;
            ccData.setSlabDesignData(strSDD);

            string[] arrSDD = strSDD.Split(',');
            if (arrSDD.Length != 14)
            {
                actDoc.Editor.WriteMessage("\nCheck slab design data. Incorrect.");
                return;
            }
            actDoc.Editor.WriteMessage("\nT={0}mm, beam D={1}mm, Dm={2}mm, SDL={3}kPa, LL={4}kPa, OL={5}kPa, f'c={6}MPa, fy={7}kPa, SWt k1={8}, SDL k2={9}, LL k3={10}, OL k4={11}, text ht.={12},OWSS={12}",
                arrSDD[0], arrSDD[1], arrSDD[2], arrSDD[3], arrSDD[4], arrSDD[5], arrSDD[6], arrSDD[7], arrSDD[8], arrSDD[9], arrSDD[10], arrSDD[11], arrSDD[12], arrSDD[13]);

            GoogleSheetsV4 gsv4 = new GoogleSheetsV4();
            using (Transaction trSPLine = aCurDB.TransactionManager.StartTransaction()) //trSPLine = Slab Polyline
            {
                DBObjectCollection dbocRefObj = new DBObjectCollection();
                //foreach(SelectedObject soRefObj in ssRefObj)
                //{
                //    dbocRefObj.Add((DBObject)trSPLine.GetObject(soRefObj.ObjectId, OpenMode.ForRead));
                //}
                dbocRefObj.Add((DBObject)trSPLine.GetObject(ssRefObj[0].ObjectId, OpenMode.ForRead));

                //Find centroid of polyline
                Solid3d solidU3e = new Solid3d();
                solidU3e.Extrude((Region)Region.CreateFromCurves(dbocRefObj)[0], 1, 0);
                Point2d centroidU3e = new Point2d(solidU3e.MassProperties.Centroid.X, solidU3e.MassProperties.Centroid.Y);
                solidU3e.Dispose();
                //actDoc.Editor.WriteMessage("\nCentroid: {0}", centroidU3e.ToString());

                BlockTable blkTblU6t = (BlockTable)trSPLine.GetObject(aCurDB.BlockTableId, OpenMode.ForRead);
                BlockTableRecord blkTblRecU6t = (BlockTableRecord)trSPLine.GetObject(blkTblU6t[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                string lyrNameU6t = "SEAS-ConcSlabDesign";

                using (DBPoint dbPointU6t = new DBPoint(new Point3d(centroidU3e.X, centroidU3e.Y, 0)))
                {
                    dbPointU6t.Layer = lyrNameU6t;
                    dbPointU6t.Color = acColor.FromColorIndex(ColorMethod.ByAci, 57);
                    blkTblRecU6t.AppendEntity(dbPointU6t);
                    trSPLine.AddNewlyCreatedDBObject(dbPointU6t, true);
                }


                //Get slab dimensions using point 1 and 3
                double sDim1;
                Curve curSPline = (Curve)trSPLine.GetObject(ssRefObj[0].ObjectId, OpenMode.ForRead);
                Point3d p3dP1U6t = curSPline.GetClosestPointTo(p3dcSelFence1[0], true);
                Point3d p3dP3U6t = curSPline.GetClosestPointTo(p3dcSelFence3[0], true);
                sDim1 = p3dP1U6t.DistanceTo(p3dP3U6t);

                //Get slab dimensions using point 2 and 4
                double sDim2;
                Point3d p3dP2U6t = curSPline.GetClosestPointTo(p3dcSelFence2[0], true);
                Point3d p3dP4U6t = curSPline.GetClosestPointTo(p3dcSelFence4[0], true);
                sDim2 = p3dP2U6t.DistanceTo(p3dP4U6t);

                double sShortU6t, sLongU6t; //Slab short dimension and slab long dimension
                if (sDim1 <= sDim2)
                {
                    sShortU6t = sDim1;
                    sLongU6t = sDim2;
                }
                else
                {
                    sShortU6t = sDim2;
                    sLongU6t = sDim1;
                }
                actDoc.Editor.WriteMessage("\nShort dimension: {0}mm; Long dimension: {1}mm", Math.Round(sShortU6t,2), Math.Round(sLongU6t,2));

                string[,] arr2SDim = new string[1, 2];
                arr2SDim[0, 0] = sShortU6t.ToString();
                arr2SDim[0, 1] = sLongU6t.ToString();
                List<IList<object>> lioSDim = new List<IList<object>>();
                for (int ii = 0; ii < arr2SDim.GetLength(0); ii++)
                {
                    List<object> loSDim = new List<object>();
                    for (int jj = 0; jj < arr2SDim.GetLength(1); jj++)
                    {
                        loSDim.Add(arr2SDim[ii, jj]);
                    }
                    lioSDim.Add(loSDim);
                }

                IList<IList<object>> iioRngVal;

                if ((sShortU6t / sLongU6t) <= 0.5)
                {
                    //Slab dimensions
                    string strCellAddr = "OWS!F13:G13";
                    ValueRange vrUpdateCellU6t = new ValueRange();
                    vrUpdateCellU6t.MajorDimension = "ROWS";
                    vrUpdateCellU6t.Values = lioSDim;
                    SpreadsheetsResource.ValuesResource.UpdateRequest urUpdateCellU6t = DataGlobal.sheetsService.Spreadsheets
                        .Values.Update(vrUpdateCellU6t, ccData.getSlabGSheetID(), strCellAddr);
                    urUpdateCellU6t.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                    UpdateValuesResponse exeUpdateCellU6t = urUpdateCellU6t.Execute();

                    //Slab design data: dimensions
                    Thread.Sleep(250); //Put a delay due to internet connectioin
                    string[,] arr2RngValU8t = new string[1, 2];
                    arr2RngValU8t[0, 0] = arrSDD[0];//slab thickness
                    arr2RngValU8t[0, 1] = arrSDD[1];//beam depth
                    List<IList<object>> lioRVU8t = new List<IList<object>>();
                    for (int ii = 0; ii < arr2RngValU8t.GetLength(0); ii++)
                    {
                        List<object> loVRU8t = new List<object>();
                        for (int jj = 0; jj < arr2RngValU8t.GetLength(1); jj++)
                        {
                            loVRU8t.Add(arr2RngValU8t[ii, jj]);
                        }
                        lioRVU8t.Add(loVRU8t);
                    }

                    string strRangeAddrU8t = "OWS!B13:C13";
                    ValueRange vrRngValU8t = new ValueRange();
                    vrRngValU8t.MajorDimension = "ROWS";
                    vrRngValU8t.Values = lioRVU8t;
                    SpreadsheetsResource.ValuesResource.UpdateRequest urRngValU8t = DataGlobal.sheetsService.Spreadsheets
                        .Values.Update(vrRngValU8t, ccData.getSlabGSheetID(), strRangeAddrU8t);
                    urRngValU8t.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                    UpdateValuesResponse uvrRngValU8t = urRngValU8t.Execute();


                    //Slab design data: Material strengths
                    Thread.Sleep(250); //Put a delay due to internet connectioin
                    string[,] arr2MatS = new string[1, 3];
                    arr2MatS[0, 0] = arrSDD[6];  //concrete strength
                    arr2MatS[0, 1] = arrSDD[7];  //steel fy
                    arr2MatS[0, 2] = arrSDD[2]; //rebar diameter
                    List<IList<object>> lioMatS = new List<IList<object>>();
                    for (int ii = 0; ii < arr2MatS.GetLength(0); ii++)
                    {
                        List<object> loMatS = new List<object>();
                        for (int jj = 0; jj < arr2MatS.GetLength(1); jj++)
                        {
                            loMatS.Add(arr2MatS[ii, jj]);
                        }
                        lioMatS.Add(loMatS);
                    }
                    string strMatS = "OWS!F18:H18";
                    ValueRange vrMatS = new ValueRange();
                    vrMatS.MajorDimension = "ROWS";
                    vrMatS.Values = lioMatS;
                    SpreadsheetsResource.ValuesResource.UpdateRequest urMatS = DataGlobal.sheetsService.Spreadsheets
                        .Values.Update(vrMatS, ccData.getSlabGSheetID(), strMatS);
                    urMatS.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                    UpdateValuesResponse uvrMatS = urMatS.Execute();


                    //Slab design data: Slab loads
                    Thread.Sleep(250); //Put a delay due to internet connectioin
                    string[,] arr2SLoad = new string[1, 3];
                    arr2SLoad[0, 0] = arrSDD[3];  //superimposed load
                    arr2SLoad[0, 1] = arrSDD[4];  //live load
                    arr2SLoad[0, 2] = arrSDD[5]; //other load
                    List<IList<object>> lioSLoad = new List<IList<object>>();
                    for (int ii = 0; ii < arr2SLoad.GetLength(0); ii++)
                    {
                        List<object> loSLoad = new List<object>();
                        for (int jj = 0; jj < arr2SLoad.GetLength(1); jj++)
                        {
                            loSLoad.Add(arr2SLoad[ii, jj]);
                        }
                        lioSLoad.Add(loSLoad);
                    }
                    string strSLoad = "OWS!C18:E18";
                    ValueRange vrSLoad = new ValueRange();
                    vrSLoad.MajorDimension = "ROWS";
                    vrSLoad.Values = lioSLoad;
                    SpreadsheetsResource.ValuesResource.UpdateRequest urSLoad = DataGlobal.sheetsService.Spreadsheets
                        .Values.Update(vrSLoad, ccData.getSlabGSheetID(), strSLoad);
                    urSLoad.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                    UpdateValuesResponse uvrSLoad = urSLoad.Execute();


                    //Slab design data: load factors
                    Thread.Sleep(250); //Put a delay due to internet connectioin
                    string[,] arr2LFac = new string[1, 4];
                    arr2LFac[0, 0] = arrSDD[8];  //self weight factor
                    arr2LFac[0, 1] = arrSDD[9];  //superimposed load factor
                    arr2LFac[0, 2] = arrSDD[10];  //live load factor
                    arr2LFac[0, 3] = arrSDD[11];  //other loads factor
                    List<IList<object>> lioLFac = new List<IList<object>>();
                    for (int ii = 0; ii < arr2LFac.GetLength(0); ii++)
                    {
                        List<object> loLFac = new List<object>();
                        for (int jj = 0; jj < arr2LFac.GetLength(1); jj++)
                        {
                            loLFac.Add(arr2LFac[ii, jj]);
                        }
                        lioLFac.Add(loLFac);
                    }
                    string strLFac = "OWS!C22:F22";
                    ValueRange vrLFac = new ValueRange();
                    vrLFac.MajorDimension = "ROWS";
                    vrLFac.Values = lioLFac;
                    SpreadsheetsResource.ValuesResource.UpdateRequest urLFac = DataGlobal.sheetsService.Spreadsheets
                        .Values.Update(vrLFac, ccData.getSlabGSheetID(), strLFac);
                    urLFac.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                    UpdateValuesResponse uvrLFac = urLFac.Execute();


                    //Edge constraints
                    Thread.Sleep(250); //Put a delay due to internet connectioin
                    string[,] arr2EdgC = new string[4, 1];
                    arr2EdgC[0,0] = ssSelFence1.Count.ToString();  //Number of slabs crossed
                    arr2EdgC[1,0] = ssSelFence3.Count.ToString();  //Number of slabs crossed
                    arr2EdgC[2,0] = ssSelFence2.Count.ToString();  //Number of slabs crossed
                    arr2EdgC[3,0] = ssSelFence4.Count.ToString();  //Number of slabs crossed
                    List<IList<object>> lioEdgC = new List<IList<object>>();
                    for (int ii = 0; ii < arr2EdgC.GetLength(0); ii++)
                    {
                        List<object> loEdgC = new List<object>();
                        for (int jj = 0; jj < arr2EdgC.GetLength(1); jj++)
                        {
                            loEdgC.Add(arr2EdgC[ii, jj]);
                        }
                        lioEdgC.Add(loEdgC);
                    }
                    string strEdgC = "OWS!C169:C172";
                    ValueRange vrEdgC = new ValueRange();
                    vrEdgC.MajorDimension = "ROWS";
                    vrEdgC.Values = lioEdgC;
                    SpreadsheetsResource.ValuesResource.UpdateRequest urEdgC = DataGlobal.sheetsService.Spreadsheets
                        .Values.Update(vrEdgC, ccData.getSlabGSheetID(), strEdgC);
                    urEdgC.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                    UpdateValuesResponse uvrEdgC = urEdgC.Execute();


                    //OWS span number
                    Thread.Sleep(250); //Put a delay due to internet connectioin
                    string[,] arr2SSpan = new string[1, 1];
                    arr2SSpan[0, 0] = arrSDD[13];  //Number of slab spans
                    List<IList<object>> lioSSpan = new List<IList<object>>();
                    for (int ii = 0; ii < arr2SSpan.GetLength(0); ii++)
                    {
                        List<object> loSSpan = new List<object>();
                        for (int jj = 0; jj < arr2SSpan.GetLength(1); jj++)
                        {
                            loSSpan.Add(arr2SSpan[ii, jj]);
                        }
                        lioSSpan.Add(loSSpan);
                    }
                    string strSSpan = "OWS!C173";
                    ValueRange vrSSpan = new ValueRange();
                    vrSSpan.MajorDimension = "ROWS";
                    vrSSpan.Values = lioSSpan;
                    SpreadsheetsResource.ValuesResource.UpdateRequest urSSpan = DataGlobal.sheetsService.Spreadsheets
                        .Values.Update(vrSSpan, ccData.getSlabGSheetID(), strSSpan);
                    urSSpan.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                    UpdateValuesResponse uvrSSpan = urSSpan.Execute();


                    //Get slab mark
                    //iioRngVal = GoogleSheetsV4.GetRange("OWS!C6", ccData.getSlabGSheetID());
                    iioRngVal = gsv4.GetRange("OWS!C6", ccData.getSlabGSheetID());

                }
                else
                {
                    //Slab dimensions
                    string strCellAddrU7t = "TWS!E13:F13";
                    ValueRange vrUpdateCellU7t = new ValueRange();
                    vrUpdateCellU7t.MajorDimension = "ROWS";
                    vrUpdateCellU7t.Values = lioSDim;
                    SpreadsheetsResource.ValuesResource.UpdateRequest urUpdateCellU7t = DataGlobal.sheetsService.Spreadsheets
                        .Values.Update(vrUpdateCellU7t, ccData.getSlabGSheetID(), strCellAddrU7t);
                    urUpdateCellU7t.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                    UpdateValuesResponse exeUpdateCellU7t = urUpdateCellU7t.Execute();

                    //Slab design data: dimensions
                    Thread.Sleep(250); //Put a delay due to internet connectioin
                    string[,] arr2RngValU8t = new string[1, 2];
                    arr2RngValU8t[0, 0] = arrSDD[0];//slab thickness
                    arr2RngValU8t[0, 1] = arrSDD[1];//beam depth
                    List<IList<object>> lioRVU8t = new List<IList<object>>();
                    for (int ii = 0; ii < arr2RngValU8t.GetLength(0); ii++)
                    {
                        List<object> loVRU8t = new List<object>();
                        for (int jj = 0; jj < arr2RngValU8t.GetLength(1); jj++)
                        {
                            loVRU8t.Add(arr2RngValU8t[ii, jj]);
                        }
                        lioRVU8t.Add(loVRU8t);
                    }

                    string strRangeAddrU8t = "TWS!A13:B13";
                    ValueRange vrRngValU8t = new ValueRange();
                    vrRngValU8t.MajorDimension = "ROWS";
                    vrRngValU8t.Values = lioRVU8t;
                    SpreadsheetsResource.ValuesResource.UpdateRequest urRngValU8t = DataGlobal.sheetsService.Spreadsheets
                        .Values.Update(vrRngValU8t, ccData.getSlabGSheetID(), strRangeAddrU8t);
                    urRngValU8t.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                    UpdateValuesResponse uvrRngValU8t = urRngValU8t.Execute();


                    //Slab design data: Material strengths
                    Thread.Sleep(250); //Put a delay due to internet connectioin
                    string[,] arr2MatS = new string[1, 3];
                    arr2MatS[0, 0] = arrSDD[6];  //concrete strength
                    arr2MatS[0, 1] = arrSDD[7];  //steel fy
                    arr2MatS[0, 2] = arrSDD[2]; //rebar diameter
                    List<IList<object>> lioMatS = new List<IList<object>>();
                    for (int ii = 0; ii < arr2MatS.GetLength(0); ii++)
                    {
                        List<object> loMatS = new List<object>();
                        for (int jj = 0; jj < arr2MatS.GetLength(1); jj++)
                        {
                            loMatS.Add(arr2MatS[ii, jj]);
                        }
                        lioMatS.Add(loMatS);
                    }
                    string strMatS = "TWS!A18:C18";
                    ValueRange vrMatS = new ValueRange();
                    vrMatS.MajorDimension = "ROWS";
                    vrMatS.Values = lioMatS;
                    SpreadsheetsResource.ValuesResource.UpdateRequest urMatS = DataGlobal.sheetsService.Spreadsheets
                        .Values.Update(vrMatS, ccData.getSlabGSheetID(), strMatS);
                    urMatS.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                    UpdateValuesResponse uvrMatS = urMatS.Execute();


                    //Slab design data: Slab loads
                    Thread.Sleep(250); //Put a delay due to internet connectioin
                    string[,] arr2SLoad = new string[1, 3];
                    arr2SLoad[0, 0] = arrSDD[3];  //superimposed load
                    arr2SLoad[0, 1] = arrSDD[4];  //live load
                    arr2SLoad[0, 2] = arrSDD[5]; //other load
                    List<IList<object>> lioSLoad = new List<IList<object>>();
                    for (int ii = 0; ii < arr2SLoad.GetLength(0); ii++)
                    {
                        List<object> loSLoad = new List<object>();
                        for (int jj = 0; jj < arr2SLoad.GetLength(1); jj++)
                        {
                            loSLoad.Add(arr2SLoad[ii, jj]);
                        }
                        lioSLoad.Add(loSLoad);
                    }
                    string strSLoad = "TWS!B23:D23";
                    ValueRange vrSLoad = new ValueRange();
                    vrSLoad.MajorDimension = "ROWS";
                    vrSLoad.Values = lioSLoad;
                    SpreadsheetsResource.ValuesResource.UpdateRequest urSLoad = DataGlobal.sheetsService.Spreadsheets
                        .Values.Update(vrSLoad, ccData.getSlabGSheetID(), strSLoad);
                    urSLoad.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                    UpdateValuesResponse uvrSLoad = urSLoad.Execute();


                    //Slab design data: load factors
                    Thread.Sleep(250); //Put a delay due to internet connectioin
                    string[,] arr2LFac = new string[1, 4];
                    arr2LFac[0, 0] = arrSDD[8];  //self weight factor
                    arr2LFac[0, 1] = arrSDD[9];  //superimposed load factor
                    arr2LFac[0, 2] = arrSDD[10];  //live load factor
                    arr2LFac[0, 3] = arrSDD[11];  //other loads factor
                    List<IList<object>> lioLFac = new List<IList<object>>();
                    for (int ii = 0; ii < arr2LFac.GetLength(0); ii++)
                    {
                        List<object> loLFac = new List<object>();
                        for (int jj = 0; jj < arr2LFac.GetLength(1); jj++)
                        {
                            loLFac.Add(arr2LFac[ii, jj]);
                        }
                        lioLFac.Add(loLFac);
                    }
                    string strLFac = "TWS!B27:E27";
                    ValueRange vrLFac = new ValueRange();
                    vrLFac.MajorDimension = "ROWS";
                    vrLFac.Values = lioLFac;
                    SpreadsheetsResource.ValuesResource.UpdateRequest urLFac = DataGlobal.sheetsService.Spreadsheets
                        .Values.Update(vrLFac, ccData.getSlabGSheetID(), strLFac);
                    urLFac.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                    UpdateValuesResponse uvrLFac = urLFac.Execute();


                    //Edge constraints
                    Thread.Sleep(250); //Put a delay due to internet connectioin
                    string[,] arr2EdgC = new string[1, 4];
                    arr2EdgC[0, 0] = ssSelFence1.Count.ToString();  //Number of slabs crossed
                    arr2EdgC[0, 1] = ssSelFence3.Count.ToString();  //Number of slabs crossed
                    arr2EdgC[0, 2] = ssSelFence2.Count.ToString();  //Number of slabs crossed
                    arr2EdgC[0, 3] = ssSelFence4.Count.ToString();  //Number of slabs crossed
                    List<IList<object>> lioEdgC = new List<IList<object>>();
                    for (int ii = 0; ii < arr2EdgC.GetLength(0); ii++)
                    {
                        List<object> loEdgC = new List<object>();
                        for (int jj = 0; jj < arr2EdgC.GetLength(1); jj++)
                        {
                            loEdgC.Add(arr2EdgC[ii, jj]);
                        }
                        lioEdgC.Add(loEdgC);
                    }
                    string strEdgC = "TWS!G37:G40";
                    ValueRange vrEdgC = new ValueRange();
                    vrEdgC.MajorDimension = "COLUMNS";
                    vrEdgC.Values = lioEdgC;
                    SpreadsheetsResource.ValuesResource.UpdateRequest urEdgC = DataGlobal.sheetsService.Spreadsheets
                        .Values.Update(vrEdgC, ccData.getSlabGSheetID(), strEdgC);
                    urEdgC.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                    UpdateValuesResponse uvrEdgC = urEdgC.Execute();

                    //Get slab mark
                    //iioRngVal = GoogleSheetsV4.GetRange("TWS!B6", ccData.getSlabGSheetID());
                    iioRngVal = gsv4.GetRange("TWS!B6", ccData.getSlabGSheetID());
                }

                Point2d p2dCntrd = centroidU3e; 
                Point2d p2dE4Pt1 = new Point2d(p3dcSelFence4[0].X, p3dcSelFence4[0].Y);
                double angleG2e = p2dCntrd.GetVectorTo(p2dE4Pt1).Angle;
                /*actDoc.Editor.WriteMessage("\nAngle 1: {0}", angleG2e);*/
                if (angleG2e >= 1.5708 && angleG2e <= 4.7124)
                {
                    angleG2e += Math.PI;
                }
                /*actDoc.Editor.WriteMessage("\nAngle 2: {0}",angleG2e);*/

                Matrix3d curUCSMatrix = actDoc.Editor.CurrentUserCoordinateSystem;
                CoordinateSystem3d curUCS = curUCSMatrix.CoordinateSystem3d;

               //Create layer
                string slabMarkLayer = "SEAS-SlabSchedule";
                seasTools.CreateLayer(slabMarkLayer, 3, true);
                Thread.Sleep(250); //Put a delay to wait spreadsheets to perform some process
                using (DBText acTxt = new DBText())
                {
                    acTxt.Position = new Point3d(centroidU3e.X, centroidU3e.Y, 0.0);
                    acTxt.Height = Convert.ToDouble(arrSDD[12]);
                    acTxt.TextString = iioRngVal[0][0].ToString();
                    acTxt.Layer = slabMarkLayer;
                    acTxt.HorizontalMode = TextHorizontalMode.TextCenter;
                    acTxt.VerticalMode = TextVerticalMode.TextVerticalMid;
                    acTxt.AlignmentPoint = new Point3d(centroidU3e.X, centroidU3e.Y, 0.0);
                    acTxt.Color = acColor.FromColorIndex(ColorMethod.ByAci, 2);
                    acTxt.TransformBy(Matrix3d.Rotation(angleG2e, curUCS.Zaxis, new Point3d(centroidU3e.X, centroidU3e.Y, 0.0)));

                    blkTblRecU6t.AppendEntity(acTxt);
                    trSPLine.AddNewlyCreatedDBObject(acTxt, true);
                }

                double txHgt = Convert.ToDouble(arrSDD[12]);
                using (Polyline plTxt = new Polyline())
                {
                    plTxt.AddVertexAt(0, new Point2d(centroidU3e.X - 2.75 * txHgt, centroidU3e.Y + txHgt), 0, 0, 0);
                    plTxt.AddVertexAt(0, new Point2d(centroidU3e.X + 2.75 * txHgt, centroidU3e.Y + txHgt), 0, 0, 0);
                    plTxt.AddVertexAt(0, new Point2d(centroidU3e.X + 2.75 * txHgt, centroidU3e.Y - txHgt), 0, 0, 0);
                    plTxt.AddVertexAt(0, new Point2d(centroidU3e.X - 2.75 * txHgt, centroidU3e.Y - txHgt), 0, 0, 0);
                    plTxt.AddVertexAt(0, new Point2d(centroidU3e.X - 2.75 * txHgt, centroidU3e.Y + txHgt), 0, 0, 0);

                    plTxt.Layer = slabMarkLayer;
                    plTxt.Color = acColor.FromColorIndex(ColorMethod.ByAci, 251);

                    plTxt.TransformBy(Matrix3d.Rotation(angleG2e, curUCS.Zaxis, new Point3d(centroidU3e.X, centroidU3e.Y, 0.0)));

                    blkTblRecU6t.AppendEntity(plTxt);
                    trSPLine.AddNewlyCreatedDBObject(plTxt, true);
                }


                Point2d p2dCntr = centroidU3e;
                //Edge 1 Txt
                Point2d p2dPnt1 = new Point2d(p3dcSelFence1[0].X, p3dcSelFence1[0].Y);
                Point3d p3dETx1;
                using (Line lTx1 = new Line())
                {
                    lTx1.StartPoint = new Point3d(p2dCntr.X, p2dCntr.Y, 0.0);
                    lTx1.EndPoint = new Point3d(p2dPnt1.X, p2dPnt1.Y, 0.0);

                    Curve cTx1 = (Curve)lTx1;

                    p3dETx1 = cTx1.GetPointAtDist(1.5 * txHgt);
                    cTx1.Dispose();
                }
                using (DBText acTx1 = new DBText())
                {
                    acTx1.Position = p3dETx1;
                    acTx1.Height = 0.5 * txHgt;
                    acTx1.TextString = "1";
                    acTx1.Layer = slabMarkLayer;
                    acTx1.HorizontalMode = TextHorizontalMode.TextCenter;
                    acTx1.VerticalMode = TextVerticalMode.TextVerticalMid;
                    acTx1.AlignmentPoint = p3dETx1;
                    acTx1.Color = acColor.FromColorIndex(ColorMethod.ByAci, 1);
                    acTx1.TransformBy(Matrix3d.Rotation(angleG2e, curUCS.Zaxis, p3dETx1));

                    blkTblRecU6t.AppendEntity(acTx1);
                    trSPLine.AddNewlyCreatedDBObject(acTx1, true);
                }

                //Edge 2 Txt
                Point2d p2dPnt2 = new Point2d(p3dcSelFence2[0].X, p3dcSelFence2[0].Y);
                Point3d p3dETx2;
                using (Line lTx2 = new Line())
                {
                    lTx2.StartPoint = new Point3d(p2dCntr.X, p2dCntr.Y, 0.0);
                    lTx2.EndPoint = new Point3d(p2dPnt2.X, p2dPnt2.Y, 0.0);

                    Curve cTx2 = (Curve)lTx2;

                    p3dETx2 = cTx2.GetPointAtDist(3.15 * txHgt);
                    cTx2.Dispose();
                }
                using (DBText acTx2 = new DBText())
                {
                    acTx2.Position = p3dETx2;
                    acTx2.Height = 0.5 * txHgt;
                    acTx2.TextString = "2";
                    acTx2.Layer = slabMarkLayer;
                    acTx2.HorizontalMode = TextHorizontalMode.TextCenter;
                    acTx2.VerticalMode = TextVerticalMode.TextVerticalMid;
                    acTx2.AlignmentPoint = p3dETx2;
                    acTx2.Color = acColor.FromColorIndex(ColorMethod.ByAci, 1);
                    acTx2.TransformBy(Matrix3d.Rotation(angleG2e, curUCS.Zaxis, p3dETx2));

                    blkTblRecU6t.AppendEntity(acTx2);
                    trSPLine.AddNewlyCreatedDBObject(acTx2, true);
                }

                //Edge 3 Txt
                Point2d p2dPnt3 = new Point2d(p3dcSelFence3[0].X, p3dcSelFence3[0].Y);
                Point3d p3dETx3;
                using (Line lTx3 = new Line())
                {
                    lTx3.StartPoint = new Point3d(p2dCntr.X, p2dCntr.Y, 0.0);
                    lTx3.EndPoint = new Point3d(p2dPnt3.X, p2dPnt3.Y, 0.0);

                    Curve cTx3 = (Curve)lTx3;

                    p3dETx3 = cTx3.GetPointAtDist(1.5 * txHgt);
                    cTx3.Dispose();
                }
                using (DBText acTx3 = new DBText())
                {
                    acTx3.Position = p3dETx3;
                    acTx3.Height = 0.5 * txHgt;
                    acTx3.TextString = "3";
                    acTx3.Layer = slabMarkLayer;
                    acTx3.HorizontalMode = TextHorizontalMode.TextCenter;
                    acTx3.VerticalMode = TextVerticalMode.TextVerticalMid;
                    acTx3.AlignmentPoint = p3dETx3;
                    acTx3.Color = acColor.FromColorIndex(ColorMethod.ByAci, 1);
                    acTx3.TransformBy(Matrix3d.Rotation(angleG2e, curUCS.Zaxis, p3dETx3));

                    blkTblRecU6t.AppendEntity(acTx3);
                    trSPLine.AddNewlyCreatedDBObject(acTx3, true);
                }

                //Edge 4 Txt
                Point2d p2dPnt4 = new Point2d(p3dcSelFence4[0].X, p3dcSelFence4[0].Y);
                Point3d p3dETx4;
                using (Line lTx4 = new Line())
                {
                    lTx4.StartPoint = new Point3d(p2dCntr.X, p2dCntr.Y, 0.0);
                    lTx4.EndPoint = new Point3d(p2dPnt4.X, p2dPnt4.Y, 0.0);

                    Curve cTx4 = (Curve)lTx4;

                    p3dETx4 = cTx4.GetPointAtDist(3.15 * txHgt);
                    cTx4.Dispose();
                }
                using (DBText acTx4 = new DBText())
                {
                    acTx4.Position = p3dETx4;
                    acTx4.Height = 0.5 * txHgt;
                    acTx4.TextString = "4";
                    acTx4.Layer = slabMarkLayer;
                    acTx4.HorizontalMode = TextHorizontalMode.TextCenter;
                    acTx4.VerticalMode = TextVerticalMode.TextVerticalMid;
                    acTx4.AlignmentPoint = p3dETx4;
                    acTx4.Color = acColor.FromColorIndex(ColorMethod.ByAci, 1);
                    acTx4.TransformBy(Matrix3d.Rotation(angleG2e, curUCS.Zaxis, p3dETx4));

                    blkTblRecU6t.AppendEntity(acTx4);
                    trSPLine.AddNewlyCreatedDBObject(acTx4, true);


                }

                trSPLine.Commit();
                trSPLine.Dispose();
            }

        }
    }
}
