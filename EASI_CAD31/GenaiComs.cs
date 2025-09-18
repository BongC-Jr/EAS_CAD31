using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.DatabaseServices.Filters;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows.Data;
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
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using acColor = Autodesk.AutoCAD.Colors.Color;
//using System.Globalization;


namespace EASI_CAD31
{
    public class GenaiComs
    {
        static Document actDoc = Application.DocumentManager.MdiActiveDocument;
        static Database aCurDB = actDoc.Database;

        static CCData ccData = new CCData();

        private static Point3d _lastSelectedPoint = Point3d.Origin; // Initialize with a default value
        // 2. (Optional but good practice) Provide a public static property to access it.
        // This gives controlled read access to the stored point.
        public static Point3d LastSelectedPoint
        {
            get { return _lastSelectedPoint; }
        }


      /**
      * Date added: 10 Sep 2025
      * Added by: Engr Bernardo Cabebe Jr.
      * Venue: 1508 CPTT, Chino Roces Ave, Makati City
      */
      [CommandMethod("AI_ColumnEstimate")]
      public static void ColumnEstimate()
      {
         Document iAcDoc = Application.DocumentManager.MdiActiveDocument; //Active document
         Database iCurDB = iAcDoc.Database;

         string userContentTxt = "";
         string convo = "";
         SEASTools sTools = new SEASTools();
         try
         {
            //Example of estimate params. The unit will be based on the given dimension of column
            //identified by its id.
            //_+DUOLDX,3280,28,12.7,10,0.00000785,mm,kg
            PromptStringOptions psoCEst = new PromptStringOptions("\nEstimate params: ");
            psoCEst.AllowSpaces = false;
            PromptResult prCEst = iAcDoc.Editor.GetString(psoCEst);
            if (prCEst.Status != PromptStatus.OK)
            {
               //iAcDoc.Editor.WriteMessage("\nCommand cancelled.");
               return;
            }
            string cEstParam = prCEst.StringResult;
            
            if (DataGlobal.isDevMessageOn)
            {
               iAcDoc.Editor.WriteMessage($"\nParams: {cEstParam}");
            }

            string[] arrStrParam = cEstParam.Split(',');
            int nParam = arrStrParam.Length;
            if (nParam < 8)
            {
               userContentTxt = $"user: {DataGlobal.UserMessage}";
               sTools.LogConversation(userContentTxt);

               convo = $"Insufficient parameters, {nParam}. I cannot perform the task.";
               sTools.LogConversation("model: " + convo);
               iAcDoc.Editor.WriteMessage($"\n{"Assistant: " + convo}\n");
               return;
            }

            /**
            * Index: 0        1    2  3    4  5          6  7
            *        _+DUOLDX,3280,28,12.7,10,0.00000785,mm,kg
            */
            //Select all objects in the layer
            string layerName = arrStrParam[0].Trim();
            double colLength = Convert.ToDouble(arrStrParam[1]);
            double nTies = Convert.ToDouble(arrStrParam[2]);
            double mainTieD = Convert.ToDouble(arrStrParam[3]);
            double crossTieD = Convert.ToDouble(arrStrParam[4]);
            double steelUW = Convert.ToDouble(arrStrParam[5]);
            string lenUnit = arrStrParam[6];
            string wtUnit = arrStrParam[7];
            
            bool layerExists = false;
            using (Transaction trLayer = iCurDB.TransactionManager.StartTransaction())
            {
               LayerTable lyrTbl = trLayer.GetObject(iCurDB.LayerTableId, OpenMode.ForRead) as LayerTable;
               foreach (ObjectId objId in lyrTbl)
               {
                  LayerTableRecord lyrtblrec;
                  lyrtblrec = trLayer.GetObject(objId, OpenMode.ForRead) as LayerTableRecord;
                   
                  if(lyrtblrec.Name == layerName)
                  {
                     layerExists = true;
                     break;
                  }
               }
            }
             
            if (!layerExists)
            {
               userContentTxt = $"user: {DataGlobal.UserMessage}";
               sTools.LogConversation(userContentTxt);

               convo = $"Column {layerName} does not exist. I cannot perform the task.";
               sTools.LogConversation("model: " + convo);
               iAcDoc.Editor.WriteMessage($"\n{"Assistant: " + convo}\n");
               return;
            }

            TypedValue[] tvColEst = new TypedValue[1];
            tvColEst.SetValue(new TypedValue((int)DxfCode.LayerName, layerName), 0);
            SelectionFilter sfColEst = new SelectionFilter(tvColEst);
            PromptSelectionResult psrColEst = iAcDoc.Editor.SelectAll(sfColEst);//GetSelection(sfColEst);
            if (psrColEst.Status != PromptStatus.OK) {
               userContentTxt = $"user: {DataGlobal.UserMessage}";
               sTools.LogConversation(userContentTxt);
               
               convo = $"Task column estimate cancelled.";
               sTools.LogConversation("model: " + convo);
               iAcDoc.Editor.WriteMessage($"\n{"Assistant: " + convo}\n");
               return;
            }
         
            SelectionSet ssColEst = psrColEst.Value;

            double cVolume = 0.0;
            double wtMt = 0.0;
            double wtCt = 0.0;
            double wtVb = 0.0;
            using (Transaction trColEst = iCurDB.TransactionManager.StartTransaction())
            {
               //Put beam marks in a list and sort it.
               foreach (SelectedObject soColEst in ssColEst)
               {
                  if (soColEst != null)
                  {
                     Entity enColEst = trColEst.GetObject(soColEst.ObjectId, OpenMode.ForRead) as Entity;
                     if (enColEst is Polyline)
                     {
                        Polyline plnColEst = enColEst as Polyline;
                        int intColor = plnColEst.ColorIndex;
                        switch (intColor)
                        {
                           case 1:
                              double cArea = plnColEst.Area;
                              cVolume = cArea * colLength;
                              break;
                           case 4:
                              //MT = main tie
                              double areaMT = 0.25 * Math.PI * mainTieD * mainTieD;
                              double lengthMT = plnColEst.Length;
                              double volMT = areaMT * lengthMT;
                              wtMt = nTies * volMT * steelUW;
                              break;
                           case 5:
                              //CT = cross tie
                              double areaCT = 0.25 * Math.PI * crossTieD * crossTieD;
                              double lengthCT = plnColEst.Length;
                              double volCT = areaCT * lengthCT;
                              wtCt += nTies * volCT * steelUW;
                              break;
                           default:
                              break;
                        }
                     }
                     if(enColEst is Circle)
                     {
                        Circle cirColEst = enColEst as Circle;
                        double diaVb = cirColEst.Diameter;
                        double areaVb = 0.25 * Math.PI * diaVb * diaVb;
                        double volVb = areaVb * colLength;
                        wtVb += volVb * steelUW;
                     }
                  }
               }
               trColEst.Commit();
            }

            try
            {
               userContentTxt = $"user: {DataGlobal.UserMessage}";
               // Create or append to the file
               sTools.LogConversation(userContentTxt);

               double totalWt = wtVb + wtMt + wtCt;
               //string strWtMT = wtMt.ToString("N",new CultureInfo("en-US"));
               convo = $"Column {layerName} has the following details: " +
                       $"length is {colLength.ToString()} " + lenUnit + ", " +
                       $"number of ties in vertical directions is {nTies.ToString()}, " +
                       $"main tie diameter is {mainTieD.ToString()}" + lenUnit + ", " +
                       $"cross tie diameter is {crossTieD.ToString()}, " + lenUnit + ", " +
                       $"AutoCAD objects of {ssColEst.Count.ToString()}," +
                       $"and the unit used is {lenUnit}-{wtUnit}. " +
                       $"The volume of concrete is {cVolume} {lenUnit}3, " +
                       $"the weight of vertical bars is {wtVb} {wtUnit}, " +
                       $"the weight of main tie is {wtMt} {wtUnit}, " +
                       $"the weight of cross ties is {wtCt} {wtUnit}, " +
                       $"and the total weight of steel is {totalWt} {wtUnit}";
                       //" ";
               //"Assistant: "
               sTools.LogConversation("model: " + convo);
               iAcDoc.Editor.WriteMessage("\nAssistant: " + convo + "\n");
            }
            catch (Autodesk.AutoCAD.Runtime.Exception exc)
            {
               userContentTxt = $"user: {DataGlobal.UserMessage}";
               // Create or append to the file
               sTools.LogConversation(userContentTxt);

               convo = $"Column Estimate system error: {exc.Message}";

               sTools.LogConversation("model: " + convo);
               iAcDoc.Editor.WriteMessage($"\n{"Assistant: " + convo}\n");
            }
            
         }
            catch (System.Exception ex)
            {
               userContentTxt = $"user: {DataGlobal.UserMessage}";
               sTools.LogConversation(userContentTxt);

               convo = $"Error in Column Estimate: {ex.Message}";
               sTools.LogConversation("model: " + convo);
               iAcDoc.Editor.WriteMessage($"\n{"Assistant: " + convo}\n");
            }
      }

      
      
      /**
      * Date added: 21 Jul 2025
      * Added by: Engr Bernardo Cabebe Jr.
      * Venue: 31H OOC, Bagumbayan, Quezon City
      */
      [CommandMethod("AI_DrawColumnSection")]
      public static void DrawColumnSection()
      {
         //Example Column params:
         //16,24,4,6,0.75,0.5,1.57,in
         PromptStringOptions psoDCS = new PromptStringOptions("\nColumn params: ");
         psoDCS.AllowSpaces = false;
         PromptResult prDCS = actDoc.Editor.GetString(psoDCS);
         if (prDCS.Status != PromptStatus.OK)
         {
            actDoc.Editor.WriteMessage("\nCommand cancelled.");
            return;
         }
         string colParam = prDCS.StringResult;

         if (DataGlobal.isDevMessageOn)
         {
            actDoc.Editor.WriteMessage($"\nDraw Column Section");
            actDoc.Editor.WriteMessage($"\nColumn parameters: {colParam}");
            actDoc.Editor.WriteMessage($"\n-------------------\n");
         }

         string[] arrStrParam = colParam.Split(',');
         double[] arrDblParam = new double[arrStrParam.Length];
         string unitUsed = "";
         for (int i = 0; i < arrStrParam.Length; i++)
         {
            if(i == 7){
               //0  1  2 3 4    5   6    7  <-- index number
               //16,24,4,6,0.75,0.5,1.57,in <-- Column params
               //                         .
               //                        / \
               //                         |
               //                         |
               //                         + index 7    
               unitUsed = arrStrParam[i];
            }
            else
            {
              arrDblParam[i] = Convert.ToDouble(arrStrParam[i]);
            }
         }
         if (DataGlobal.isDevMessageOn)
         {
            for (int aa = 0; aa < arrDblParam.Length; aa++)
            {
               actDoc.Editor.WriteMessage($"\narrDblParam[{aa}] = {arrDblParam[aa]}");
            }
         }
         
         SEASTools sTools = new SEASTools();
         string mixText = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
         string shuffledText = sTools.ShuffleString(mixText);
         string layerName = "_+" + shuffledText.Substring(0, 6);
         sTools.CreateLayer(layerName, 3, true);

         Editor curEd = actDoc.Editor;
         Point2d curEdCtr = curEd.GetCurrentView().CenterPoint;
         double cW = arrDblParam[0]; // column width
         double cH = arrDblParam[1]; // column depth
         Point2d pA1 = new Point2d(curEdCtr.X - cW / 2, curEdCtr.Y + cH / 2);
         Point2d pA2 = new Point2d(pA1.X + cW, pA1.Y);
         Point2d pA3 = new Point2d(pA2.X, pA2.Y - cH);
         Point2d pA4 = new Point2d(pA3.X - cW, pA3.Y);
         Point2d pA5 = new Point2d(pA1.X, pA1.Y);
         using (Transaction tr = aCurDB.TransactionManager.StartTransaction())
         {
            try
            {
               BlockTable blktbl;
               blktbl = tr.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
               BlockTableRecord blktbr;
               blktbr = tr.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

               Polyline plCol = new Polyline();
               plCol.AddVertexAt(0, pA1, 0, 0, 0);
               plCol.AddVertexAt(1, pA2, 0, 0, 0);
               plCol.AddVertexAt(2, pA3, 0, 0, 0);
               plCol.AddVertexAt(3, pA4, 0, 0, 0);
               plCol.AddVertexAt(4, pA5, 0, 0, 0);
               plCol.Color = acColor.FromColorIndex(ColorMethod.ByAci, 1);
               plCol.Layer = layerName;
               blktbr.AppendEntity(plCol);
               tr.AddNewlyCreatedDBObject(plCol, true);
               tr.Commit();
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
               tr.Abort();
               curEd.WriteMessage("\nError A2W: " + ex.Message + "\n");
            }
            catch (System.Exception ex)
            {
               tr.Abort();
               curEd.WriteMessage("\nError S5Y: " + ex.Message + "\n");
            }
         }

         double verBarDia = arrDblParam[4];
         double tieBarDia = arrDblParam[5];
         double concCover = arrDblParam[6];

         //Create text
         string aiTextLayer = "AI_Text";
         sTools.CreateLayer(aiTextLayer, 1, false);

         Vector2d v2yTexti2q = pA1.GetVectorTo(pA4);
         v2yTexti2q = v2yTexti2q.GetNormal();
         Vector2d v2xTexti2q = pA1.GetVectorTo(pA2);
         v2xTexti2q = v2xTexti2q.GetNormal();
         double txtDist = pA4.GetDistanceTo(pA3) / 2;
         Point2d txtPti2q = pA4 + v2xTexti2q * txtDist;
         txtPti2q = txtPti2q + v2yTexti2q * verBarDia * 3.5;
         
         using (Transaction tr = aCurDB.TransactionManager.StartTransaction())
         {
            BlockTable bt = tr.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
            BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

            // Create a new DBText entity
            DBText text = new DBText();
            text.SetDatabaseDefaults();
            text.TextString = layerName;
            text.Layer = aiTextLayer;
            //text.Layer = layerName;
            text.Position = new Point3d(txtPti2q.X, txtPti2q.Y, 0);
            text.Height = verBarDia * 2;
            text.HorizontalMode = TextHorizontalMode.TextCenter;
            text.VerticalMode = TextVerticalMode.TextTop;
            text.AlignmentPoint = text.Position;
            text.Color = acColor.FromColorIndex(ColorMethod.ByAci, 251); 

            // Add the text to the current space
            btr.AppendEntity(text);
            tr.AddNewlyCreatedDBObject(text, true);

            tr.Commit();
         }

         Point3d cirCPt1 = new Point3d(pA1.X + concCover + tieBarDia + verBarDia * 0.5, pA1.Y - concCover - tieBarDia - verBarDia * 0.5, 0.0);
         Point3d cirCPt2 = new Point3d(pA2.X - concCover - tieBarDia - verBarDia * 0.5, pA2.Y - concCover - tieBarDia - verBarDia * 0.5, 0.0);
         Point3d cirCPt3 = new Point3d(pA3.X - concCover - tieBarDia - verBarDia * 0.5, pA3.Y + concCover + tieBarDia + verBarDia * 0.5, 0.0);
         Point3d cirCPt4 = new Point3d(pA4.X + concCover + tieBarDia + verBarDia * 0.5, pA4.Y + concCover + tieBarDia + verBarDia * 0.5, 0.0);
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
               cVerBar.Color = acColor.FromColorIndex(ColorMethod.ByAci, 2);
               cVerBar.Layer = layerName;

               bltrec.AppendEntity(cVerBar);
               trVertBar.AddNewlyCreatedDBObject(cVerBar, true);
            }

            trVertBar.Commit();
         }
         

         //Draw main tie
         GenaiComs genaico = new GenaiComs();
         genaico.DrawCloseTie(cirCPt1, cirCPt3, verBarDia, tieBarDia, layerName, 4);

         // Draw cross ties
         double Nb = arrDblParam[2]; // number of vertical bars
         double Nh = arrDblParam[3]; // number of horizontal bars
         double bs = cirCPt1.DistanceTo(cirCPt2); // bar spacing
         double db = bs / (Nb - 1); // distance spacing
         double hs = cirCPt1.DistanceTo(cirCPt4); // vertical distance
         double dh = hs/(Nh - 1); // vertical distance spacing
         int Nt = 2*(int)Nb+2*(int)Nh-4; // number of bars
         
         Vector2d v2yTextk2q = pA1.GetVectorTo(pA4);
         v2yTextk2q = v2yTextk2q.GetNormal();
         Point2d txtPtk2q = pA4 + v2yTextk2q * verBarDia * 12;
         
         using (Transaction tr = aCurDB.TransactionManager.StartTransaction())
         {
            BlockTable bt = tr.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
            BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
            
            // Create a new MText entity
            MText mText = new MText();
            mText.SetDatabaseDefaults();
            mText.Contents = $"DIMENSION  : {cW}X{cH}{unitUsed}\n"+
                             $"VERT. BARS : {Nt} - {verBarDia}%%C\n" +
                             $"COL. TIES  : {tieBarDia}%%C - N1 @ S1, REST @ S2 O.C. TO C.L.\n" +
                             $"JOINT CONF.: {concCover}%%C - N3 @ S3 O.C.";
            //mText.Layer = aiTextLayer;
            mText.Layer = layerName;
            mText.Location = new Point3d(txtPtk2q.X, txtPtk2q.Y, 0);
            mText.TextHeight = verBarDia * 3;
            mText.Attachment = AttachmentPoint.TopLeft;
            mText.Color = acColor.FromColorIndex(ColorMethod.ByAci, 3); 

            // Add the text to the current space
            btr.AppendEntity(mText);
            tr.AddNewlyCreatedDBObject(mText, true);

            tr.Commit();
         }


         List<Point3d> arrBTP3d = new List<Point3d>();
         List<Point3d> arrBBP3d = new List<Point3d>();
         List<Point3d> arrHLP3d = new List<Point3d>();
         List<Point3d> arrHRP3d = new List<Point3d>();
         using (Transaction tr = aCurDB.TransactionManager.StartTransaction())
         {
            BlockTable blktbl;
            blktbl = tr.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
            BlockTableRecord blktbr;
            blktbr = tr.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

            Vector3d v3dB = cirCPt1.GetVectorTo(cirCPt2);
            v3dB = v3dB.GetNormal();
            int aa;
            for (aa = 0; aa < (Nb - 2); aa++)
            {
               Point3d aaPti = cirCPt1 + v3dB * db * (aa + 1);
               arrBTP3d.Add(aaPti);
               Circle aaVBarT = new Circle();
               aaVBarT.SetDatabaseDefaults();
               aaVBarT.Center = aaPti;
               aaVBarT.Diameter = verBarDia;
               aaVBarT.Color = acColor.FromColorIndex(ColorMethod.ByAci, 2);
               aaVBarT.Layer = layerName;
               blktbr.AppendEntity(aaVBarT);
               tr.AddNewlyCreatedDBObject(aaVBarT, true);

               Point3d aaPtj = cirCPt4 + v3dB * db * (aa + 1);
               arrBBP3d.Add(aaPtj);
               Circle aaVBarB = new Circle();
               aaVBarB.SetDatabaseDefaults();
               aaVBarB.Center = aaPtj;
               aaVBarB.Diameter = verBarDia;
               aaVBarB.Color = acColor.FromColorIndex(ColorMethod.ByAci, 2);
               aaVBarB.Layer = layerName;
               blktbr.AppendEntity(aaVBarB);
               tr.AddNewlyCreatedDBObject(aaVBarB, true);

            }

            Vector3d v3dH = cirCPt1.GetVectorTo(cirCPt4);
            v3dH = v3dH.GetNormal();
            int bb;
            for (bb = 0; bb < (Nh - 2); bb++)
            {
               Point3d bbPti = cirCPt1 + v3dH * dh * (bb + 1);
               arrHLP3d.Add(bbPti);
               Circle bbVBarL = new Circle();
               bbVBarL.SetDatabaseDefaults();
               bbVBarL.Center = bbPti;
               bbVBarL.Diameter = verBarDia;
               bbVBarL.Color = acColor.FromColorIndex(ColorMethod.ByAci, 2);
               bbVBarL.Layer = layerName;
               blktbr.AppendEntity(bbVBarL);
               tr.AddNewlyCreatedDBObject(bbVBarL, true);

               Point3d bbPtj = cirCPt2 + v3dH * dh * (bb + 1);
               arrHRP3d.Add(bbPtj);
               Circle bbVBarR = new Circle();
               bbVBarR.SetDatabaseDefaults();
               bbVBarR.Center = bbPtj;
               bbVBarR.Diameter = verBarDia;
               bbVBarR.Color = acColor.FromColorIndex(ColorMethod.ByAci, 2);
               bbVBarR.Layer = layerName;
               blktbr.AppendEntity(bbVBarR);
               tr.AddNewlyCreatedDBObject(bbVBarR, true);
            }
            tr.Commit();
         }
         
         Point3d[][] arrPairBT = genaico.ProcessPoint3DList(arrBTP3d);
         Point3d[][] arrPairBB = genaico.ProcessPoint3DList(arrBBP3d);
         int cc;
         for (cc = 0; cc < arrPairBT.Length; cc++)
         {
            if (arrPairBT[cc].Length == 2 && arrPairBB[cc].Length == 2)
            {
               Point3d ccPt1 = arrPairBT[cc][0];
               Point3d ccPt2 = arrPairBB[cc][1];
               genaico.DrawCloseTie(ccPt1, ccPt2, verBarDia, tieBarDia, layerName, 5);
            }
            else
            {
               Point3d ccPt1 = arrPairBT[cc][0];
               Point3d ccPt2 = arrPairBB[cc][0];
               genaico.DrawOpenTieH(ccPt1, ccPt2, verBarDia, tieBarDia, layerName, 5);
            }
         }

         Point3d[][] arrPairHL = genaico.ProcessPoint3DList(arrHLP3d);
         Point3d[][] arrPairHR = genaico.ProcessPoint3DList(arrHRP3d);
         int dd;
         for(dd = 0; dd < arrPairHL.Length; dd++)
         {
            if (arrPairHL[dd].Length == 2 && arrPairHR[dd].Length == 2)
            {
               Point3d ddPt1 = arrPairHL[dd][0];
               Point3d ddPt2 = arrPairHR[dd][1];
               genaico.DrawCloseTie(ddPt1, ddPt2, verBarDia, tieBarDia, layerName, 5);
            }
            else
            {
               Point3d ddPt1 = arrPairHL[dd][0];
               Point3d ddPt2 = arrPairHR[dd][0];
               genaico.DrawOpenTieB(ddPt1, ddPt2, verBarDia, tieBarDia, layerName, 5);
            }
         }

         string userContentTxt = "";
         string convo = "";
         
         userContentTxt = $"user: {DataGlobal.UserMessage}";
         // Create or append to the file
         sTools.LogConversation(userContentTxt);
         
         convo = $"The column {layerName} with width of {cW} {unitUsed}, depth of {cH} {unitUsed}, Nb of {arrDblParam[2]}, an Nh of {arrDblParam[3]}, and total number of rebar {Nt} " +
                 $"is drawn at the center of the current view. The vertical bar diameter is {verBarDia} {unitUsed}, tie bar diameter is {tieBarDia} {unitUsed}, and concrete cover is {concCover} {unitUsed}.";
         
         sTools.LogConversation("model: " + convo);
         actDoc.Editor.WriteMessage($"\n{"Assistant: " + convo}\n");
      }
      private Point3d[][] ProcessPoint3DList(List<Point3d> liP3d)
      {
         List<Point3d[]> arrPair = new List<Point3d[]>();

         while (liP3d.Count > 0)
         {
            if (liP3d.Count >= 2)
            {
               // Get the first two items
               Point3d[] pair = new Point3d[2];
               pair[0] = liP3d[0];
               pair[1] = liP3d[1];

               // Add the pair to arrPair
               arrPair.Add(pair);

               // Remove the first two items
               liP3d.RemoveRange(0, 2);
            }
            else if (liP3d.Count == 1)
            {
               // Get the last item
               Point3d[] pair = new Point3d[1];
               pair[0] = liP3d[0];

               // Add the pair to arrPair
               arrPair.Add(pair);

               // Remove the last item
               liP3d.Clear();
            }

            if (liP3d.Count > 0)
            {
               if (liP3d.Count >= 2)
               {
                  // Get the last two items
                  Point3d[] pair = new Point3d[2];
                  pair[0] = liP3d[liP3d.Count - 2];
                  pair[1] = liP3d[liP3d.Count - 1];

                  // Add the pair to arrPair
                  arrPair.Add(pair);

                  // Remove the last two items
                  liP3d.RemoveRange(liP3d.Count - 2, 2);
               }
               else if (liP3d.Count == 1)
               {
                  // Get the last item
                  Point3d[] pair = new Point3d[1];
                  pair[0] = liP3d[0];

                  // Add the pair to arrPair
                  arrPair.Add(pair);

                  // Remove the last item
                  liP3d.Clear();
               }
            }
         }

         return arrPair.ToArray();
      }
      private void DrawCloseTie(Point3d cirCPt1, Point3d cirCPt2, double verBarDia, double tieBarDia, string layerName, short colorIndex = 5)
      {
         using (Transaction trCloseTie = aCurDB.TransactionManager.StartTransaction())
            {
                BlockTable blktbl;
                blktbl = trCloseTie.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord bltrec;
                bltrec = trCloseTie.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

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

                    plSeg1.Color = acColor.FromColorIndex(ColorMethod.ByAci, colorIndex);
                    plSeg1.Layer = layerName;

                    bltrec.AppendEntity(plSeg1);
                    trCloseTie.AddNewlyCreatedDBObject(plSeg1, true);
                }

                trCloseTie.Commit();
            }
      }
      private void DrawOpenTieH(Point3d cirCPt1, Point3d cirCPt2, double verBarDia, double tieBarDia, string layerName, short colorIndex = 5)
      {
         using (Transaction trCloseTie = aCurDB.TransactionManager.StartTransaction())
         {
            BlockTable blktbl;
            blktbl = trCloseTie.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
            BlockTableRecord bltrec;
            bltrec = trCloseTie.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

            Matrix3d curUCSMatrix = actDoc.Editor.CurrentUserCoordinateSystem;
            CoordinateSystem3d curUCS = curUCSMatrix.CoordinateSystem3d;

            Point3d plpt02u = new Point3d(cirCPt1.X + verBarDia * 0.5 + tieBarDia, cirCPt1.Y, cirCPt1.Z);
            Point3d plpt02 = plpt02u.TransformBy(Matrix3d.Rotation(0.785398, curUCS.Zaxis, cirCPt1));

            Point3d plpt01u = new Point3d(plpt02.X + 6 * tieBarDia, plpt02.Y, plpt02.Z);
            Point3d plpt01 = plpt01u.TransformBy(Matrix3d.Rotation(-0.785398, curUCS.Zaxis, plpt02));

            Point3d appPt1 = new Point3d(cirCPt1.X, cirCPt2.Y, cirCPt1.Z); //apparent point 1
            Point3d appPt2 = new Point3d(cirCPt2.X, cirCPt1.Y, cirCPt1.Z); //apparent point 2
            Point3d plpt04 = new Point3d(appPt1.X - verBarDia * 0.5 - tieBarDia, appPt1.Y, appPt1.Z);
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

            Vector3d v3P56 = plpt04.GetVectorTo(appPt1);
            v3P56 = v3P56.GetNormal();
            Point3d plpt06a = plpt05 + v3P56 * tieBarDia * 8;

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

               plSeg1.AddVertexAt(5, new Point2d(plpt06a.X, plpt06a.Y), 0, 0, 0);
               //plSeg1.AddVertexAt(5, new Point2d(plpt06.X, plpt06.Y), 0.414214, 0, 0);
               //plSeg1.AddVertexAt(6, new Point2d(plpt07.X, plpt07.Y), 0, 0, 0);
               // 
               //plSeg1.AddVertexAt(7, new Point2d(plpt08.X, plpt08.Y), 0.414214, 0, 0);
               //plSeg1.AddVertexAt(8, new Point2d(plpt09.X, plpt09.Y), 0, 0, 0);
               // 
               //plSeg1.AddVertexAt(9, new Point2d(plpt10.X, plpt10.Y), 0.668179, 0, 0);
               //plSeg1.AddVertexAt(10, new Point2d(plpt11.X, plpt11.Y), 0, 0, 0);
               // 
               //plSeg1.AddVertexAt(11, new Point2d(plpt12.X, plpt12.Y), 0, 0, 0);

               plSeg1.Color = acColor.FromColorIndex(ColorMethod.ByAci, colorIndex);
               plSeg1.Layer = layerName;

               bltrec.AppendEntity(plSeg1);
               trCloseTie.AddNewlyCreatedDBObject(plSeg1, true);
            }
            trCloseTie.Commit();
         }
      }
      private void DrawOpenTieB(Point3d cirCPt1, Point3d cirCPt2, double verBarDia, double tieBarDia, string layerName, short colorIndex = 5)
      {
         using (Transaction trCloseTie = aCurDB.TransactionManager.StartTransaction())
         {
            BlockTable blktbl;
            blktbl = trCloseTie.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;
            BlockTableRecord bltrec;
            bltrec = trCloseTie.GetObject(blktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

            Matrix3d curUCSMatrix = actDoc.Editor.CurrentUserCoordinateSystem;
            CoordinateSystem3d curUCS = curUCSMatrix.CoordinateSystem3d;

            Point3d plpt02u = new Point3d(cirCPt1.X + verBarDia * 0.5 + tieBarDia, cirCPt1.Y, cirCPt1.Z);
            Point3d plpt02 = plpt02u.TransformBy(Matrix3d.Rotation(0.785398, curUCS.Zaxis, cirCPt1));

            Point3d plpt01u = new Point3d(plpt02.X + 6 * tieBarDia, plpt02.Y, plpt02.Z);
            Point3d plpt01 = plpt01u.TransformBy(Matrix3d.Rotation(-0.785398, curUCS.Zaxis, plpt02));

            Point3d appPt1 = new Point3d(cirCPt1.X, cirCPt2.Y, cirCPt1.Z); //apparent point 1
            Point3d appPt2 = new Point3d(cirCPt2.X, cirCPt1.Y, cirCPt1.Z); //apparent point 2
            Point3d plpt04 = new Point3d(appPt1.X - verBarDia * 0.5 - tieBarDia, appPt1.Y, appPt1.Z);
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

            Vector3d v3P56 = plpt09.GetVectorTo(appPt2);
            v3P56 = v3P56.GetNormal();
            Point3d plpt07a = plpt08 + v3P56 * tieBarDia * 8;

            Polyline plSeg1;
            using (plSeg1 = new Polyline())
            {
               plSeg1.SetDatabaseDefaults();
               //plSeg1.AddVertexAt(0, new Point2d(plpt01.X, plpt01.Y), 0, 0, 0);

               //double arcRad = verBarDia * 0.5 + tieBarDia;
               //Point3d plpt03 = new Point3d(cirCPt1.X - verBarDia * 0.5 - tieBarDia, cirCPt1.Y, cirCPt1.Z);
               //plSeg1.AddVertexAt(1, new Point2d(plpt02.X, plpt02.Y), 0.668179, 0, 0);
               //plSeg1.AddVertexAt(2, new Point2d(plpt03.X, plpt03.Y), 0, 0, 0);
               // 
               //plSeg1.AddVertexAt(3, new Point2d(plpt04.X, plpt04.Y), 0.414214, 0, 0);
               //plSeg1.AddVertexAt(4, new Point2d(plpt05.X, plpt05.Y), 0, 0, 0);
               //
               //plSeg1.AddVertexAt(5, new Point2d(plpt06.X, plpt06.Y), 0.414214, 0, 0);
               //plSeg1.AddVertexAt(6, new Point2d(plpt07.X, plpt07.Y), 0, 0, 0);

               plSeg1.AddVertexAt(0, new Point2d(plpt07a.X, plpt07a.Y), 0, 0, 0);
                
               plSeg1.AddVertexAt(1, new Point2d(plpt08.X, plpt08.Y), 0.414214, 0, 0);
               plSeg1.AddVertexAt(2, new Point2d(plpt09.X, plpt09.Y), 0, 0, 0);
                
               plSeg1.AddVertexAt(3, new Point2d(plpt10.X, plpt10.Y), 0.668179, 0, 0);
               plSeg1.AddVertexAt(4, new Point2d(plpt11.X, plpt11.Y), 0, 0, 0);
                
               plSeg1.AddVertexAt(5, new Point2d(plpt12.X, plpt12.Y), 0, 0, 0);

               plSeg1.Color = acColor.FromColorIndex(ColorMethod.ByAci, colorIndex);
               plSeg1.Layer = layerName;

               bltrec.AppendEntity(plSeg1);
               trCloseTie.AddNewlyCreatedDBObject(plSeg1, true);
            }
            trCloseTie.Commit();
         }
      }

      


      /**
       * Date added: 19 Jul 2025
       * Addede by: Engr Bernardo Cabebe Jr.
       * Venue: Subway Eastwood, Bagumbayan, Quezon City
       */
      [CommandMethod("AI_DrawLineM01")]
        public static void AI_DrawLineM01()
        {
             PromptPointOptions ppoStart = new PromptPointOptions("\nEnter start point: ");
             PromptPointResult pprStart = actDoc.Editor.GetPoint(ppoStart);
             if (pprStart.Status != PromptStatus.OK)
             {
                 actDoc.Editor.WriteMessage("\nCommand cancelled.");
                 return;
             }
             Point3d startPt = pprStart.Value;
 
             PromptPointOptions ppoEnd = new PromptPointOptions("\nEnter end point: ");
             PromptPointResult pprEnd = actDoc.Editor.GetPoint(ppoEnd);
             if (pprEnd.Status != PromptStatus.OK)
             {
                 actDoc.Editor.WriteMessage("\nCommand cancelled.");
                 return;
             }
             Point3d endPt = pprEnd.Value;
 
             PromptIntegerOptions pioColor = new PromptIntegerOptions("\nEnter color index (1-255): ");
             pioColor.AllowNegative = false;
             pioColor.AllowZero = false;
             pioColor.DefaultValue = 73; // Default to pale green
             PromptIntegerResult pirColor = actDoc.Editor.GetInteger(pioColor);
             if (pirColor.Status != PromptStatus.OK)
             {
                 actDoc.Editor.WriteMessage("\nCommand cancelled.");
                 return;
             }
             int colorIndex = pirColor.Value;
             if (colorIndex < 1 || colorIndex > 255)
             {
                actDoc.Editor.WriteMessage("\nInvalid color index. Please enter a value between 1 and 255.");
                return;
             }

             double lineLength = 0.0;
             using (Transaction tr = aCurDB.TransactionManager.StartTransaction())
             {
                Line lineObj = new Line(startPt, endPt);
                lineObj.ColorIndex = colorIndex;
                lineLength = lineObj.Length;
                BlockTable bt = (BlockTable)tr.GetObject(aCurDB.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btrModelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                btrModelSpace.AppendEntity(lineObj);
                tr.AddNewlyCreatedDBObject(lineObj, true);
                tr.Commit();
             }
             string filePath = Path.Combine(DataGlobal.convofilepath, "conversation.txt");

             string userContentTxt = $"user: {DataGlobal.UserMessage}";
             // Create or append to the file
             using (StreamWriter writer = File.AppendText(filePath))
             {
                 writer.WriteLine(userContentTxt);
             }
 
             string contentTxt = $"model: Drawn a line from {startPt} to {endPt} with color index of {colorIndex}. The length of the line is {lineLength}";
             // Create or append to the file
             using (StreamWriter writer = File.AppendText(filePath))
             {
                 writer.WriteLine(contentTxt);
             }
             actDoc.Editor.WriteMessage($"\nLine drawn from {startPt} to {endPt}.");
        }


      /**
       * Date added: 18 Jul 2025
       * Addede by: Engr Bernardo Cabebe Jr.
       * Venue: Subway Eastwood, Bagumbayan, Quezon City
       */
      [CommandMethod("AI_SelectPointM01")]
        public static void AI_SelectPointM01()
        {
             PromptIntegerOptions pioPtColor = new PromptIntegerOptions("\nEnter color index: ");
             pioPtColor.AllowNegative = false;
             pioPtColor.AllowZero = false;
             pioPtColor.DefaultValue = 12; // Default to bright Red
             
             string filePath = Path.Combine(DataGlobal.convofilepath, "conversation.txt");
             string contentTxt = "";
             PromptIntegerResult pirPtColor = actDoc.Editor.GetInteger(pioPtColor);
             if (pirPtColor.Status != PromptStatus.OK)
             {
                 actDoc.Editor.WriteMessage("\nAssistant: Command cancelled.");
                 string userContentTxt = $"user: {DataGlobal.UserMessage}";
                 // Create or append to the file
                 using (StreamWriter writer = File.AppendText(filePath))
                 {
                     writer.WriteLine(userContentTxt);
                 }
      
                 contentTxt = $"model: Command cancelled.";
                 // Create or append to the file
                 using (StreamWriter writer = File.AppendText(filePath))
                 {
                     writer.WriteLine(contentTxt);
                 }
                 return;
             }
             int iPtColor = pirPtColor.Value;

             TypedValue[] tvFilterG5w = new TypedValue[2];
             tvFilterG5w[0] = new TypedValue((int)DxfCode.Start, "POINT");
             tvFilterG5w[1] = new TypedValue((int)DxfCode.Color, iPtColor);

             PromptSelectionResult psrPtObj = actDoc.Editor.SelectAll(new SelectionFilter(tvFilterG5w));
             if (psrPtObj.Status != PromptStatus.OK)
             {
                 actDoc.Editor.WriteMessage($"\nAssistant: There is no object point found with the specified color index {iPtColor}.");
                 
                 string userContentTxt = $"user: {DataGlobal.UserMessage}";
                 // Create or append to the file
                 using (StreamWriter writer = File.AppendText(filePath))
                 {
                     writer.WriteLine(userContentTxt);
                 }
      
                 contentTxt = $"model: No point objects found with the specified color index {iPtColor}.";
                 // Create or append to the file
                 using (StreamWriter writer = File.AppendText(filePath))
                 {
                     writer.WriteLine(contentTxt);
                 }
                 return;
             }
    
             using (Transaction trPtObj = aCurDB.TransactionManager.StartTransaction())
             {
                SelectionSet ssPtObj = psrPtObj.Value;
                if (ssPtObj.Count == 0)
                {
                   actDoc.Editor.WriteMessage($"\nNo point objects found with the specified color index {iPtColor}.");
                   filePath = Path.Combine(DataGlobal.convofilepath, "conversation.txt");
                   contentTxt = $"model: No point objects found with the specified color index {iPtColor}.";
                   // Create or append to the file
                   File.AppendAllText(filePath, contentTxt + Environment.NewLine);
                   return;
                }
                // Get the first selected point object
                ObjectId ptObjId = ssPtObj[0].ObjectId;
                DBPoint ptObj = (DBPoint)trPtObj.GetObject(ptObjId, OpenMode.ForRead);
                // Store the point's position
                _lastSelectedPoint = ptObj.Position;
    
                string userContentTxt = $"user: {DataGlobal.UserMessage}";
                // Create or append to the file
                using (StreamWriter writer = File.AppendText(filePath))
                {
                    writer.WriteLine(userContentTxt);
                }
     
                contentTxt = $"model: Object is Point, Color index is {iPtColor}, Point coordinate is {_lastSelectedPoint}, and Point Id is {ptObjId}";
                // Create or append to the file
                using (StreamWriter writer = File.AppendText(filePath))
                {
                    writer.WriteLine(contentTxt);
                }
    
                // Optionally, you can display the point's position in the command line
                actDoc.Editor.WriteMessage($"\nAssistant: The coordinates of a point with "+
                                           $"color index of {iPtColor} is { _lastSelectedPoint}\n");
                trPtObj.Commit();
             }
        }
    }
}
