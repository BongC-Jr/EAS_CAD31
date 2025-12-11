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

namespace EASI_CAD31
{
    public class CommonCom20
    {
        static Document actDoc = Application.DocumentManager.MdiActiveDocument;
        static Database aCurDB = actDoc.Database;

        static CCData ccData = new CCData();

        static int colorNoG = 1;

        static string oldTxtG = "";

        static int txtColorG = 7;
        static int layerColG = 1;

        //Default sheet address
        static string shAddrG = "Sheet1!A2:O2";

        static string textPatternG = "*";

        //Footing
        static string sheetsUrl = "https://docs.google.com/spreadsheets/d/1DYTcredribXAXBho2Z4k7aD8y-KF-3LZTGHARk_EDMM/edit?gid=0#gid=0";
        static string ftgSpreadSheetIDG = "1DYTcredribXAXBho2Z4k7aD8y-KF-3LZTGHARk_EDMM";
        static int startRowG = 4;
        static int endRowG = 10;
        static double textHgtG = 250;
        static double rowHgtScaleG = 2.5;

        /**
         * Date added: 11 Jan. 2024
         * Addede by: Engr Bernardo Cabebe Jr.
         * Venue: 3407 Cityland Pasong Tamo Tower
         */
        [CommandMethod("ICC_InitializedCommonCommands")]
        public static void InitializedCommonCommand()
        {
            /**
             *https://docs.google.com/spreadsheets/d/1OCmwmrTTy9627APO6gEy-TxIPBj-q0otIDWxl5RcgPU/edit?gid=0#gid=0
             */
            DataGlobal.licSSheetsID = "1OCmwmrTTy9627APO6gEy-TxIPBj-q0otIDWxl5RcgPU";

            /**
             * Google sheet for Get Polyline Area
             * https://docs.google.com/spreadsheets/d/15r1dGpnTtlsgRPm8s2u170bMPTdAqSELk-XBX72QDPk/edit?gid=0#gid=0
             */
            DataGlobal.gpaSSheetID = "15r1dGpnTtlsgRPm8s2u170bMPTdAqSELk-XBX72QDPk";

            GoogleSheetsV4 gsv4 = new GoogleSheetsV4();
            gsv4.GoogleSheetInit(); //Initialize Google Sheets API

            //GoogleSheetsV4.GoogleSheetInit();


            /**
             * Sample content of SEASLicV21.txt
             * --------------------------------------
             * T24A09M51R23F   (<- this is the first line in text file)
             * A16:I16
             * 
             */
            string licPath = DataGlobal.credLicPath + "\\SEASLicV21.txt";
            actDoc.Editor.WriteMessage("\nLP: {0}", licPath);

            StreamReader srSTx = new StreamReader(licPath); //stream reader SEAS text file
            List<string> lsLine = new List<string>();
            string strLine = srSTx.ReadLine();
            while (strLine != null)
            {
                lsLine.Add(strLine);
                strLine = srSTx.ReadLine();
            }
            srSTx.Close();

            /**
             * DISTRIBUTION DATA
             */
            //1.License
            ccData.setLicense(lsLine[0]);                   //When releasing with different license, this line
            string strAddrP2w = "CADLicense!" + lsLine[1];  //and this line must be updated.

            SpreadsheetsResource.ValuesResource.GetRequest grLic = DataGlobal.sheetsService.Spreadsheets.Values.Get(DataGlobal.licSSheetsID, strAddrP2w);
            ValueRange vrLic = grLic.Execute();
            IList<IList<object>> iioLic = vrLic.Values;
            actDoc.Editor.WriteMessage("\nLN: " + iioLic[0][1].ToString());

            if (iioLic[0][1].ToString() == ccData.getLicense() && iioLic[0][3].ToString() == "notrevoked")
            {
                ccData.setLicenseStatus(true);
                ccData.setLicenseNote(iioLic[0][4].ToString());
                actDoc.Editor.WriteMessage("\n" + iioLic[0][4].ToString());
            }
            else
            {
                actDoc.Editor.WriteMessage("\n" + iioLic[0][4].ToString());
                return;
            }

            /**
             * 2. Slab design
             *    https://docs.google.com/spreadsheets/d/13-99FEDfAhzm5bDDbZWu8_rACk1M9FSyzbOnUYxO6KM/edit#gid=0
             */
            string strGSheetID = iioLic[0][7].ToString(); // Google sheet as EDITOR
            DataGlobal.slabGSheetViewer = iioLic[0][6].ToString(); //Google sheet as VIEWER
            string strGSheetViewer = DataGlobal.slabGSheetViewer;

            actDoc.Editor.WriteMessage("\nSlab Google sheet viewer: {0}", strGSheetViewer);
            ccData.setSlabGSheetID(strGSheetID);

            //Initialize slab design data
            //slab thickness, bema depth, bar diameter, superimposed load, live load, others, concrete fc, steel fy, sw factor, sd factor, ll factor, others factor
            string slabDesignData = "150,600,10,3.47,0.48,0.0,27.58,275.8,1.2,1.2,1.6,1.0,250,1";
            ccData.setSlabDesignData(slabDesignData);

            //BBS - Bar bending schedule
            DataGlobal.bbsGSheet = "https://docs.google.com/spreadsheets/d/1xmapL82ZIP0CPuvVMQmPWKPLXpsyzOANs6x5gSyPdFQ/edit#gid=2098576256";
            DataGlobal.bbsGSheetId = "1xmapL82ZIP0CPuvVMQmPWKPLXpsyzOANs6x5gSyPdFQ";
            actDoc.Editor.WriteMessage("\nBBS Google sheet: {0}", DataGlobal.bbsGSheet);
            actDoc.Editor.WriteMessage("\nBBS: Copy the google sheet above and make your copy as your working Google Sheet.");

            /**
             * AutoCAD AI
             * https://docs.google.com/spreadsheets/d/17lo6yeCzmVMpghV7RKUqi0vg2kciqFxcROjY312s5rY/edit?gid=0#gid=0
             */
            DataGlobal.aiConvoGSheetId = "17lo6yeCzmVMpghV7RKUqi0vg2kciqFxcROjY312s5rY";
            IList<IList<object>> iioRangeAddress = gsv4.GetRange("ProgramParameters!B2", DataGlobal.aiConvoGSheetId);
            string rangeAddress = iioRangeAddress[0][0].ToString();
            DataGlobal.trainingConversation = gsv4.GetRange(rangeAddress, DataGlobal.aiConvoGSheetId);

            //CAI_2
            IList<IList<object>> iioRangeAddress2 = gsv4.GetRange("ProgramParameters!B3", DataGlobal.aiConvoGSheetId);
            string rangeAddress2 = iioRangeAddress2[0][0].ToString();
            DataGlobal.trainingConvCai2 = gsv4.GetRange(rangeAddress2, DataGlobal.aiConvoGSheetId);

            actDoc.Editor.WriteMessage("\nSEAS commands data initialized...");
        }


        [CommandMethod("CLI_CommonComandsList")]
        public static void CommonCommandsList(string[] args)
        {
            //Document actDoc = Application.DocumentManager.MdiActiveDocument;
            //Database aCurDB = actDoc.Database;

            actDoc.Editor.WriteMessage("\n1. ACC_ -> Assign Color to Cad objects.");
            actDoc.Editor.WriteMessage("\n2. RTC_ -> Replace Text Content.");
            actDoc.Editor.WriteMessage("\n3. CLC_ -> Create Layer with Color.");
            actDoc.Editor.WriteMessage("\n4. SOR_ -> Select object by reference object properties.");
            actDoc.Editor.WriteMessage("\n5. GPA_ -> Get polyline object.");
            actDoc.Editor.WriteMessage("\nEnd of list.");
        }

      /**
         * Author: Bernardo A. Cabebe Jr.
         * Date: 03 June 2024 11:48am
         * Vanue: 3407 Cityland Pasong Tamo Tower
         */
      [CommandMethod("AIP_AIPoint")]
      public static void AIPoint()
      {
         PromptPointOptions ppoAiP = new PromptPointOptions("");
         ppoAiP.Message = "\nCreate reference point: ";
         PromptPointResult pprAiP = actDoc.Editor.GetPoint(ppoAiP);
         Point3d p3dAiP = pprAiP.Value;

         if (pprAiP.Status != PromptStatus.OK)
         {
            actDoc.Editor.WriteMessage("\nCommand cancelled");
            return;
         }

         using (Transaction trAiP = aCurDB.TransactionManager.StartTransaction())
         {
            BlockTable acBlkTbl;
            acBlkTbl = trAiP.GetObject(aCurDB.BlockTableId, OpenMode.ForRead) as BlockTable;

            BlockTableRecord aBTblRec;
            aBTblRec = trAiP.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

            string layerNamet0y = "Cai_Column_Refpoint";
            SEASTools stools = new SEASTools();
            stools.CreateLayer(layerNamet0y, 111, true);

            using (DBPoint dbpAiP = new DBPoint(p3dAiP))
            {
               dbpAiP.ColorIndex = 111;
               dbpAiP.Layer = layerNamet0y;
               aBTblRec.AppendEntity(dbpAiP);
               trAiP.AddNewlyCreatedDBObject(dbpAiP, true);
            }
            trAiP.Commit();
         }

      }

      [CommandMethod("DMO_DevMessageOn")]
      public static void DevMessageOn()
      {
         PromptStringOptions psoDMO = new PromptStringOptions("");
         psoDMO.Message = "\nEnter passcode: ";
         psoDMO.AllowSpaces = false;
         PromptResult prDMO = actDoc.Editor.GetString(psoDMO);
         if (prDMO.Status != PromptStatus.OK) return;
         string passcode = prDMO.StringResult;
         if (passcode == "KiWi988_874123")
         {
            PromptIntegerOptions pioDMO = new PromptIntegerOptions("");
            int devMsgOn = DataGlobal.isDevMessageOn ? 1 : 0;
            pioDMO.DefaultValue = devMsgOn;
            pioDMO.Message = "\nEnter value: ";
            PromptIntegerResult pirDMO = actDoc.Editor.GetInteger(pioDMO);
            if (pirDMO.Status != PromptStatus.OK) return;
            int iDevMsgOn = pirDMO.Value;
            DataGlobal.isDevMessageOn = iDevMsgOn == 1 ? true : false;
            actDoc.Editor.WriteMessage($"\nDeveloper message is " + (DataGlobal.isDevMessageOn ? "ON" : "OFF"));
            return;
         }
         else
         {
            DataGlobal.isDevMessageOn = false;
            actDoc.Editor.WriteMessage("\nDeveloper message is OFF.");
            return;
         }
      }

      
      /**
       * Author: Bernardo A. Cabebe Jr.
       * Date: 11 Jan 2024 7:55pm
       * Vanue: 3407 Cityland Pasong Tamo Tower
       */
      [CommandMethod("GPA_GetPolylineArea")]
        public static void GetPolylineArea()
        {
            if (!ccData.isLicenseActive())
            {

                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            PromptSelectionOptions psoPLine = new PromptSelectionOptions();
            psoPLine.MessageForAdding = "\nSelect reference object: ";

            PromptSelectionResult psrPLine = actDoc.Editor.GetSelection(psoPLine);
            if (psrPLine.Status != PromptStatus.OK) return;

            SelectionSet ssPLine = psrPLine.Value;

            //if (ssPLine.Count != 1)
            //{
            //    actDoc.Editor.WriteMessage("\nSelect one reference object only.");
            //    return;
            //}

            PromptStringOptions psoSheetAddr = new PromptStringOptions("");
            psoSheetAddr.DefaultValue = shAddrG;
            psoSheetAddr.Message = "\nEnter cell address: ";
            psoSheetAddr.AllowSpaces = false; 
            PromptResult prSheetAddr = actDoc.Editor.GetString(psoSheetAddr);

            string sheetAddr = prSheetAddr.StringResult;
            shAddrG = sheetAddr;
            actDoc.Editor.WriteMessage("\nCell address: {0}", shAddrG);

            PromptStringOptions psoGSheetID = new PromptStringOptions("");
            psoGSheetID.DefaultValue = DataGlobal.gpaSSheetID;
            psoGSheetID.Message = "\nEnter Google Spreadsheet ID: ";
            psoGSheetID.AllowSpaces = false;
            PromptResult prGSheetID = actDoc.Editor.GetString(psoGSheetID);

            string spreadSheetID = prGSheetID.StringResult;
            DataGlobal.gpaSSheetID = spreadSheetID;
            actDoc.Editor.WriteMessage("\nSpreadsheet ID: {0}", DataGlobal.gpaSSheetID);

            double pLineArea = 0;
            using (Transaction trPLine = aCurDB.TransactionManager.StartTransaction())
            {
                int iiObj = 1;
                foreach (SelectedObject soPLine in ssPLine)
                {
                    DBObject dboPLine = trPLine.GetObject(soPLine.ObjectId, OpenMode.ForRead);
                    if (dboPLine is Polyline)
                    {
                        Polyline pLine = dboPLine as Polyline;
                        pLineArea += pLine.Area;
                        actDoc.Editor.WriteMessage("\n{0} Area = {1}", iiObj, pLine.Area.ToString());

                    }
                    else
                    {
                        actDoc.Editor.WriteMessage("\n{0} not polyline", iiObj);
                    }

                    iiObj++;
                }

                actDoc.Editor.WriteMessage("\nTotal area= {0}", pLineArea);

                string[,] arr2Value = new string[1, 2];
                arr2Value[0, 0] = "Area=";
                arr2Value[0, 1] = pLineArea.ToString();

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

                trPLine.Commit();
                trPLine.Dispose();
            }

            return;
        }


        /**
         * Author: Bernardo A. Cabebe Jr.
         * Date: 11 Jan 2024 7:55pm
         * Vanue: 3407 Cityland Pasong Tamo Tower
         */
        [CommandMethod("SOR_SelectObjectByRef")]
        public static void SelectObjectByRef()
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
            else if (objType == "DBTEXT")
            {
                tvFilter[0] = new TypedValue((int)DxfCode.Start, "TEXT");
            }
            else
            {
                tvFilter[0] = new TypedValue((int)DxfCode.Start, objType);
            }

            tvFilter[1] = new TypedValue((int)DxfCode.LayerName, layerName);
            tvFilter[2] = new TypedValue((int)DxfCode.Color, colIndex);

            SelectionFilter sfObjs = new SelectionFilter(tvFilter);
            PromptSelectionResult psrObjs = actDoc.Editor.GetSelection(sfObjs);

            if (psrObjs.Status != PromptStatus.OK) return;

            SelectionSet ssObjs = psrObjs.Value;

            ccData.setNumOfSelection(ssObjs.Count);

            return;

        }


        [CommandMethod("ACC_AssignColorToCADObj")]
        public static void ChangeColor(string[] args)
        {
            //Document actDoc = Application.DocumentManager.MdiActiveDocument;
            //Database aCurDB = actDoc.Database;
            if (!ccData.isLicenseActive())
            {

                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            PromptIntegerOptions colIntPIO = new PromptIntegerOptions("");
            colIntPIO.Message = "\nColor number: ";
            colIntPIO.DefaultValue = colorNoG;
            PromptIntegerResult colIntPRes = actDoc.Editor.GetInteger(colIntPIO);
            int colorNo = colIntPRes.Value;
            colorNoG = colorNo;

            using (Transaction acTrans = aCurDB.TransactionManager.StartTransaction())
            {
                PromptSelectionResult objPSRes = actDoc.Editor.GetSelection();

                if (objPSRes.Status == PromptStatus.OK)
                {
                    SelectionSet objSSet = objPSRes.Value;
                    foreach (SelectedObject iiObj in objSSet)
                    {
                        if (iiObj != null)
                        {
                            Entity acEnt = acTrans.GetObject(iiObj.ObjectId, OpenMode.ForWrite) as Entity;
                            if (acEnt != null)
                            {
                                acEnt.ColorIndex = colorNo;
                            }
                        }
                    }

                    acTrans.Commit();
                    actDoc.Editor.WriteMessage("\nAssigned color to object: " + colorNo);
                }
            }
        }


        [CommandMethod("RTC_ReplaceTextContent")]
        public static void ReplaceTextContent(string[] args)
        {
            //Document actDoc = Application.DocumentManager.MdiActiveDocument;
            //Database aCurDB = actDoc.Database;
            if (!ccData.isLicenseActive())
            {

                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            PromptStringOptions oldTxtPSO = new PromptStringOptions("");
            oldTxtPSO.Message = "\nOld text/pattern ;; New text: ";
            oldTxtPSO.DefaultValue = oldTxtG;
            oldTxtPSO.AllowSpaces = true;
            PromptResult oldTxtRes = actDoc.Editor.GetString(oldTxtPSO);
            actDoc.Editor.WriteMessage("\nOld text/pattern: " + oldTxtRes.StringResult);
            string oldTxt = oldTxtRes.StringResult;
            oldTxtG = oldTxt;
            string[] txtSep = { ";;" };
            string[] arrTxt = oldTxt.Split(txtSep, StringSplitOptions.RemoveEmptyEntries);

            //PromptStringOptions newTxtPSO = new PromptStringOptions("");
            //newTxtPSO.Message = "\nNew text: ";
            //newTxtPSO.DefaultValue = newTxtG;
            //newTxtPSO.AllowSpaces = true;
            //PromptResult newTxtRes = actDoc.Editor.GetString(newTxtPSO);
            //actDoc.Editor.WriteMessage("\nNew text: " + newTxtG);
            //string newTxt = newTxtRes.StringResult;
            //newTxtG = newTxt;

            PromptIntegerOptions colIntPIO = new PromptIntegerOptions("");
            colIntPIO.Message = "\nText color mark: ";
            colIntPIO.DefaultValue = txtColorG;
            PromptIntegerResult colIntPRes = actDoc.Editor.GetInteger(colIntPIO);
            int colorNo = colIntPRes.Value;
            colorNo = colorNo <= 0 ? 1 : colorNo;
            txtColorG = colorNo;

            using (Transaction acTransTxt = aCurDB.TransactionManager.StartTransaction())
            {
                TypedValue[] tvTxtSelFil = new TypedValue[1];
                tvTxtSelFil.SetValue(new TypedValue((int)DxfCode.Start, "*TEXT"), 0);
                SelectionFilter sfTxtSelFil = new SelectionFilter(tvTxtSelFil);

                PromptSelectionResult psrTxt = actDoc.Editor.GetSelection(sfTxtSelFil);
                if (psrTxt.Status != PromptStatus.OK)
                {
                    actDoc.Editor.WriteMessage("\nSelection cancelled.");
                    return;
                }
                SelectionSet ssTxt = psrTxt.Value;
                foreach (SelectedObject soiiTxt in ssTxt)
                {
                    if (soiiTxt == null)
                    {
                        actDoc.Editor.WriteMessage("\nNo selected text objects.");
                        return;
                    }

                    Entity entTxt = acTransTxt.GetObject(soiiTxt.ObjectId, OpenMode.ForWrite) as Entity;
                    if (entTxt is DBText)
                    {
                        DBText dbTxt = entTxt as DBText;
                        string txtCont = dbTxt.TextString;

                        string txtOld = arrTxt[0].Trim(' ');
                        string txtNew = arrTxt[1].Trim(' ');
                        string newTxtCont = txtCont.Replace(txtOld, txtNew);
                        actDoc.Editor.WriteMessage("\nSearch text: {0}; Replacement text: {1}", txtOld, txtNew);
                        actDoc.Editor.WriteMessage("\nOld text: {0}; New text: {1}", txtCont, newTxtCont);

                        dbTxt.TextString = newTxtCont;
                        if (txtCont != newTxtCont)
                        {
                            dbTxt.ColorIndex = colorNo;
                        }

                    }
                    if (entTxt is MText)
                    {
                        MText mTxt = entTxt as MText;
                        string txtCont = mTxt.Contents;

                        string txtOld = arrTxt[0].Trim(' ');
                        string txtNew = arrTxt[1].Trim(' ');
                        string newTxtCont = txtCont.Replace(txtOld, txtNew);
                        actDoc.Editor.WriteMessage("\nSearch text: {0}; Replacement text: {1}", txtOld, txtNew);
                        actDoc.Editor.WriteMessage("\nOld text: {0}; New text: {1}", txtCont, newTxtCont);

                        mTxt.Contents = newTxtCont;
                        if (txtCont != newTxtCont)
                        {
                            mTxt.ColorIndex = colorNo;
                        }

                    }
                }
                acTransTxt.Commit();
                actDoc.Editor.WriteMessage("\nOld text(s) replaced.");
            }
        }


        [CommandMethod("CLC_CreateLayerWithColor")]
        public static void CreateLayer(string[] args)
        {
            //Document actDoc = Application.DocumentManager.MdiActiveDocument;
            //Database aCurDB = actDoc.Database;
            if (!ccData.isLicenseActive())
            {

                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            PromptStringOptions psoLyr = new PromptStringOptions("");
            psoLyr.Message = "\nLayer name: ";
            psoLyr.AllowSpaces = true;
            PromptResult prLyr = actDoc.Editor.GetString(psoLyr);
            actDoc.Editor.WriteMessage("\nLayer name: " + prLyr.StringResult);
            string layerNme = prLyr.StringResult;


            PromptIntegerOptions pioLyr = new PromptIntegerOptions("");
            pioLyr.Message = "\nColor number: ";
            pioLyr.DefaultValue = layerColG;
            PromptIntegerResult pirLyr = actDoc.Editor.GetInteger(pioLyr);
            int layerCol = pirLyr.Value;
            layerCol = layerCol < 0 ? 0 : layerCol;
            layerCol = layerCol > 255 ? 255 : layerCol;
            layerColG = layerCol;

            PromptKeywordOptions pkoIsPlottable = new PromptKeywordOptions("");
            pkoIsPlottable.Message = "\nIs plottable? ";
            pkoIsPlottable.Keywords.Add("Yes");
            pkoIsPlottable.Keywords.Add("No");
            pkoIsPlottable.Keywords.Default = "Yes";
            pkoIsPlottable.AllowNone = true;

            PromptResult prIsPlottable = actDoc.Editor.GetKeywords(pkoIsPlottable);
            string isPlottable = "";
            if (prIsPlottable.Status != PromptStatus.OK) return;
            isPlottable = prIsPlottable.StringResult;
            bool plotJ3w = true;
            switch (isPlottable)
            {
                case "Yes":
                    plotJ3w = true;
                    break;
                case "No":
                    plotJ3w = false;
                    break;
                default:
                    return;
            }

            using (Transaction lyTrans = aCurDB.TransactionManager.StartTransaction())
            {
                LayerTable ltLayer;
                ltLayer = lyTrans.GetObject(aCurDB.LayerTableId, OpenMode.ForRead) as LayerTable;
                if (ltLayer.Has(layerNme) == false)
                {
                    LayerTableRecord ltrLayer = new LayerTableRecord();
                    ltrLayer.Color = acColor.FromColorIndex(ColorMethod.ByAci, (short)layerCol);
                    ltrLayer.Name = layerNme;
                    ltrLayer.IsPlottable = plotJ3w;
                    ltLayer.UpgradeOpen();
                    ltLayer.Add(ltrLayer);
                    lyTrans.AddNewlyCreatedDBObject(ltrLayer, true);

                    actDoc.Editor.WriteMessage($"\nNew layer created: {layerNme}.");
                }

                lyTrans.Commit();

            }

        }


        /**
         * Author: Bernardo A. Cabebe Jr.
         * Date: 03 June 2024 11:48am
         * Vanue: 3407 Cityland Pasong Tamo Tower
         */
        [CommandMethod("STC_SelectTextByContent")]
        public static void SelectTextByContent()
        {
            if (!ccData.isLicenseActive())
            {
                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            PromptStringOptions psoSTCB4r = new PromptStringOptions("");
            psoSTCB4r.Message = "\nEnter text pattern: ";
            psoSTCB4r.UseDefaultValue = true;
            psoSTCB4r.DefaultValue = textPatternG;
            psoSTCB4r.AllowSpaces = false;

            PromptResult prSTCB4r = actDoc.Editor.GetString(psoSTCB4r);
            if (prSTCB4r.Status != PromptStatus.OK) return;
            string textPattern = prSTCB4r.StringResult;
            textPatternG = textPattern;


            TypedValue[] tvFilterG5w = new TypedValue[2];
            tvFilterG5w[0] = new TypedValue((int)DxfCode.Start, "*TEXT");
            tvFilterG5w[1] = new TypedValue((int)DxfCode.Text, textPattern);
            SelectionFilter sfTCGG5w = new SelectionFilter(tvFilterG5w);

            PromptSelectionOptions psoSTCB5r = new PromptSelectionOptions();
            psoSTCB5r.MessageForAdding = "\nSelect text objects: ";

            PromptSelectionResult psrSTCB5r = actDoc.Editor.GetSelection(psoSTCB5r, sfTCGG5w);
            if (psrSTCB5r.Status != PromptStatus.OK) return;

            SelectionSet ssSTCB5r = psrSTCB5r.Value;
            actDoc.Editor.WriteMessage("\nSelected objects: {0}", ssSTCB5r.Count);

            return;

        }

        [CommandMethod("SOR2_SelectObjectByRef2")]
        public static void SelectObjectByRef2()
        {
            if (!ccData.isLicenseActive())
            {
                actDoc.Editor.WriteMessage("\n" + ccData.getLicenseNote());
                return;
            }

            actDoc.Editor.WriteMessage("\nExecuting SOR...\n");
            actDoc.Editor.Command("SOR_SelectObjectByRef");
        }


        [CommandMethod("GFSC_GenerateFootingScheduleInCAD")]
        public static void GenerateFootingScheduleInCAD(string[] args)
        {
            Document actDoc = Application.DocumentManager.MdiActiveDocument;
            Database aCurDB = actDoc.Database;


            try
            {
                actDoc.Editor.WriteMessage($"\nMake a copy of {sheetsUrl}");

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
                psoGSheetID.DefaultValue = ftgSpreadSheetIDG;
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
                ftgSpreadSheetIDG = spreadSheetID;

                IList<IList<object>> iioShtRngData = GetRangeFooting(startRow, endRow, spreadSheetID);
                IList<IList<object>> iioShtRngColW = GetRangeFooting(3, 3, spreadSheetID); //Column widths
                IList<IList<object>> iioShtRngHead = GetRangeFooting(2, 2, spreadSheetID); //Header

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
                    while (ii < iioShtRngData.Count && iiVertDist < verDist)
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

        public static IList<IList<object>> GetRangeFooting(int startRow, int endRow, string spreadSheetID)
        {
            string shtRange = "FootingSchedule!A" + startRow.ToString() + ":L" + endRow.ToString();
            SpreadsheetsResource.ValuesResource.GetRequest grBeamSched = DataGlobal.sheetsService.Spreadsheets.Values.Get(spreadSheetID, shtRange);
            ValueRange vrBeamSched = grBeamSched.Execute();
            IList<IList<object>> iioBeamSched = vrBeamSched.Values;

            return iioBeamSched;
        }




        [CommandMethod("TC_TestCommand")]
        public static void TestCommand()
        {
            try
            {
                /**
                *string filePath = DataGlobal.credLicPath + "\\Resources\\BCECSiSlabDesignCNegM.tsv";
                *
                *StreamReader srFile = new StreamReader(filePath);
                *
                *List<object> loLines = new List<object>();
                *string strLine = srFile.ReadLine();
                *while (strLine != null)
                *{
                *    strLine = srFile.ReadLine();
                *    loLines.Add(strLine);
                *}
                *
                *srFile.Close();
                *
                *int aa = 0;
                *foreach (object objL in loLines)
                *{
                *    actDoc.Editor.WriteMessage("\nLine {0}: {1}", aa, objL);
                *    aa++;
                *}
                */

                IList<IList<object> > iioConvoData = DataGlobal.trainingConversation;

                for (int outerIndex = 0; outerIndex < iioConvoData.Count; outerIndex++)
                {
                    for (int innerIndex = 0; innerIndex < iioConvoData[outerIndex].Count; innerIndex++)
                    {
                        object item = iioConvoData[outerIndex][innerIndex];
                        actDoc.Editor.WriteMessage($"\n[{outerIndex}][{innerIndex}]:{item}");
                    }
                    actDoc.Editor.WriteMessage("\n-----------");
                }



            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                actDoc.Editor.WriteMessage("\nError T2: {0}", e.Message);
            }
        }
    }
}
