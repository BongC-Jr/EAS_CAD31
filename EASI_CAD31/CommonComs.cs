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
   public class CommonComs
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

            actDoc.Editor.WriteMessage("\nSEAS commands data initialized...");
        }

      [CommandMethod("ACC_AssignColorToCADObj")]
        public static void ChangeColor(string[] args)
        {
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
   }
}
