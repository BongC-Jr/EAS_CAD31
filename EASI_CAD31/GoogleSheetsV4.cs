using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using static Google.Apis.Sheets.v4.SheetsService;

namespace EASI_CAD31
{
    public class GoogleSheetsV4
    {
        static Document actDoc = Application.DocumentManager.MdiActiveDocument;
        static Database aCurDB = actDoc.Database;

        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "GoogleSheetsDotNetCadOAuth";

        /**
         * Date added: 11 Jan. 2024
         * Addede by: Engr Bernardo Cabebe Jr.
         * Venue: 3407 Cityland Pasong Tamo Tower
         */

        public void GoogleSheetInit()
        {
            if(DataGlobal.credentialDirectory == null)
            {
                DataGlobal.credentialDirectory = "C:\\SEAS";
            }
            
            UserCredential uCredential;

            //Get initial source directory
            actDoc.Editor.WriteMessage("\nThe SEAS folder files can be placed or located in your custom folder" +
                                       "\nlike C:\\<myFolder>. As specific example is D:\\KiWi, where inside the folder KiWi" +
                                       "\nare the SEAS files. As an important note, the C:\\<myFolder> or D:\\KiWi folder must" +
                                       "\ncontain the files BCECSiGoogleSheetCredV21.json and SEASLicV21.txt as minimum.");
            PromptStringOptions psoSDr = new PromptStringOptions("");
            psoSDr.Message = "\nEnter SEAS folder: ";
            psoSDr.DefaultValue = DataGlobal.credentialDirectory;
            psoSDr.AllowSpaces = false;
            PromptResult prSDr = actDoc.Editor.GetString(psoSDr);
            if (prSDr.Status != PromptStatus.OK) return;
            string strSDr = prSDr.StringResult;
            actDoc.Editor.WriteMessage("\nSearch path: {0}",strSDr);
            DataGlobal.credentialDirectory = strSDr;
            if(strSDr.Length <= 3)
            {
                actDoc.Editor.WriteMessage("\nInvalid SEAS folder.");
                return;
            }

            string credFile = "BCECSiGoogleSheetCredV21.json";
            string searchPath = strSDr;
            SEASTools seasTools = new SEASTools();
            string pathResult = seasTools.FilePath(credFile, searchPath);
            if(pathResult == null)
            {
                actDoc.Editor.WriteMessage("\nInvalid SEAS folder.");
                return;
            }
            string credFilePath = searchPath + "\\" + pathResult;
            string credDirectory = Path.GetDirectoryName(credFilePath);
            DataGlobal.credLicPath = credDirectory;
            actDoc.Editor.WriteMessage("Credential file path: {0} and the directory is {1}", credFilePath, credDirectory);

            using (var stream = new FileStream(credFilePath, FileMode.Open, FileAccess.Read))
            {
                /* The file token.json stores the user's access and refresh tokens, and is created
                 automatically when the authorization flow completes for the first time. */
                string credPath = credDirectory;//"BCECSiGoogleSheetCred2.json";
                uCredential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                //actDoc.Editor.WriteMessage("\nCredential file saved to: " + credPath);
                //Console.WriteLine("Credential file saved to: " + credPath);
            }
        
            // Create Google Sheets API service.
            DataGlobal.sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = uCredential,
                ApplicationName = ApplicationName
            });
        }

        public IList<IList<object>> GetRange(string rangeAddress, string spreadSheetID)
        {
            string strRngVal = rangeAddress;
            SpreadsheetsResource.ValuesResource.GetRequest grRangVal = DataGlobal.sheetsService.Spreadsheets.Values.Get(spreadSheetID,strRngVal);
            ValueRange vrRngVal = grRangVal.Execute();
            IList<IList<object>> iioRngVal = vrRngVal.Values;

            return iioRngVal;
        }

        
    }
}
