using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using System.Threading.Tasks;

namespace EASI_CAD31
{
   public class ZzTestSite
   {
      
      [CommandMethod("ZTS_CheckLayer")]
      public void ZTS_CheckLayer()
      {
         Document iAcDoc = Application.DocumentManager.MdiActiveDocument; //Active document
         Database iCurDB = iAcDoc.Database;
         
         PromptStringOptions psoCEst = new PromptStringOptions("\nEstimate params: ");
         psoCEst.AllowSpaces = false;
         PromptResult prCEst = iAcDoc.Editor.GetString(psoCEst);
         if (prCEst.Status != PromptStatus.OK)
         {
            iAcDoc.Editor.WriteMessage("\nCommand cancelled.");
            return;
         }
         string layerName = prCEst.StringResult;

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

         if(layerExists)
         {
            iAcDoc.Editor.WriteMessage($"\nLayer {layerName} exists.");
         }
         else
         {
            iAcDoc.Editor.WriteMessage($"\nLayer {layerName} does not exist.");
         }
      }
   }
}