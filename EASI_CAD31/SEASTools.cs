using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Colors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

using acColor = Autodesk.AutoCAD.Colors.Color;
using System.IO;
using System.Data.Odbc;

namespace EASI_CAD31
{
   public class SEASTools
   {
      static Document actDoc = Application.DocumentManager.MdiActiveDocument;
      static Database aCurDB = actDoc.Database;

      public void CreateLayer(string layerName, short colorIndex, bool isPlottable = true)
      {
         using (Transaction trLayerK1w = aCurDB.TransactionManager.StartTransaction())
         {
            LayerTable ltLayerK1w = (LayerTable)trLayerK1w.GetObject(aCurDB.LayerTableId, OpenMode.ForRead);
            if (ltLayerK1w.Has(layerName) == false)
            {
               using (LayerTableRecord ltrLayerT1w = new LayerTableRecord())
               {
                  //Assign the layer the necessary properties
                  ltrLayerT1w.Name = layerName;
                  ltrLayerT1w.Color = acColor.FromColorIndex(ColorMethod.ByAci, colorIndex);
                  ltrLayerT1w.IsPlottable = isPlottable;

                  //Upgrade the layer table for write
                  trLayerK1w.GetObject(aCurDB.LayerTableId, OpenMode.ForWrite);

                  //Append the new layer to the layer talbe and the transaction
                  ltLayerK1w.Add(ltrLayerT1w);
                  trLayerK1w.AddNewlyCreatedDBObject(ltrLayerT1w, true);
               }
            }

            trLayerK1w.Commit();
            trLayerK1w.Dispose();
         }

         return;
      }

      public string FilePath(string searchFile, string searchDirectory)
      {
         try
         {
            var txtFiles = Directory.EnumerateFiles(searchDirectory, searchFile, SearchOption.AllDirectories);
            foreach (string currentFile in txtFiles)
            {
               string fileName = currentFile.Substring(searchDirectory.Length + 1);
               if (fileName.Contains(searchFile))
               {
                  return fileName;
               }
            }
         }
         catch (Exception e)
         {
            return e.Message;
         }

         return null;
      }

      public int?  LogConversation(string conversationContent)
      {
         string filePath = Path.Combine(DataGlobal.convofilepath, "conversation.txt");
         if (Directory.Exists(DataGlobal.convofilepath))
         {
            string content = conversationContent;
            File.AppendAllText(filePath, content + Environment.NewLine);
            return 1; // Success
         }
         else
         {
            return -1; // Directory does not exist
         }
      }
        
    }
}
