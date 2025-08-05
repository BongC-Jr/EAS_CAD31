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
         * Date added: 19 Jul 2025
         * Addede by: Engr Bernardo Cabebe Jr.
         * Venue: Subway Eastwood, Bagumbayan, Quezon City
         */
        [CommandMethod ("AI_DrawLineM01")]
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
            string content = $"Assistant: Drawn a line from {startPt} to {endPt} with color index of {colorIndex}. The length of the line is {lineLength}\n";
            File.AppendAllText(filePath, content + Environment.NewLine);
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

            PromptIntegerResult pirPtColor = actDoc.Editor.GetInteger(pioPtColor);
            if (pirPtColor.Status != PromptStatus.OK)
            {
                actDoc.Editor.WriteMessage("\nCommand cancelled.");
                return;
            }
            int iPtColor = pirPtColor.Value;

            TypedValue[] tvFilterG5w = new TypedValue[2];
            tvFilterG5w[0] = new TypedValue((int)DxfCode.Start, "POINT");
            tvFilterG5w[1] = new TypedValue((int)DxfCode.Color, iPtColor);

            PromptSelectionResult psrPtObj = actDoc.Editor.SelectAll(new SelectionFilter(tvFilterG5w));
            if (psrPtObj.Status != PromptStatus.OK)
            {
                actDoc.Editor.WriteMessage("\nNo point objects found with the specified color.");
                return;
            }

            using(Transaction trPtObj = aCurDB.TransactionManager.StartTransaction())
            {
                SelectionSet ssPtObj = psrPtObj.Value;
                if (ssPtObj.Count == 0)
                {
                    actDoc.Editor.WriteMessage("\nNo point objects found with the specified color.");
                    return;
                }
                // Get the first selected point object
                ObjectId ptObjId = ssPtObj[0].ObjectId;
                DBPoint ptObj = (DBPoint)trPtObj.GetObject(ptObjId, OpenMode.ForRead);
                // Store the point's position
                _lastSelectedPoint = ptObj.Position;
                string filePath = Path.Combine(DataGlobal.convofilepath, "conversation.txt");
                string content = $"Assistant: Object is Point, Color index is {iPtColor}, Point coordinate is {_lastSelectedPoint}, and Point Id is {ptObjId}\n";
                // Create or append to the file
                File.AppendAllText(filePath, content + Environment.NewLine);

                // Optionally, you can display the point's position in the command line
                actDoc.Editor.WriteMessage($"\nSelected point: {_lastSelectedPoint}");
                trPtObj.Commit();
            }
        }
    }
}
