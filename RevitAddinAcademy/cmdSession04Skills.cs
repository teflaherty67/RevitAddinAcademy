#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Forms = System.Windows.Forms;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB.Architecture;

#endregion

namespace RevitAddinAcademy
{
    [Transaction(TransactionMode.Manual)]
    public class cmdSession04Skills : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            IList<Element> pickList = uidoc.Selection.PickElementsByRectangle("Select elements");
            List<CurveElement> curveList = new List<CurveElement>();

            WallType curWallType = GetWallTypeByName(doc, @"Generic - 8""");
            Level curLevel = GetLevelByName(doc, "Level 1");

            MEPSystemType curSystemType = GetSystemTypeByName(doc, "Domestic Hot Water");
            PipeType curPipeType = GetPipeTypeByName(doc, "Default");

            using(Transaction t = new Transaction(doc))
            {
                t.Start("Create Revit Elements");

                foreach (Element element in pickList)
                {
                    if (element is CurveElement)
                    {
                        CurveElement curve = (CurveElement)element;
                        CurveElement curve2 = element as CurveElement;

                        curveList.Add(curve);

                        GraphicsStyle curGS = curve.LineStyle as GraphicsStyle;

                        Curve curCurve = curve.GeometryCurve;
                        XYZ startPoint = curCurve.GetEndPoint(0);
                        XYZ endPoint = curCurve.GetEndPoint(1);

                        // Wall newWall = Wall.Create(doc, curCurve, curWallType.Id, curLevel.Id, 15, 0, false, false);
                        Pipe newPipe = Pipe.Create(doc, curSystemType.Id, curPipeType.Id, curLevel.Id, startPoint, endPoint);

                        Debug.Print(curGS.Name);
                    }
                }

                t.Commit();
            }         

            TaskDialog.Show("Complete", curveList.Count.ToString());

            return Result.Succeeded;
        }

        private WallType GetWallTypeByName(Document doc, string wallTypeName)
        {
            FilteredElementCollector colWTypes = new FilteredElementCollector(doc);
            colWTypes.OfClass(typeof(WallType));

            foreach(Element curElem in colWTypes)
            {
                WallType wallType = curElem as WallType;

                if(wallType.Name == wallTypeName)
                    return wallType;
            }
            return null;
        }

        private Level GetLevelByName(Document doc, string levelName)
        {
            FilteredElementCollector colLevels = new FilteredElementCollector(doc);
            colLevels.OfClass(typeof(Level));

            foreach (Element curElem in colLevels)
            {
                Level level = curElem as Level;

                if (level.Name == levelName)
                    return level;
            }
            return null;
        }

        private MEPSystemType GetSystemTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector colMEPSysType = new FilteredElementCollector(doc);
            colMEPSysType.OfClass(typeof(MEPSystemType));

            foreach (Element curElem in colMEPSysType)
            {
                MEPSystemType curType = curElem as MEPSystemType;

                if (curType.Name == typeName)
                    return curType;
            }
            return null;
        }
        private PipeType GetPipeTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector colPipeType = new FilteredElementCollector(doc);
            colPipeType.OfClass(typeof(MEPSystemType));

            foreach (Element curElem in colPipeType)
            {
               PipeType curType = curElem as PipeType;

                if (curType.Name == typeName)
                    return curType;
            }
            return null;
        }
    }
}
