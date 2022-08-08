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
    public class cmdWallsFromLines : IExternalCommand
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

            List<string> lineStyles = GetAllLineStyleNames(doc);
            List<string> wallTypes = GetAllWallTypeNames(doc);

            frmWallsFromLines curForm = new frmWallsFromLines(lineStyles, wallTypes);
            curForm.Height = 450;
            curForm.Width = 550;
            curForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;

            if(curForm.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string selectedLineStyle = "";
                string selectedWallType = "";
                double wallHeight = 0;
                bool isStructural = false;
            }
            
            return Result.Succeeded;
        }

        private List<string> GetAllLineStyleNames(Document doc)
        {
            List<string> results = new List<string>();

            FilteredElementCollector collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
            collector.OfClass(typeof(CurveElement));

            foreach(CurveElement element in collector)
            {
                GraphicsStyle curGS = element.LineStyle as GraphicsStyle;

                if(results.Contains(curGS.Name) == false)
                {
                    results.Add(curGS.Name);
                }
            }

            return results;
         }

        private List<string> GetAllWallTypeNames(Document doc)
        {
            List<string> results = new List<string>();

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(WallType));

            foreach (WallType wallType in collector)
            {
               results.Add(wallType.Name);
            }

            return results;
        }       
    }
}
