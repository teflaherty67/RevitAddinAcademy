#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

#endregion

namespace RevitAddinAcademy
{
    [Transaction(TransactionMode.Manual)]
    public class cmdDeleteUnused : IExternalCommand
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

            // create list of views to delete
            List<View> viewsToDelete = new List<View>();

            // create list of views to keep
            List<View> viewsToKeep = new List<View>();

            // get all the views in the project
            FilteredElementCollector colViews = new FilteredElementCollector(doc);
            colViews.OfCategory(BuiltInCategory.OST_Views);

            // get all the sheets in the project
            FilteredElementCollector colSheets = new FilteredElementCollector(doc);
            colSheets.OfClass(typeof(ViewSheet));

            // make sure there are sheets in the project
            if(colSheets.GetElementCount() < 1)
            {
                // alert user
                TaskDialog.Show("Error", "There are no sheets in the project. Please add some");
            }

            // loop through views
            foreach(View curView in colViews)
            {
                // check if view name has a prefix
                if(curView.Name.Contains("working_") == false)
                {
                    // check if view is already on a sheet
                    if(Viewport.CanAddViewToSheet(doc, colSheets.FirstElementId(), curView.Id))
                    {
                        // check if view has dependent views
                        if(curView.GetDependentViewIds().Count == 0)
                        {
                            // add views to lis of views to delete
                            viewsToDelete.Add(curView);
                        }
                    }
                }
            }

            // start a transaction
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Delete Unused Views");

                // loop through list of views to delete
                try
                {
                    foreach(View deleteView in viewsToDelete)
                    {
                        // delete the views
                        doc.Delete(deleteView.Id);
                    }
                }
                catch (Exception)
                {
                    TaskDialog.Show("Error", "Could not delete view");
                }

                // close the transaction
                t.Commit();
            }

            // alert the user
            TaskDialog.Show("Complete", "Deleted " + viewsToDelete.Count.ToString() + " views.");

            return Result.Succeeded;
        }
    }
}
