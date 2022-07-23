using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using Forms = System.Windows.Forms;

namespace RevitAddinAcademy
{
    internal class Utility
    {
        public static View GetViewByName(Document doc, string viewName)
        {
            FilteredElementCollector colViews = new FilteredElementCollector(doc);
            colViews.OfCategory(BuiltInCategory.OST_Views);

            foreach (View curView in colViews)
            {
                if (curView.Name == viewName)
                    return curView;
            }

            return null;
        }

        public static void SetParameterValue(ViewSheet newSheet, string paramName, string paramValue)
        {
            foreach (Parameter curParam in newSheet.Parameters)
            {
                if (curParam.Definition.Name == paramName)
                {
                    curParam.Set(paramValue);
                }
            }
        }
    }
}
