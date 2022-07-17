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
    public class cmdProjectSetup : IExternalCommand
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

            string excelFile = @"C:\temp\Session 02_Challenge.xlsx";
            
            // open Excel

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile);

            Excel.Worksheet excelWs1 = excelWb.Worksheets.Item[1];
            Excel.Worksheet excelWs2 = excelWb.Worksheets.Item[2];

            Excel.Range excelRng1 = excelWs1.UsedRange;
            Excel.Range excelRng2 = excelWs2.UsedRange;

            int rowCount1 = excelRng1.Rows.Count;
            int rowCount2 = excelRng2.Rows.Count;

            for(int i = 2; i <=rowCount1; i++)
            {

            }

            for(int j = 2; j <=rowCount2; j++)
            {

            }



            excelWb.Close();
            excelApp.Quit();
           

            return Result.Succeeded;
        }
    }
}
