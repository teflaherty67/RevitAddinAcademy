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
using Forms = System.Windows.Forms;

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

            Forms.OpenFileDialog dialog = new Forms.OpenFileDialog();
            dialog.InitialDirectory = @"S:\Shared Folders\Lifestyle USA Design\04-Customization";
            dialog.Multiselect = false;
            dialog.Filter = "Excel Files | *.xls; *.xlsx; | All files | *.*";

            if (dialog.ShowDialog() != Forms.DialogResult.OK)
            {
                TaskDialog.Show("Error", "Please select an Excel file");
                return Result.Failed;
            }

            string excelFile = dialog.FileName;
            int levelCounter = 0;
            int sheetCounter = 0;

            try
            {
                // open excel
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWb = excelApp.Workbooks.Open(excelFile);
              

                Excel.Worksheet excelWs1 = GetExcelWorksheetByName(excelWb, "Levels");
                Excel.Worksheet excelWs2 = GetExcelWorksheetByName(excelWb, "Sheets");

                List<LevelStruct> levelData = GetLevelDataFromExcel(excelWs1);
                List<SheetStruct> sheetData = GetSheetDataFromExcel(excelWs2);

                Excel.Range excelRng1 = excelWs1.UsedRange;
                Excel.Range excelRng2 = excelWs2.UsedRange;

                int rowCount1 = excelRng1.Rows.Count;
                int rowCount2 = excelRng2.Rows.Count;

                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Setup project");

                    for (int i = 2; i <= rowCount1; i++)
                    {
                        Excel.Range levelData1 = excelWs1.Cells[i, 1];
                        Excel.Range levelData2 = excelWs1.Cells[i, 2];

                        string levelName = levelData1.Value.ToString();
                        double levelElev = levelData2.Value;

                        try
                        {
                            Level newLevel = Level.Create(doc, levelElev);
                            newLevel.Name = levelName;
                            levelCounter++;
                        }
                        catch (Exception ex)
                        {
                            Debug.Print(ex.Message);
                        }

                    }

                    FilteredElementCollector colSheet = new FilteredElementCollector(doc);
                    colSheet.OfCategory(BuiltInCategory.OST_TitleBlocks);
                    colSheet.WhereElementIsElementType();

                    for (int j = 2; j <= rowCount2; j++)
                    {
                        Excel.Range sheetData1 = excelWs2.Cells[j, 1];
                        Excel.Range sheetData2 = excelWs2.Cells[j, 2];

                        string sheetNum = sheetData1.Value.ToString();
                        string sheetName = sheetData2.Value.ToString();

                        try
                        {
                            ViewSheet newSheet = ViewSheet.Create(doc, colSheet.FirstElementId());
                            newSheet.SheetNumber = sheetNum;
                            newSheet.Name = sheetName;
                        }
                        catch (Exception ex)
                        {
                            Debug.Print(ex.Message);
                        }

                    }

                    t.Commit();
                }

                excelWb.Close();
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                Debug.Print(ex.Message);
            }

            TaskDialog.Show("Complete", "Created " + levelCounter.ToString() + " levels.");

            return Result.Succeeded;
        }

        private List<LevelStruct> GetLevelDataFromExcel(Excel.Worksheet excelWs1)
        {
            List<LevelStruct> returnList = new List<LevelStruct>();
            Excel.Range excelRng1 = excelWs1.UsedRange;
           
            int rowCount1 = excelRng1.Rows.Count;

            for (int i = 2; i <= rowCount1; i++)
            {
                Excel.Range levelData1 = excelWs1.Cells[i, 1];
                Excel.Range levelData2 = excelWs1.Cells[i, 2];

                string levelName = levelData1.Value.ToString();
                double levelElev = levelData2.Value;

                LevelStruct curlevel = new LevelStruct(levelName, levelElev);
                returnList.Add(curlevel);
            }

            return returnList;
        }

        private List<SheetStruct> GetSheetDataFromExcel(Excel.Worksheet excelWs2)
        {
            
        }

        private Excel.Worksheet GetExcelWorksheetByName(Excel.Workbook curWb, string wsName)
        {
            foreach(Excel.Worksheet sheet in curWb.Worksheets)
            {
                if(sheet.Name == wsName)
                {  
                    return sheet;
                }
            }

            return null;
        }
        private struct LevelStruct
        {
            public string LevelName;
            public double LevelElev;

            public LevelStruct(string name, double elev)
            {
                LevelName = name;
                LevelElev = elev;
            }
        }

        private struct SheetStruct
        {
            public string SheetNumber;
            public string SheetName;
            public string SheetView;
            public string DrawnBy;
            public string CheckedBy;

            public SheetStruct(string number, string name, string view, string db, string cb)
            {
                SheetNumber = number;
                SheetName = name;
                SheetView = view;
                DrawnBy = db;
                CheckedBy = cb;
            }
        }
    }
}
