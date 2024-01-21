#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using APP = System.Windows;

#endregion

namespace sHEETS_WPF
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
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

            // put any code needed for the form here

            // open form

            List<Element> sheetcollector = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType().ToElements().Cast<Element>().ToList();


            var currentForm = new MyForm(sheetcollector, GetViews(doc))
            {
                Width = 800,
                Height = 450,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                Topmost = false,
            };

            currentForm.ShowDialog();

            Transaction t = new Transaction(doc);

            t.Start("sheet");
            {
                if (currentForm.DialogResult == true)
                {
                    List<Dataclass> dataclasses = currentForm.GetData();
                    foreach (Dataclass curData in currentForm.GetData())
                    {
                        ViewSheet newsheet = null;
                        if (curData.Column3 == true) { newsheet = ViewSheet.CreatePlaceholder(doc); }
                        else
                        {
                            newsheet = ViewSheet.Create(doc, curData.Column4.Id);
                        }
                        newsheet.Name = curData.Column1.ToString();
                        newsheet.SheetNumber = curData.Column2.ToString();

                        if(curData.Column5 != null && curData.Column3 == false)
                        {
                            Viewport viewport = Viewport.Create(doc, newsheet.Id, curData.Column5.Id, new XYZ());
                        }
                    }
                }
                //ViewSheet.CreatePlaceholder(doc);
            }
            t.Commit();
            t.Dispose();



            // get form data and do something

            return Result.Succeeded;
        }



        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }

        private List<View> GetViews(Document doc)
        {
            FilteredElementCollector viewList = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views);
            FilteredElementCollector sheetcollector = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets);
            List<View> allview = new List<View>();
            try
            {
                foreach (View curView in viewList)
                {
                    if (curView.IsTemplate == false)
                    {
                        if (Viewport.CanAddViewToSheet(doc, sheetcollector.FirstElementId(), curView.Id) == true)
                        {
                            allview.Add(curView);
                        }
                    }
                }
            }

            catch { }

            return allview;
        }



    }


}
