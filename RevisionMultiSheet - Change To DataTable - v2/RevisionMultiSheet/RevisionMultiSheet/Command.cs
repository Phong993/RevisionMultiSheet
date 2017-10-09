#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Windows;

#endregion

namespace RevisionMultiSheet
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
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;

            // Access current selection

            //Selection sel = uidoc.Selection;            
            Window window = new Window();
            window.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            window.Height = 785; window.Width = 890;
            window.Title = "Revision To MultiSheet";
            window.ResizeMode = ResizeMode.NoResize;
            FormMaster formMaster = new FormMaster();

            formMaster.App = app;
            formMaster.Doc = doc;
            formMaster.UIDoc = uidoc;
            window.Content = formMaster;
            window.ShowDialog();

            return Result.Succeeded;
        }
    }
}
