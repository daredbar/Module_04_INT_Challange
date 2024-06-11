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
using Forms = System.Windows.Forms;

#endregion

namespace Module_04_INT_Challange
{
    [Transaction(TransactionMode.Manual)]
    public class Mod_04_INT_B : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Prompt user to select RVT file
            string revitFile = "";

            Forms.OpenFileDialog ofd = new Forms.OpenFileDialog();
            ofd.Title = "Select Revit file";
            ofd.InitialDirectory = @"C:\";
            ofd.Filter = "Revit files (*.rvt)|*.rvt";
            ofd.RestoreDirectory = true;

            if (ofd.ShowDialog() != Forms.DialogResult.OK)
                return Result.Failed;
            
            revitFile = ofd.FileName;

            // Open selected file in background
            UIDocument closedUIDoc = uiapp.OpenAndActivateDocument(revitFile);
            Document closedDoc = closedUIDoc.Document;

            FilteredElementCollector modelGroupCollector = new FilteredElementCollector(closedDoc)
                .OfCategory(BuiltInCategory.OST_IOSModelGroups);

            TaskDialog.Show("Copy Model Groups", $"There are {modelGroupCollector.Count()} groups in the model I just opened");

            // Make other document active then close document
            uiapp.OpenAndActivateDocument(doc.PathName);
            closedDoc.Close(false);

            // inserting groups
            // get group type
            string groupName1 = "Group 1";
            string groupName2 = "Group 2";
            string groupName3 = "Group 3";

            GroupType curGroup1 = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_IOSModelGroups)
                .WhereElementIsElementType()
                .Where(r => r.Name == groupName1)
                .Cast<GroupType>().First();

            GroupType curGroup2 = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_IOSModelGroups)
                .WhereElementIsElementType()
                .Where(r => r.Name == groupName2)
                .Cast<GroupType>().First();

            GroupType curGroup3 = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_IOSModelGroups)
                .WhereElementIsElementType()
                .Where(r => r.Name == groupName3)
                .Cast<GroupType>().First();

            // insert group
            XYZ insPoint = new XYZ();

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Place Group");

                doc.Create.PlaceGroup(insPoint, curGroup);

                t.Commit();
            }


            return Result.Succeeded;
        }

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}
