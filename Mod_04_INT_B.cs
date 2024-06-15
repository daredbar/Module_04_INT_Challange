#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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

            List<ElementId> elemIdList = modelGroupCollector.Select(elem => elem.Id).ToList();
            // copy element
            CopyPasteOptions options = new CopyPasteOptions();

            TaskDialog.Show("Copy Model Groups", $"There are {modelGroupCollector.Count()} groups in the model I just opened");

            int gA = 0;
            int gB = 0;
            int gC = 0;

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Place Group");

                // inserting groups
                ElementTransformUtils.CopyElements(closedDoc, elemIdList, doc, null, options);

                // get group type
                GroupType curGroupA = GetGroupTypeByName(doc, "GroupA");
                GroupType curGroupB = GetGroupTypeByName(doc, "GroupB");
                GroupType curGroupC = GetGroupTypeByName(doc, "GroupC");


                //Loop through Spaces
                FilteredElementCollector curSpace = new FilteredElementCollector(doc);
                curSpace.OfCategory(BuiltInCategory.OST_MEPSpaces);

                foreach (Space space in curSpace)
                {
                    // get room data
                    string spaceComm = space.LookupParameter("Comments").AsString();
                    // get room location point
                    LocationPoint spacePoint = space.Location as LocationPoint;
                    // insert group
                    XYZ insPoint = spacePoint.Point;

                    switch (spaceComm)
                    {
                        case "GroupA":
                            doc.Create.PlaceGroup(insPoint, curGroupA);
                            gA++;
                            break;
                        case "GroupB":
                            doc.Create.PlaceGroup(insPoint, curGroupB);
                            gB++;
                            break;
                        case "GroupC":
                            doc.Create.PlaceGroup(insPoint, curGroupC);
                            gC++;
                            break;
                    }
                }

                t.Commit();
            }

            // Make other document active then close document
            uiapp.OpenAndActivateDocument(doc.PathName);
            closedDoc.Close(false);

            TaskDialog.Show("Model Groups", $"{gA} Group A added, {gB} Group B added, {gC} Group C added.");

            return Result.Succeeded;
        }

        public static GroupType GetGroupTypeByName(Document doc, string groupName)
        {
            return new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_IOSModelGroups) 
                .WhereElementIsElementType().Where(r => r.Name == groupName)
                .Cast<GroupType>().First();
        }

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}
