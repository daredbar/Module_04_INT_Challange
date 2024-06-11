#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

#endregion

namespace Module_04_INT_Challange
{
    [Transaction(TransactionMode.Manual)]
    public class Mod_04_INT_A : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Get all Links
            FilteredElementCollector linkCollector = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkType));

            // Loop Through links and get doc if loaded
            Document linkDoc       = null;
            RevitLinkInstance link = null;

            foreach (RevitLinkType rvtLink in linkCollector)
            {
                if (rvtLink.GetLinkedFileStatus() == LinkedFileStatus.Loaded)
                {
                    link = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_RvtLinks)
                        .OfClass(typeof(RevitLinkInstance))
                        .Where(x => x.GetTypeId() == rvtLink.Id).First() as RevitLinkInstance;

                    linkDoc = link.GetLinkDocument();
                }   
            }

            FilteredElementCollector collectorA = new FilteredElementCollector(linkDoc)
                .OfCategory(BuiltInCategory.OST_Rooms);

            //TaskDialog.Show("Test Method A", $"There are {collectorA.Count()} rooms in the Link");

            // create space from Room
            // use LINQ to get room
            //Room curRoom = new FilteredElementCollector(doc)
            //    .OfCategory(BuiltInCategory.OST_Rooms)
            //    .Cast<Room>()
            //    .First();

            FilteredElementCollector curRoom = new FilteredElementCollector(linkDoc);
            curRoom.OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType();

            int roomCount = 0;

            foreach (Room room in curRoom)
            {
                // Get level from current view
                Level curLevel = doc.ActiveView.GenLevel;

                // get room data
                string roomName = room.Name;
                string roomNum = room.Number;
                string roomComm = room.LookupParameter("Comments").AsString();

                // get room location point
                LocationPoint roomPoint =room.Location as LocationPoint;

                using(Transaction t = new Transaction(doc))
                {
                    t.Start("Create Space");

                    // create space and transfer properties
                    SpatialElement newSpace = doc.Create.NewSpace
                        (curLevel, new UV(roomPoint.Point.X, roomPoint.Point.Y));
                    newSpace.Name = roomName;
                    newSpace.Number = roomNum;
                    newSpace.LookupParameter("Comments").Set(roomComm);

                    roomCount++;

                    t.Commit();
                }
            }

            TaskDialog.Show("Rooms Created", $"There are {roomCount} rooms created!");

            return Result.Succeeded;
        }

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}
