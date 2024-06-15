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

#endregion

namespace Module_04_INT_Challange
{
    [Transaction(TransactionMode.Manual)]
    public class Mod_04_INT_C : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Get open file with specific name
            // create document varible
            Document openDoc = null;

            // Loop through open documents and look for match
            foreach (Document curDoc in uiapp.Application.Documents)
            {
                if (curDoc.PathName.Contains("Sample 03"))
                    openDoc = curDoc;
            }

            // copy elements from other doc
            // create filtered element collector to get elements
            FilteredElementCollector collector = new FilteredElementCollector(openDoc);

            List<BuiltInCategory> catList = new List<BuiltInCategory>();
            catList.Add(BuiltInCategory.OST_Walls);
            catList.Add(BuiltInCategory.OST_GenericModel);

            ElementMulticategoryFilter catFilter = new ElementMulticategoryFilter(catList);
            collector.WherePasses(catFilter).WhereElementIsNotElementType();

            // get list of element Ids
            List<ElementId> elemIdList = collector.Select(elem =>  elem.Id).ToList();
            int elementCount = elemIdList.Count;

            // copy element
            CopyPasteOptions options = new CopyPasteOptions();

            using(Transaction t = new Transaction(doc))
            {
                t.Start("Copy Elements");

                ElementTransformUtils.CopyElements(openDoc, elemIdList, doc, null, options);

                t.Commit();
            }

            // Display the number of copied elements
            TaskDialog.Show("Copy Elements", $"Number of elemets copied: {elementCount}");

            return Result.Succeeded;
        }

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}
