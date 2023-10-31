#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Windows;
using Application = Autodesk.Revit.ApplicationServices.Application;
using System.IO;
using System.Runtime.InteropServices;

#endregion

namespace TextAnno
{
    [Transaction(TransactionMode.ReadOnly)]
    public class RetrieveText : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            List<string> testStrings = new FilteredElementCollector(doc)
            .OfClass(typeof(TextNote))
            .OfType<TextNote>()
            .Cast<TextNote>()
            .Where(tn => tn.OwnerViewId != ElementId.InvalidElementId) // Check if it's associated with a view (sheet)
            .Select<TextNote, string>(tn => tn.Text)
            .ToList<string>();

            List<string> distinctTextStrings = testStrings.Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            string concatenatedText = string.Join("\n", distinctTextStrings);

            // Create and show a TaskDialog
            TaskDialog taskDialog = new TaskDialog("Text Notes in Revit");
            taskDialog.MainContent = concatenatedText;
            taskDialog.Show();

            // Define the path and filename for the Notepad file
            string filePath = "C:\\ProgramData\\Autodesk\\Revit\\TextNotes.txt";

         

            // Create a StreamWriter to write to the file
            using (StreamWriter sw = new StreamWriter(filePath))
            {
                // Write each text note to the file
                foreach (string text in distinctTextStrings)
                {
                    sw.WriteLine(text);

                    for (int i = 0; i < distinctTextStrings.Count; i++)
                    {
                        sw.Write(distinctTextStrings[i]);

                        // Add a comma if it's not the last text
                        if (i < distinctTextStrings.Count - 1)
                        {
                            sw.Write(",");
                        }
                    }
                }
            }


            return Result.Succeeded;
        }

        
    }
    
}
