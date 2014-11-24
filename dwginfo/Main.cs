using Autodesk.AutoCAD.Interop;
using System;
using System.IO;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;

namespace dwginfo
{
    public class DWGInfo
    {
        public const int ConsoleWidth = 80;

        public static void PrintObjectTypeSummary(AcadDocument doc, string collectionName)
        {
            Dictionary<string, int> count = new Dictionary<string, int>();
            dynamic collection = doc.GetType().InvokeMember(collectionName, BindingFlags.GetProperty, null, doc, null);
            foreach (dynamic item in collection)
            {
                string objectName = item.ObjectName as string;
                if (!String.IsNullOrWhiteSpace(objectName))
                {
                    if (objectName.StartsWith("Ac"))
                    {
                        objectName = objectName.Substring(2);
                    }
                    if (objectName.StartsWith("Db"))
                    {
                        objectName = objectName.Substring(2);
                    }

                    int value = 0;
                    count.TryGetValue(objectName, out value);
                    count[objectName] = value + 1;
                }
            }

            System.Console.WriteLine(collectionName + " summary:");
            foreach (KeyValuePair<string, int> pair in count)
            {
                System.Console.WriteLine(String.Format(" - {0}x {1}", pair.Value, pair.Key));
            }

        }

        public static void PrintSummaryInfo(dynamic summaryInfo)
        {
            string retval = "";
            string[] properties = { "Title", "Author", "Subject", "RevisionNumber", "Comments", "LastSavedBy", "Keywords" };
            foreach (string property in properties)
            {
                string value = summaryInfo.GetType().InvokeMember(property, BindingFlags.GetProperty, null, summaryInfo, null);
                if (!String.IsNullOrWhiteSpace(value))
                {
                    retval += property + ": " + value + "\n";
                }
            }

            if (String.IsNullOrWhiteSpace(retval))
            {
                retval = "(No document summary information)\n";
            }

            System.Console.Write(retval);
        }

        public static void PrintCollectionInfo(
            AcadDocument doc,
            string collectionName,
            Action<dynamic, List<string>> action = null,
            bool printIfZero = true,
            string printName = null)
        {
            List<string> names = new List<string>();

            dynamic collection = doc.GetType().InvokeMember(collectionName, BindingFlags.GetProperty, null, doc, null);

            foreach (dynamic item in collection)
            {
                if (action != null)
                {
                    action(item, names);
                }
                else
                {
                    names.Add(item.Name);
                }
            }
            if (printIfZero || names.Count > 0)
            {
                string status = (printName ?? collectionName) + ": ";
                status += StringTools.FormatCountAndExamples(names, ConsoleWidth - status.Length);
                System.Console.WriteLine(status);
            }
        }

        // STAThread is required to use dialog boxes.
        // See http://stackoverflow.com/questions/1361033/what-does-stathread-do
        [STAThread]
        private static void Main(string[] args)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "DWG files (*.dwg)|*.dwg";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                string fn = dlg.FileName;
                try
                {
                    using (var ac = new AutoCADDocument(fn, true))
                    {
                        System.Console.WriteLine("Info on AutoCAD DWG file: " + Path.GetFileName(fn));
                        AcadDocument doc = ac.Document;

                        PrintSummaryInfo(doc.SummaryInfo);

                        PrintCollectionInfo(doc, "Blocks", printIfZero: true, action: (item, names) =>
                        {
                            // Standard layout blocks:
                            // See http://help.autodesk.com/view/ACD/2015/ENU/?guid=GUID-E6F7B03B-F5CC-4A18-9C48-BBF1D32A31FD
                            if (item.Name.StartsWith("*MODEL_SPACE", true, null)) { return; }
                            if (item.Name.StartsWith("*PAPER_SPACE", true, null)) { return; }

                            // Anonymous blocks (mostly for hatch patterns and dimensions)
                            // See http://knowledge.autodesk.com/support/autocad/troubleshooting/caas/sfdcarticles/sfdcarticles/Anonymous-blocks-explained.html
                            if (item.Name.StartsWith("*", true, null)) { return; }

                            // Other blocks:
                            names.Add(item.Name);
                        });

                        string[] defaultcollections = { "Layers" };
                        foreach (var i in defaultcollections)
                        {
                            PrintCollectionInfo(doc, i, null, true);
                        }

                        string[] optionalcollections = { "Groups", "Layouts" };
                        foreach (var i in optionalcollections)
                        {
                            PrintCollectionInfo(doc, i, null, false);
                        }

                        PrintCollectionInfo(doc, "FileDependencies", printIfZero: false,
                            action: (item, names) => { names.Add(item.FileName); });

                        PrintCollectionInfo(doc, "Dictionaries", printIfZero: false, printName: "Xrecords",
                            action: (item, names) => { if (item.ObjectName == "AcDbXrecord") names.Add(item.Name); });

                        PrintCollectionInfo(doc, "RegisteredApplications", printIfZero: true,
                            action: (item, names) => { names.Add(item.Name); });


                        PrintObjectTypeSummary(doc, "ModelSpace");
                    }
                }
                catch (Exception e)
                {
                    System.Console.WriteLine(e);
                    System.Console.WriteLine("Main() Error: " + e.Message);
                }
            }
        }
    }
}