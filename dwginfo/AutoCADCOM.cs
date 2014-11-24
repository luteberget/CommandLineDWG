using Autodesk.AutoCAD.Interop;
using System;
using System.Runtime.InteropServices;

namespace dwginfo
{
    public sealed class AutoCADCOM : IDisposable
    {
        private AcadApplication app = null;
        private const string progid = "AutoCAD.Application";
        private bool reuse = true;

        public AcadApplication Context { get { return app; } }

        public AutoCADCOM()
        {
            try
            {
                var obj = Marshal.GetActiveObject(progid);
                if (obj != null)
                {
                    app = obj as AcadApplication;
                }
            }
            catch
            {
                System.Console.WriteLine("Note: Could not find running AutoCAD instance. Will launch, this is slow.");
            }

            if (app == null)
            {
                try
                {
                    var programtype = Type.GetTypeFromProgID(progid, true);
                    reuse = false;
                    var obj = Activator.CreateInstance(programtype);
                    app = obj as AcadApplication;
                }
                catch
                {
                }
            }

            if (app == null)
            {
                Dispose();
                throw new Exception("Could not load AutoCAD application context.");
            }
        }

        public void Dispose()
        {
            if (!reuse && app != null)
            {
                app.Quit();
            }

            if (app != null)
            {
                Marshal.ReleaseComObject(app);
            }

            app = null;

            // Make sure that the COM objects are released before
            // application exits. See
            // http://stackoverflow.com/questions/158706/how-to-properly-clean-up-excel-interop-objects
            // http://stackoverflow.com/questions/26925939/how-to-properly-close-dispose-excel-coms-in-c-sharp-excel-process-not-closing
            // http://stackoverflow.com/questions/25134024/clean-up-excel-interop-objects-with-idisposable

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    public sealed class AutoCADDocument : IDisposable
    {
        private AutoCADCOM context = null;
        private AcadDocument doc = null;

        public AcadDocument Document { get { return doc; } }

        public AutoCADDocument(string filename, bool rdonly = false)
        {
            try
            {
                context = new AutoCADCOM();
                var obj = context.Context.Documents.Open(filename, rdonly);
                doc = obj as AcadDocument;
            }
            catch (Exception e)
            {
                Dispose();
                throw e;
            }

            if (doc == null)
            {
                throw new Exception("Unknown error.");
            }
        }

        public void Dispose()
        {
            if (doc != null)
            {
                doc.Close();
            }

            doc = null;
            context.Dispose();
        }
    }
}