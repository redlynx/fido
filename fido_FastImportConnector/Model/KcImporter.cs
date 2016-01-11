using System;
using System.Collections.Generic;
using System.Linq;
using Kofax.Capture.CaptureModule.InteropServices;
using System.IO;
using System.Threading;

namespace fido_FastImportConnector.Model
{
    public class KcImporter
    {

        public string Username { get; set; }
        public string Password { get; set; }
        public string BatchClassName { get; set; }
        public string FormTypeName { get; set; }
        public string ImportRoot { get; set; }
        public string MoveToPath { get; set; }
        public string ImportFilePattern { get; set; }
        public int Seconds { get; set; }
        public bool ImportRunning { get; private set; }
        public Dictionary<string, string> BatchFields { get; set; }
        public Dictionary<string, string> IndexFields { get; set; }

        private ImportLogin login; // = new ImportLogin();
        private IApplication app; // = new Application();
        private IBatchClass batchclass;
        private IBatch batch;
        private IDocument document, nulldocument;
        private IPage nullpage;
        private Guid g;

        private Timer t;

        public KcImporter()
        {
            Username = "";
            Password = "";
            BatchClassName = "";
            FormTypeName = "";
            ImportRoot = @"C:\";
            MoveToPath = @"C:\";
            ImportFilePattern = "*.tif";
            Seconds = 0;
            BatchFields = new Dictionary<string, string>();
            IndexFields = new Dictionary<string, string>();
        }
        
        private void Import(Object state)
        {
            if (ImportRunning)
            {
                Console.WriteLine(String.Format("{0} import in progress, ignoring timer for now", DateTime.Now));
                return;
            }
                
            ImportRunning = true;
            IEnumerable<string> importFiles;
            int docCount, pageCount;
            
            // fetch import files and only start when there's at least one file to import
            importFiles = Directory.EnumerateFiles(ImportRoot, ImportFilePattern, SearchOption.AllDirectories);

            if (importFiles.Count() > 0)
            {
                Console.WriteLine(String.Format("{0} {1} files found", DateTime.Now, importFiles.Count()));
                login = new ImportLogin();
                app = new Application();
                g = Guid.NewGuid();

                // potential problem: username or password are incorrect, user hasn't got all privilegues required 
                // as there is no other way to validate a user, we make use of a try/catch
                try
                {
                    app = login.Login(Username, Password);
                }
                catch (Exception e)
                {
                    Console.WriteLine(String.Format("{0} import failed: {1}", DateTime.Now, e.Message));
                    ImportRunning = false;
                    return;
                }

                // make sure the batch class exists
                if (!BatchClassExists(BatchClassName))
                {
                    Console.WriteLine(String.Format("{0} the batch class {1} was not found", DateTime.Now, BatchClassName));
                    ImportRunning = false;
                    return;
                    //throw new Exception(String.Format("The Batch Class {0} was not found", BatchClassName)); 
                }

                batchclass = app.BatchClasses[BatchClassName];
                batch = app.CreateBatch(ref batchclass, g.ToString());
                nulldocument = null;
                nullpage = null;

                // set batch fields
                SetAllBatchFields();

                foreach (var f in importFiles)
                {
                    // now, we have loose pages
                    batch.ImportFile(f);
                    // create a document per attachment
                    // note: if we wanted to create loose pages, we need to skip creating a document and moving pages to it entirely
                    document = batch.CreateDocument(ref nulldocument);
                    // then, move all loose pages to that document and assign the form type
                    foreach (Page p in batch.LoosePages)
                    {
                        p.MoveToDocument(ref document, ref nullpage);
                    }
                    // if the form type exists, assign it. note that we won't throw an exception this time, but rather leave documents unassigned.
                    if (FormTypeExists(FormTypeName))
                    {
                        document.FormType = batch.FormTypes[FormTypeName];
                        // then, set index fields
                        SetAllIndexFields();
                    }
                    else
                    {
                        Console.WriteLine(String.Format("{0} form type {1} was not found", DateTime.Now, FormTypeName));
                    }

                }

                docCount = batch.Documents.Count;
                pageCount = batch.LoosePageCount;

                // close the batch to finalize its creation and logout
                app.CloseBatch();
                login.Logout();
                Console.WriteLine(String.Format("{0} batch '{1}' successfully imported, {2} documents and {3} loose pages", DateTime.Now, g, docCount, pageCount));

                // finally, we move the imported files to the MoveToPath location
                Directory.CreateDirectory(this.MoveToPath);
                foreach (var f in importFiles)
                {
                    File.Move(f, Path.Combine(this.MoveToPath, Path.GetFileName(f)));
                }
            }
            else
            {
                Console.WriteLine(String.Format("{0} {1}", DateTime.Now, "no matching files found in import directory"));
            }

            ImportRunning = false;
        }

        public void StartTimedImport(int seconds)
        {
            t = new Timer(Import, null, 0, seconds * 1000);
        }

        public void StartTimedImport()
        {
            StartTimedImport(Seconds);           
        }

        public void StopTimedImport()
        {
            t.Dispose();
        }

        public void ImportOnce()
        {
            Import(null);
        }

        private bool BatchClassExists(string BatchClassName)
        {
            foreach (BatchClass b in app.BatchClasses)
            {
                if (b.Name == BatchClassName) return true;   
            }
            return false;
        }

        private bool FormTypeExists(string FormTypeName)
        {
            foreach (FormType ft in batch.FormTypes)
            {
                if (ft.Name == FormTypeName) return true;
            }
            return false;
        }

        private bool BatchFieldExists(string BatchFieldKey)
        {
            foreach (BatchField bf in batch.BatchFields)
            {
                if (bf.Name == BatchFieldKey)
                {
                    return true;
                }
            }
            return false;
        }

        private bool IndexFieldExists(string IndexFieldKey)
        {
            foreach (IndexField f in document.IndexFields)
            {
                if (f.Name == IndexFieldKey)
                {
                    return true;
                }
            }
            return false;
        }

        private void SetAllBatchFields()
        {
            foreach (var bf in BatchFields)
            {
                if (BatchFieldExists(bf.Key))
                {
                    batch.BatchFields[bf.Key].Value = bf.Value;
                }
            }
        }
        private void SetAllIndexFields()
        {
            foreach (var idx in IndexFields)
            {
                if (IndexFieldExists(idx.Key))
                {
                    document.IndexFields[idx.Key].Value = idx.Value;
                }
            }
        }


    }
}
