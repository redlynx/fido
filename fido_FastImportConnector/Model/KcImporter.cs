using System;
using System.Collections.Generic;
using System.Linq;
using Kofax.Capture.CaptureModule.InteropServices;
using System.IO;
using System.Threading;
using Newtonsoft.Json;

namespace fido_FastImportConnector.Model
{
    [Serializable]
    public class KcImporter
    {
        // public properties, exposed and serializable
        public string Username { get; set; }
        public string Password { get; set; }
        public string BatchClassName { get; set; }
        public string FormTypeName { get; set; }
        public string ImportDir { get; set; }
        public bool TopDirectoryOnly { get; set; } = false;
        public bool CreateDocumentPerFile { get; set; } = true;
        public string MoveToDir { get; set; }
        public string ImportFilePattern { get; set; }
        public int Seconds { get; set; }
        public Dictionary<string, string> BatchFields { get; set; }
        public Dictionary<string, string> IndexFields { get; set; }
        [JsonIgnore]
        public bool ImportRunning { get; private set; }
        private bool importValid = true;

        // KC objects
        private ImportLogin login;
        private IApplication app;
        private IBatchClass batchclass;
        private IBatch batch;
        private IDocument document, nulldocument;
        private IPage nullpage;
        private Guid g;

        private Timer t;

        // since we plan on deserializing objects of this class, we might only need this parameterless constructor
        public KcImporter()
        {
            Username = "";
            Password = "";
            BatchClassName = "";
            FormTypeName = "";
            ImportDir = @"C:\";
            MoveToDir = @"C:\";
            ImportFilePattern = "*.tif";
            Seconds = 0;
            BatchFields = new Dictionary<string, string>();
            IndexFields = new Dictionary<string, string>();
        }
        
        /// <summary>
        /// This method is called either every time the timer ticks (and the import isn't yet running), 
        /// or once in case the import has to be performed only a single time. An invalid configuration
        /// skips the import entirely, but the timer will still continue to run.
        /// </summary>
        /// <param name="state"></param>
        private void Import(Object state)
        {
            if (ImportRunning)
            {
                Console.WriteLine($"{DateTime.Now} import in progress, ignoring timer for now");
                return;
            }

            if (!importValid)
            {
                Console.WriteLine($"{DateTime.Now} import not possible due to invalid configuration.");
                return;
            }

            // fetch import files and only start when there's at least one file to import
            ImportRunning = true;            
            var searchOption = TopDirectoryOnly ? SearchOption.TopDirectoryOnly : SearchOption.AllDirectories;
            var files = Directory.EnumerateFiles(ImportDir, ImportFilePattern, searchOption);
            ImportFiles(files);
            ImportRunning = false;
        }

        private void ImportFiles(IEnumerable<string> importFiles)
        {
            if (importFiles.Count() == 0)
            {
                Console.WriteLine($"{DateTime.Now} no matching files found in import directory");
                return;
            }

            Console.WriteLine($"{DateTime.Now} {importFiles.Count()} files found");
            login = new ImportLogin();
            app = new Application();
            g = Guid.NewGuid();

            // potential problem: username or password are incorrect, user hasn't got all privilegues required 
            // as there is no other way to validate a user in KC, we make use of a try/catch
            try
            {
                app = login.Login(Username, Password);
            }
            catch (Exception e)
            {
                Console.WriteLine($"{DateTime.Now} import failed: {e.Message}");
                ImportRunning = false;
                importValid = false;
                return;
                // possible improvement: rethrow the error here.
            }

            // make sure the batch class exists
            if (!BatchClassExists(BatchClassName))
            {
                Console.WriteLine($"{DateTime.Now} the batch class ${BatchClassName} was not found");
                ImportRunning = false;
                importValid = false;
                return;
            }

            batchclass = app.BatchClasses[BatchClassName];
            batch = app.CreateBatch(ref batchclass, g.ToString());
            nulldocument = null;
            nullpage = null;

            // set batch fields
            SetAllBatchFields();

            importFiles.ToList().ForEach(file =>
            {
                // now, we have loose pages
                batch.ImportFile(file);
                if (CreateDocumentPerFile)
                {
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
                        Console.WriteLine($"{DateTime.Now} form type {FormTypeName} was not found");
                    }
                }
            });

            // close the batch to finalize its creation and logout
            Console.WriteLine($"{DateTime.Now} batch '{g}' successfully imported, {batch.Documents.Count} documents and {batch.LoosePageCount} loose pages");
            app.CloseBatch();
            login.Logout();
            

            // finally, we move the imported files to the MoveToPath location
            Directory.CreateDirectory(this.MoveToDir);
            foreach (var f in importFiles)
            {
                string movedFile = Path.Combine(MoveToDir, Path.GetFileName(f));
                if (File.Exists(movedFile))
                {
                    Console.WriteLine($"{DateTime.Now} target file {movedFile} already exists, overwriting");
                    File.Delete(movedFile);
                }
                File.Move(f, movedFile);
            }
        }


        /// <summary>
        /// Starts a timed import. Will ignore importing files if another import is in progress. 
        /// </summary>
        /// <param name="seconds">Seconds in between imports.</param>
        public void StartTimedImport(int seconds)
        {
            t = new Timer(Import, null, 0, seconds * 1000);
        }


        /// <summary>
        /// Starts a timed import. Will ignore importing files if another import is in progress. 
        /// </summary>
        public void StartTimedImport()
        {
            StartTimedImport(Seconds);           
        }


        /// <summary>
        /// Stops the timed import.
        /// </summary>
        public void StopTimedImport()
        {
            t.Dispose();
        }


        /// <summary>
        /// Imports files once.
        /// </summary>
        public void ImportOnce()
        {
            Import(null);
        }


        // helpers for checking misc Kofax objects
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
                if (ft.Name == FormTypeName)
                    return true;
            }
            return false;
        }


        private bool BatchFieldExists(string BatchFieldKey)
        {
            foreach (BatchField bf in batch.BatchFields)
            {
                if (bf.Name == BatchFieldKey)
                    return true;
            }
            return false;
        }


        private bool IndexFieldExists(string IndexFieldKey)
        {
            foreach (IndexField f in document.IndexFields)
            {
                if (f.Name == IndexFieldKey)
                    return true;
            }
            return false;
        }


        private void SetAllBatchFields()
        {
            foreach (var bf in BatchFields)
            {
                if (BatchFieldExists(bf.Key))
                    batch.BatchFields[bf.Key].Value = bf.Value;
            }
        }


        private void SetAllIndexFields()
        {
            foreach (var idx in IndexFields)
            {
                if (IndexFieldExists(idx.Key))
                    document.IndexFields[idx.Key].Value = idx.Value;
            }
        }
    }
}
