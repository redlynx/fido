using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
            BatchFields = new Dictionary<string, string>();
            IndexFields = new Dictionary<string, string>();
        }

        private void Import(Object state)
        {

            IEnumerable<string> importFiles;
            int docCount, pageCount;
            
            // fetch import files and only start when there's at least one file to import
            importFiles = Directory.EnumerateFiles(ImportRoot, ImportFilePattern, SearchOption.AllDirectories);

            if (importFiles.Count() > 0)
            {

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
                    throw e;
                    //return String.Format("{0} Import failed ({1})", DateTime.Now, e.Message);
                }
                
                // make sure the batch class exists
                if (!BatchClassExists(BatchClassName)) throw new Exception(String.Format("The Batch Class {0} was not found", BatchClassName)); //return string.Format("{0} The Batch Class {1} was not found", DateTime.Now, BatchClassName);
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
                    document = batch.CreateDocument(ref nulldocument);
                    // then, move all loose pages to that document and assign the form type
                    foreach (Page p in batch.LoosePages)
                    {
                        p.MoveToDocument(ref document, ref nullpage);
                    }
                    // if the form type exists, assign it. note that we won't throw an exception this time, but rather leave documents unassigned.
                    if (FormTypeExists(FormTypeName)) document.FormType = batch.FormTypes[FormTypeName];
                    
                    // then, set index fields
                    SetAllIndexFields();

                }

                docCount = batch.Documents.Count;
                pageCount = batch.LoosePageCount;

                // close the batch to finalize its creation and logout
                app.CloseBatch();
                login.Logout();
                //return String.Format("{0} Batch '{1}' successfully imported, {2} documents and {3} loose pages", DateTime.Now, g, docCount, pageCount);

                // finally, we move the imported files to the MoveToPath location
                Directory.CreateDirectory(this.MoveToPath);
                foreach (var f in importFiles)
                {
                    File.Move(f, Path.Combine(this.MoveToPath, Path.GetFileName(f)));
                }
            }
        }

        public void StartTimedImport(int seconds)
        {
            t = new Timer(this.Import, null, 0, seconds * 1000);
        }

        public void StopTimedImport()
        {
            t.Dispose();
        }

        public void ImportOnce()
        {
            Import(null);
        }

        private Boolean BatchClassExists(string BatchClassName)
        {
            foreach (BatchClass b in app.BatchClasses)
            {
                if (b.Name == BatchClassName) return true;   
            }
            return false;
        }

        private Boolean FormTypeExists(string FormTypeName)
        {
            foreach (FormType ft in batch.FormTypes)
            {
                if (ft.Name == FormTypeName) return true;
            }
            return false;
        }

        private Boolean BatchFieldExists(string BatchFieldKey)
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

        private Boolean IndexFieldExists(string IndexFieldKey)
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
