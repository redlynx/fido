using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using fido_FastImportConnector.Model;
  
namespace fido_FastImportConnector
{
    class Program
    {
        static void Main(string[] args)
        {

            KcImporter kcimp = new KcImporter();

            kcimp.Username = "kfxservice"; // -u
            kcimp.Password = "kfxservice"; // -p
            kcimp.BatchClassName = "Test"; // -bc
            kcimp.FormTypeName = "Test";  // -ft
            kcimp.ImportRoot = @"C:\ProgramData\Kofax\CaptureSV\Projects\import\in"; // -i
            kcimp.MoveToPath = @"C:\ProgramData\Kofax\CaptureSV\Projects\import\done"; // -o
            kcimp.ImportFilePattern = "*.tif"; // -ext

            // some batch and index fields
            kcimp.BatchFields.Add("BatchField", "test"); 
            kcimp.IndexFields.Add("DocumentGuid", "123");

            // alternative 1: a timed import (repeat the import every n seconds until canceled)
            /*
            ConsoleKeyInfo ki;

            Console.WriteLine("Timed import started. Hit ESC to cancel");
            kcimp.StartTimedImport(60);

            do
            {
                ki = Console.ReadKey();
            } while (ki.Key != ConsoleKey.Escape);

            Console.WriteLine("Canceling timed import");
            kcimp.StopTimedImport();

            Console.WriteLine("Import canceled. Exiting.");
            */

            // alternative 2: execute the import once
            Console.WriteLine("Importing once");
            kcimp.ImportOnce();
            Console.WriteLine("Finished importing!");






            
        }
    }
}
