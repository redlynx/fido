using System;
using fido_FastImportConnector.Model;
using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json;
using System.IO;

namespace fido_FastImportConnector
{
    class Program
    {
        static void Main(string[] args)
        {
            KcImporter kcImp;
            // a config file is mandatory and has to be passed as the single program argument
            if (args.Length != 1)
            {
                Console.WriteLine("Please provide a valid configuration file as the single program argument.");
                return;
            }

            string settingsJson = args[0];

            try
            {
                kcImp = JsonConvert.DeserializeObject<KcImporter>(File.ReadAllText(settingsJson));
            }
            catch (Exception)
            {
                throw;
            }
            
            if (kcImp.Seconds == 0)
            {
                kcImp.ImportOnce();
            }
            else
            {
                kcImp.StartTimedImport();
                // keep the program in an endless loop
                Console.WriteLine("Hit CTRL+C to cancel.");
                do
                {
                    
                } while (true);

            }
        }

        static void SampleApiCall()
        {
            var kcImp = new KcImporter();
            kcImp.Username = "admin";
            kcImp.Password = "";
            kcImp.BatchClassName = "Test";
            kcImp.ImportDir = @"C:\ProgramData\Kofax\CaptureSV\Projects\import\fido";
            kcImp.MoveToDir = @"C:\ProgramData\Kofax\CaptureSV\Projects\import\fido\done";
            kcImp.TopDirectoryOnly = true;
            kcImp.ImportFilePattern = "*.tif";
            kcImp.ImportOnce();
        }
    }
}
