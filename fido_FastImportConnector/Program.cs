using System;
using fido_FastImportConnector.Model;
using System.Collections.Generic;
using System.Text;

namespace fido_FastImportConnector
{
    class Program
    {
        static KcImporter kcimp;
        static void Main(string[] args)
        {
            kcimp = new KcImporter();
            // read command line arguments
            Dictionary<string, string> arguments = new Dictionary<string, string>();
            //if (args.Length % 2 == 0)
            {

                //fetching the arguments

                // test


                for (int i = 0; i < args.Length; i++)
                {
                    kcimp.ImportRoot = args[i] == "--importpath" ? args[i + 1] : "";
                    kcimp.ImportFilePattern = args[i] == "--extension" ? args[i + 1] : "";
                    kcimp.MoveToPath = args[i] == "--movetopath" ? args[i + 1] : "";
                    kcimp.BatchClassName = args[i] == "--batchclass" ? args[i + 1] : "";
                    kcimp.FormTypeName = args[i] == "--formtype" ? args[i + 1] : "";
                    kcimp.Username = args[i] == "--username" ? args[i + 1] : "";
                    kcimp.Password = args[i] == "--password" ? args[i + 1] : "";
                    int seconds
                }

                kcimp.BatchClassName = arguments.ContainsKey("--batchclass") ? arguments["--batchclass"] : "";
                kcimp.FormTypeName = arguments.ContainsKey("--formtype") ? arguments["--formtype"] : "";
                kcimp.ImportFilePattern = arguments.ContainsKey("--extension") ? arguments["--extension"] : "";
                kcimp.ImportRoot = arguments.ContainsKey("--importpath") ? arguments["--importpath"] : "";
                kcimp.MoveToPath = arguments.ContainsKey("--movetopath") ? arguments["--movetopath"] : "";
                kcimp.Password = arguments.ContainsKey("--password") ? arguments["--password"] : "";
                kcimp.Username = arguments.ContainsKey("--username") ? arguments["--username"] : "";
                int seconds;
                int.TryParse(arguments.ContainsKey("--seconds") ? arguments["--seconds"] : "0", out seconds);
                kcimp.Seconds = seconds;

                // map batch fields

                // map index fields


            }
            
            // some batch and index fields
            kcimp.BatchFields.Add("BatchField", "test"); // --batchfields key=value key=value
            kcimp.IndexFields.Add("DocumentGuid", "123"); // --indexfields key=value key=value


            // depending on the seconds property (whether set or not), we start a normal or timed import
            if (kcimp.Seconds != 0)
            {
                // alternative 1: a timed import (repeat the import every n seconds until canceled)
                ConsoleKeyInfo ki;

                Console.WriteLine("Timed import started. Hit ESC to cancel");
                kcimp.StartTimedImport();

                do
                {
                    ki = Console.ReadKey();
                } while (ki.Key != ConsoleKey.Escape);

                Console.WriteLine("Canceling timed import");
                kcimp.StopTimedImport();

                Console.WriteLine("Import canceled. Exiting.");

            }
            else
            {
                // alternative 2: execute the import once
                Console.WriteLine("Importing once");
                kcimp.ImportOnce();
                Console.WriteLine("Finished importing!");

            }

            Console.ReadKey();




        }
        
        /// <summary>
        /// shows the help in case the arguments were invalid
        /// </summary>
        private static void ShowHelp()
        {
            Console.WriteLine("illegal operation");
            Console.WriteLine("usage: --inputpath [path] --movetopath [path] --extension [ext] --username [name] --password [pass] --batchclass [name] --formtype [name] --batchfields [key=>value] --indexfields [key=>value]");
            Console.ReadKey();
        }
        
    }
}
