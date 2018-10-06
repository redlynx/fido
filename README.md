# fido - Fast Import Connector for Kofax Capture

fido is two things: a lightweight importer for Kofax Capture that makes use of configuration files in JSON format, and a wrapper for the Kofax Import API.

## Sample Usage and Logging
Just call fido with a configuration JSON as the single program argument:
```
fido_FastImportConnector.exe fido_config_test.json
```

If you require a log, just redirect stdout:
```
fido_FastImportConnector.exe fido_config_test.json >> C:\fido.log
```

## Configuration Files
fido makes use of simple configuraton files in JSON format. Here is a very simple example - more samples are included in the `SampleConfigs` folder.

```
{
  "Username": "admin",
  "Password": "",
  "BatchClassName": "Test",
  "FormTypeName": "Test",
  "ImportDir": "C:\\ProgramData\\Kofax\\CaptureSV\\Projects\\import\\fido",
  "TopDirectoryOnly": true,
  "CreateDocumentPerFile": true,
  "MoveToDir": "C:\\ProgramData\\Kofax\\CaptureSV\\Projects\\import\\fido\\done",
  "ImportFilePattern": "*.tif",
  "Seconds": 10,
  "BatchFields": {
    "Fruit": "Banana"
  },
  "IndexFields": {
    "Name0": "Jane"
  }
}

```

* `Username`: the user to log onto Kofax Capture, may be blank if User Profiles are disabled. Make sure the user may access both the Batch Class as well as the Scan module!
* `Password`: the user's password.
* `BatchClassName`: name of the Batch Class to be used.
* `FormTypeName`: name of the Form Type to be used. Determines the Document Class and available Index Fields.
* `ImportDir`: the directory to parse. 
* `TopDirectoryOnly`: if set to true, will only look in the ImportRoot directory. If false, will consider subfolders as well.
* `CreateDocumentPerFile`: if set to true, will create one document per imported file.
* `MoveToDir`: all imported files will be moved to this directory (and overwritten, if applicable).
* `ImportFilePattern`: considers files that match this search pattern only.
* `Seconds`: will check for new files in ImportDir every n seconds. If set to 0, will import only once and then terminate.
* `BatchFields`and `IndexFields`: will assign a value to the respective field. If you assing a value to an Index Field each document will use that value.

## KcImporter class
Alternatively, the KcImport class can be used as a wrapper. The following sample imports TIFF files, looking at the top directory only:

```
var kcImp = new KcImporter();
kcImp.Username = "admin";
kcImp.Password = "";
kcImp.BatchClassName = "Test";
kcImp.ImportDir = @"C:\ProgramData\Kofax\CaptureSV\Projects\import\fido";
kcImp.MoveToDir = @"C:\ProgramData\Kofax\CaptureSV\Projects\import\fido\done";
kcImp.TopDirectoryOnly = true;
kcImp.ImportFilePattern = "*.tif";
kcImp.ImportOnce();
```

