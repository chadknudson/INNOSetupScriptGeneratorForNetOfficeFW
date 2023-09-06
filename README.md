# INNOSetupScriptGeneratorForNetOfficeFw

This is a command-line application that generates a basic INNO Setup script file for installing a NetOfficeFw COM AddIn for Microsoft Office. The application will load the AddIn DLL file and use reflection to identify all COM visible classes and create the necessary entries in the [Registry] section of the script. The generated script will produce a Setup.exe application that does not require Administrator privileges. All the registry entries for the DLL are created under HKEY_CURRENT_USER for the current user, and the installation folder defaults to a folder created under AppData for the user's Roaming Profile. The script file will also include the necessary files your .DLL depends upon by examing the files in the folder where your .DLL is located.

Make sure that you fill out the properties of your Visual Studio project in the AssemblyInfo.cs file as this tool will utilize those for configuring your INNO Setup script file. 

## Requirements

- The application requires a .NET Framework 4.8 or higher runtime environment.
- The application requires the INNO Setup Compiler to compile the generated script file into a setup.exe file.

## Usage

The application takes three command-line parameters:

- `--addin`: The full path and filename of the .DLL file that implements the COM AddIn.
- `--officeapps`: A semi-colon separated list of Microsoft Office applications that the AddIn supports. The supported applications are Word, Excel, Access, and PowerPoint.
- `--target`: The full path and filename of the INNO Setup script file to generate.

For example, to generate a script file for an AddIn named MyAddIn.dll that supports Word and Excel, you can run the following command:

```
INNOSetupScriptGeneratorForNetOfficeFw.exe --addin "C:\MyAddIn\MyAddIn.dll" --officeapps "Word;Excel" --target "C:\MyAddIn\MyAddIn.iss"
```

You should use the .DLL file that is present in the bin\Release folder under your Visual Studio project for your AddIn for the --addin parameter value when running the app. This will enable the application to pickup all the required dependencies and include them in the [Files] section of the setup script. 

