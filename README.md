# INNOSetupScriptGeneratorForNetOfficeFw

The INNO Setup [^1] Generator for NetOfficeFw[^2] helps to get you started on creating an installer for your NetOfficeFw AddIn. 

This is a command-line application that generates a basic INNO Setup script file for installing a NetOfficeFw COM AddIn for Microsoft Office. The application will load the AddIn DLL file and use reflection to identify all COM visible classes and create the necessary entries in the [Registry] section of the script. The generated script will produce a Setup.exe application that does not require Administrator privileges. All the registry entries for the DLL are created under HKEY_CURRENT_USER for the current user, and the installation folder defaults to a folder created under AppData for the user's Roaming Profile. The script file will also include the necessary files your .DLL depends upon by examing the files in the folder where your .DLL is located.

Make sure that you fill out the properties of your Visual Studio project in the AssemblyInfo.cs file as this tool will utilize those for configuring your INNO Setup script file. 

## Requirements

- The application requires a .NET Framework 4.8 or higher runtime environment.
- The application requires the INNO Setup Compiler to compile the generated script file into a setup.exe file.

## Installation

Clone the project and build it on your development machine. You will find the executable in the bin folder after you build the project

## Usage

The INNO Setup Script Generator for NetOffice FW is a command line tool that will read your COM 
AddIn .DLL file and generate a basic script file for the INNO Setup Compiler to build an application
setup for you. 

The generator creates a setup.exe file that will install your COM AddIn in User Space in the user's
roaming profile and perform all necessary Windows registry entries under HKEY_CURRENTUSER so that
administrator privileges are not required to run your setup application. You can change these options
in your generated INNO Setup script file if you wish.

Make certain that you have completed filling in the AssemblyInfo.cs file for your project before you run
the INNO Setup Script Generator. The generator will read the metadata that originates from the 
AssemblyInfo.cs file to get the information for your COM AddIn.

The program is self documenting, so you can run it with the --help option to see the command line parameters
that are available to you.

You can use it like this:

```
INNOSetupScriptGeneratorForNetOfficeFW.exe --help
```

The application takes three command-line parameters:

- `--addin`: The full path and filename of the .DLL file that implements the COM AddIn.
- `--officeapps`: A semi-colon separated list of Microsoft Office applications that the AddIn supports. The supported applications are Word, Excel, Access, and PowerPoint.
- `--target`: The full path and filename of the INNO Setup script file to generate.

For example, to generate a script file for an AddIn named MyAddIn.dll that supports Word and Excel, you can run the following command:

```
INNOSetupScriptGeneratorForNetOfficeFw.exe --addin "C:\MyAddIn\bin\Release\MyAddIn.dll" --officeapps "Word;Excel" --target "C:\MyAddIn\Setup\MyAddIn.iss"
```

You should use the .DLL file that is present in the bin\Release folder under your Visual Studio project for your AddIn for the --addin parameter value when running the app. This will enable the application to pickup all the required dependencies and include them in the [Files] section of the setup script. 

## Customizing Your INNO Setup Script File

After you generate the INNO Setup script file, you will need to edit the generated file to reflect your
installation options. The tool provides a roughed out script file that contains everything needed to 
install your program files and perform the necessary Windows registry operations to enable your COM AddIn
to run inside Microsoft Office. After you complete your customizations of the script file, you can use the
INNO Setup Compiler to build your setup.exe

## Generating your Setup.exe file with the INNOSetup Compiler

After you have created your INNO Setup script file for your application, you will need to run the INNO
Setup Compiler to build your setup.exe file. You can download INNOSetup from here: https://jrsoftware.org/isdl.php.

## Integrating the INNO Setup Compiler into your Visual Studio Project

You can integrate the INNO Setup Compiler into your Visual Studio project by adding a NuGet package
reference to the Tools.INNOSetup package. This will add the INNO Setup Compiler to your project and
you can then add a post build event to your project to run the INNO Setup Compiler to build your 
setup.exe file.

```
dotnet add package Tools.InnoSetup
```

In the following example, we are going to define several INNO Setup preprocessor variables that we will use
in our setup script file. We will define the SourcePath variable to be the path to the folder where our
project output files are located. We will define the ProjectDir preprocessor variable to be the path to the
root of our project folder. We will then run the INNO Setup Compiler to build our setup.exe file.

```
if exist "$(SolutionDir)packages\Tools.InnoSetup.6.2.2\tools\iscc.exe" (
    "$(SolutionDir)packages\Tools.InnoSetup.6.2.2\tools\iscc.exe" /dSourcePath="$(TargetDir)" /dProjectDir="$(ProjectDir)" /dTargetPath="$(TargetPath)" "$(ProjectDir)Setup\MyProgram.iss"
) else (
    echo Inno Setup Compiler not found. Skipping setup.exe creation.
)
```

By defining preprocessor variables for use in our INNO Setup script file, we can use the same script file
to build the setup.exe file for our project on our development machine and on our build server.

Thank you goes out to Jozef Izso for all his work on NetOfficeFw and NetOfficeFw.BuildTasks! 

### Supported Projects and Frameworks

INNO Setup Generator for NetOfficeFw supports.NET Frameworks 4.8.

[^1] [INNO Setup](https://jrsoftware.org/isinfo.php)
[^2] [NetOfficeFw](https://github.com/NetOfficeFw)
[^3] [INNO Setup Script Studio](http://www.innoscriptstudio.com/)
