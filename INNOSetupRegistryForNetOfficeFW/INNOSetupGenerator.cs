﻿using INNOSetupScriptGeneratorForNetOfficeFW.Extensions;
using NetOfficeFw.Build;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;

namespace NetOfficeFwInstallTools
{
    public class INNOSetupGenerator
    {
        public string AddInPath { get; set; }
        public string AppName { get; set; }
        public string Version { get; set; }
        public string CompanyName { get; set; }
        public string Description { get; set; }
        public string[] OfficeApps { get; set; }
        public INNOSetupGenerator(string addInPath, string officeApps)
        {
            AddInPath = addInPath;
            OfficeApps = officeApps.Split(new char[] { ';' });
        }

        public string Execute()
        {
            StringBuilder sbInstallationScript = new StringBuilder();

            sbInstallationScript.AppendLine("; Script generated by the Inno Setup Script Generator for NetOfficeFW.");
            sbInstallationScript.AppendLine("; SEE THE INNO SETUP DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!");
            sbInstallationScript.AppendLine();

            sbInstallationScript.Append(GetApplicationInformation(AddInPath));
            sbInstallationScript.Append(GenerateSetup());
            sbInstallationScript.Append(GenerateFiles());
            sbInstallationScript.Append(GenerateRegistry());

            return sbInstallationScript.ToString();
        }

        public string GenerateRegistry()
        {
            StringBuilder innoRegistry = new StringBuilder();
            AppDomain domain = null;
            try
            {
                var assemblyPath = AddInPath;
                var assemblyDir = Path.GetDirectoryName(assemblyPath);

                AppDomain.CurrentDomain.ReflectionOnlyAssemblyResolve += (sender, args) => AssemblyEx.ReflectionOnlyAssemblyResolve(args, assemblyDir);

                var assembly = Assembly.LoadFrom(assemblyPath);
                var assemblyName = assembly.GetName();
                var assemblyCodebase = assemblyPath.GetCodebase();
                var publicTypes = assembly.GetExportedTypes();

                innoRegistry.AppendLine("[Registry]");

                foreach (var publicType in publicTypes)
                {
                    var isComVisible = publicType.IsComVisibleType();
                    var isAddinType = publicType.IsComAddinType();

                    if (isComVisible)
                    {
                        var name = publicType.Name;
                        var guid = publicType.GUID;
                        var progId = publicType.GetProgId();

                        Console.WriteLine($@"Processing class {progId} with GUID {guid.ToRegistryString()}");

                        INNOClassRegistry classRegistry = new INNOClassRegistry();

                        innoRegistry.Append(classRegistry.RegisterProgId(progId, guid));

                        innoRegistry.Append(classRegistry.RegisterComClassNative(progId, guid, assemblyName, assemblyCodebase, publicType.FullName));

                        if (Environment.Is64BitOperatingSystem)
                        {
                            innoRegistry.Append(classRegistry.RegisterComClassWOW6432(progId, guid, assemblyName, assemblyCodebase, publicType.FullName));
                        }

                        if (isAddinType && OfficeApps != null)
                        {
                            foreach (var officeApp in OfficeApps)
                            {
                                Console.WriteLine($@"Registering add-in {progId} to Microsoft Office {officeApp}");
                                innoRegistry.Append(classRegistry.RegisterOfficeAddin(officeApp, progId, AppName, AppName));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                if (domain != null)
                {
                    AppDomain.Unload(domain);
                }
            }
            return innoRegistry.ToString();
        }

        public string GenerateSetup()
        {
            StringBuilder sbSetup = new StringBuilder();

            string companyName = CompanyName.Sanitize();
            string appName = AppName.Sanitize();
            string outputFolder = Path.GetDirectoryName(AddInPath);
            string iconFilePath = FindIcoFile(outputFolder);

            int binPosition = outputFolder.IndexOf("\\bin\\");
            if (binPosition > 0)
            {
                outputFolder = outputFolder.Substring(0, binPosition) + $"\\{appName}.Setup\\Output";
            }

            sbSetup.AppendLine("[Setup]");
            sbSetup.AppendLine("; NOTE: The value of AppId uniquely identifies this application.");
            sbSetup.AppendLine("; Do not use the same AppId value in installers for other applications.");
            sbSetup.AppendLine("; (To generate a new GUID, click Tools | Generate GUID inside the IDE.");
            sbSetup.AppendLine("AppId = {" + Guid.NewGuid().ToRegistryString());
            sbSetup.AppendLine($"AppName = \"{AppName}\"");
            sbSetup.AppendLine("AppVersion ={#MyAppVersion}");
            sbSetup.AppendLine("AppVerName ={#MyAppName} {#MyAppVersion}");
            sbSetup.AppendLine($"AppPublisher = \"{CompanyName}\"");
            sbSetup.AppendLine("AppPublisherURL ={#MyAppURL}");
            sbSetup.AppendLine("AppSupportURL ={#MyAppURL}");
            sbSetup.AppendLine("AppUpdatesURL ={#MyAppURL}");
            sbSetup.AppendLine($"DefaultDirName ={{userappdata}}\\{companyName}\\{appName}");
            sbSetup.AppendLine("DefaultGroupName ={#MyAppName}");
            sbSetup.AppendLine($"OutputDir={outputFolder}");
            sbSetup.AppendLine($"OutputBaseFilename = {appName}Setup");
            if (!string.IsNullOrEmpty(iconFilePath))
            {
                string iconFile = Path.GetFileName(iconFilePath);
                sbSetup.AppendLine($"SetupIconFile = \"{{ProjectDir}}{iconFile}\"");
            }
            else
            {
                sbSetup.AppendLine("; SetupIconFile = \"path to your icon file goes here\"");
            }
            sbSetup.AppendLine("Compression = lzma");
            sbSetup.AppendLine("SolidCompression = yes");
            sbSetup.AppendLine("PrivilegesRequired = lowest");
            sbSetup.AppendLine();
            sbSetup.AppendLine("[Languages]");
            sbSetup.AppendLine("Name: \"english\"; MessagesFile: \"compiler:Default.isl\"");
            sbSetup.AppendLine();

            return sbSetup.ToString();
        }
        public string GetApplicationInformation(string path)
        {
            StringBuilder sbFileInfo = new StringBuilder();

            FileVersionInfo fileInfo = FileVersionInfo.GetVersionInfo(path);

            CompanyName = fileInfo.CompanyName;
            Version = fileInfo.ProductVersion;
            AppName = fileInfo.ProductName;

            sbFileInfo.AppendLine($"#define MyAppName \"{fileInfo.ProductName}\"");
            sbFileInfo.AppendLine($"#define MyAppVersion \"{fileInfo.ProductVersion}\"");
            sbFileInfo.AppendLine($"#define MyAppPublisher \"{fileInfo.CompanyName}\"");
            sbFileInfo.AppendLine($"#define MyAppURL \"https://your-app-url-here\"");
            sbFileInfo.AppendLine();

            return sbFileInfo.ToString();
        }

        public string GenerateFiles()
        {
            StringBuilder sbFiles = new StringBuilder();
            
            string folder = Path.GetDirectoryName(AddInPath);

            if (Directory.Exists(folder))
            {
                sbFiles.AppendLine("[Files]");
                string[] files = Directory.GetFiles(folder);

                foreach (string filePath in files)
                {
                    string file = Path.GetFileName(filePath);

                    if (file.EndsWith(".pdb") || file.EndsWith(".xml"))
                        continue;

                    sbFiles.AppendLine($"Source: \"{{#SourcePath}}{file}\"; DestDir: \"{{app}}\"; Flags: ignoreversion");
                }
                sbFiles.AppendLine();
            }
            else
            {
                Console.WriteLine($"Directory {folder} does not exist.");
            }

            return sbFiles.ToString();
        }

        public static string FindIcoFile(string folder)
        {
            DirectoryInfo currentDirectory = new DirectoryInfo(folder);

            while (currentDirectory != null)
            {
                // Search for .ico files in the current directory
                var icoFiles = currentDirectory.GetFiles("*.ico");

                if (icoFiles.Length > 0)
                {
                    return icoFiles[0].FullName; // Return the first .ico file found
                }

                currentDirectory = currentDirectory.Parent; // Move up to the parent directory
            }

            return null; // Return null if no .ico file was found up to the root
        }
    }
}