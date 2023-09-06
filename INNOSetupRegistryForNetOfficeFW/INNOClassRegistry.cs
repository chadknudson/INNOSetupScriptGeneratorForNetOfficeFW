using NetOfficeFw.Build;
using System;
using System.IO;
using System.Reflection;
using System.Text;

namespace NetOfficeFwInstallTools
{
    public class INNOClassRegistry
    {
        public string RegisterProgId(string progId, Guid guid)
        {
            StringBuilder sbProgId = new StringBuilder();
            string regGuid = guid.ToRegistryString();
            sbProgId.AppendLine($"Root: HKCU; Subkey: \"Software\\Classes\\{progId}\" ; ValueType: string; ValueName: \"\"; ValueData: \"{progId}\"; Flags: uninsdeletekey");
            sbProgId.AppendLine($"Root: HKCU; Subkey: \"Software\\Classes\\{progId}\\CLSID\" ; ValueType: string; ValueName: \"\"; ValueData: \"{{{regGuid}\"");
            return sbProgId.ToString();
        }

        public string RegisterComClassNative(string progId, Guid guid, AssemblyName assemblyName, string assemblyCodebase, string typeFullName)
        {
            return RegisterComClass(@"", progId, guid, assemblyName, assemblyCodebase, typeFullName);
        }

        public string RegisterComClassWOW6432(string progId, Guid guid, AssemblyName assemblyName, string assemblyCodebase, string typeFullName)
        {
            return RegisterComClass(@"WOW6432Node\", progId, guid, assemblyName, assemblyCodebase, typeFullName);
        }

        public string RegisterComClass(string wow6432, string progId, Guid guid, AssemblyName assemblyName, string assemblyCodebase, string typeFullName)
        {
            try
            {
                StringBuilder sbComClass = new StringBuilder();
                string win64Check = (string.IsNullOrEmpty(wow6432) ? null : "; Check: IsWin64");
                string bit = (string.IsNullOrEmpty(wow6432) ? "32" : "64");
                var guidValue = guid.ToRegistryString();

                string installedAddinPath = "file:///{app}\\" + Path.GetFileName(assemblyCodebase);

//                string keyBase = $@"Software\Classes\{wow6432}CLSID\{{{guidValue}}}";
                string keyBase = $@"Software\Classes\CLSID\{{{guidValue}";
                sbComClass.AppendLine($"Root: HKCU{bit}; Subkey: \"{keyBase}\"; ValueType: string; ValueName: \"\"; ValueData: \"{progId}\"; Flags: uninsdeletekey{win64Check}");
                sbComClass.AppendLine($"Root: HKCU{bit}; Subkey: \"{keyBase}\\Implemented Categories\"; ValueType: string; ValueName: \"\"{win64Check}");
                sbComClass.AppendLine($"Root: HKCU{bit}; Subkey: \"{keyBase}\\Implemented Categories\\{{{{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}}}}\"{win64Check}");
                sbComClass.AppendLine($"Root: HKCU{bit}; Subkey: \"{keyBase}\\Implemented Categories\\{{{{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}}}}\"; ValueType: string; ValueName: \"\"; ValueData: \"{progId}\"{win64Check}");

                string keyInProcServer32 = keyBase + "\\InprocServer32";
                sbComClass.AppendLine($"Root: HKCU{bit}; Subkey: \"{keyInProcServer32}\"; ValueType: string; ValueName: \"\"; ValueData: \"mscoree.dll\"{win64Check}");
                sbComClass.AppendLine($"Root: HKCU{bit}; Subkey: \"{keyInProcServer32}\"; ValueType: string; ValueName: \"Assembly\"; ValueData: \"{assemblyName.FullName}\"{win64Check}");
                sbComClass.AppendLine($"Root: HKCU{bit}; Subkey: \"{keyInProcServer32}\"; ValueType: string; ValueName: \"Class\"; ValueData: \"{typeFullName}\"{win64Check}");
                sbComClass.AppendLine($"Root: HKCU{bit}; Subkey: \"{keyInProcServer32}\"; ValueType: string; ValueName: \"Codebase\"; ValueData: \"{installedAddinPath}\"{win64Check}");
                sbComClass.AppendLine($"Root: HKCU{bit}; Subkey: \"{keyInProcServer32}\"; ValueType: string; ValueName: \"RuntimeVersion\"; ValueData: \"v4.0.30319\"{win64Check}");
                sbComClass.AppendLine($"Root: HKCU{bit}; Subkey: \"{keyInProcServer32}\"; ValueType: string; ValueName: \"ThreadingModel\"; ValueData: \"Both\"{win64Check}");

                return sbComClass.ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            return null;
        }

        public string RegisterOfficeAddin(string officeApp, string progId, string friendlyName, string description)
        {
            StringBuilder sbOfficeAddIn = new StringBuilder();
            string keyBase = $@"Software\Microsoft\Office\{officeApp}\Addins\{progId}";

            sbOfficeAddIn.AppendLine($"Root: HKCU; Subkey: \"{keyBase}\"; Flags: uninsdeletekey");
            sbOfficeAddIn.AppendLine($"Root: HKCU; Subkey: \"{keyBase}\"; ValueType: string; ValueName: \"\"; ValueData: \"{progId}\"");
            sbOfficeAddIn.AppendLine($"Root: HKCU; Subkey: \"{keyBase}\"; ValueType: string; ValueName: \"FriendlyName\"; ValueData: \"{friendlyName}\"");
            sbOfficeAddIn.AppendLine($"Root: HKCU; Subkey: \"{keyBase}\"; ValueType: string; ValueName: \"Description\"; ValueData: \"{description}\"");
            sbOfficeAddIn.AppendLine($"Root: HKCU; Subkey: \"{keyBase}\"; ValueType: dword; ValueName: \"LoadBehavior\"; ValueData: \"3\"");

            return sbOfficeAddIn.ToString();
        }
    }
}
