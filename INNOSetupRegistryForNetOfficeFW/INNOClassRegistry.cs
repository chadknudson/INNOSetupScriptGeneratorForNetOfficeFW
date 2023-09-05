using NetOfficeFw.Build;
using System;
using System.Reflection;
using System.Text;

namespace NetOfficeFwInstallTools
{
    public class INNOClassRegistry
    {
        public string RegisterProgId(string progId, Guid guid)
        {
            StringBuilder sbProgId = new StringBuilder();

            sbProgId.AppendLine("Root: HKCU; Subkey: \"Software\\Classes\" ; ValueType: string; ValueName: ");

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

                var guidValue = guid.ToRegistryString().Substring(1);
                guidValue = guidValue.Substring(0, guidValue.Length - 1);

                string keyBase = $@"Software\Classes\{wow6432}CLSID\{guidValue}";
                sbComClass.AppendLine($"Root: HKCR; Subkey: \"{keyBase}\"; ValueType: string; ValueName: \"\"; ValueData: \"{progId}\"");
                sbComClass.AppendLine($"Root: HKCR; Subkey: \"{keyBase}\\Implemented Categories\"; ");
                sbComClass.AppendLine($"Root: HKCR; Subkey: \"{keyBase}\\Implemented Categories\\{{{{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}}}}\"; ");
                sbComClass.AppendLine($"Root: HKCR; Subkey: \"{keyBase}\\Implemented Categories\\{{{{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}}}}\"; ValueType: string; ValueName: \"\"; ValueData: \"{progId}\"");

                string keyInProcServer32 = keyBase + "\\InprocServer32";
                sbComClass.AppendLine($"Root: HKCR; Subkey: \"{keyInProcServer32}\"; ValueType: string; ValueName: \"\"; ValueData: \"mscoree.dll\"");
                sbComClass.AppendLine($"Root: HKCR; Subkey: \"{keyInProcServer32}\"; ValueType: string; ValueName: \"Assembly\"; ValueData: \"{assemblyName.FullName}\"");
                sbComClass.AppendLine($"Root: HKCR; Subkey: \"{keyInProcServer32}\"; ValueType: string; ValueName: \"Class\"; ValueData: \"{typeFullName}\"");
                sbComClass.AppendLine($"Root: HKCR; Subkey: \"{keyInProcServer32}\"; ValueType: string; ValueName: \"Codebase\"; ValueData: \"{assemblyCodebase}\"");
                sbComClass.AppendLine($"Root: HKCR; Subkey: \"{keyInProcServer32}\"; ValueType: string; ValueName: \"RuntimeVersion\"; ValueData: \"v4.0.30319\"");
                sbComClass.AppendLine($"Root: HKCR; Subkey: \"{keyInProcServer32}\"; ValueType: string; ValueName: \"ThreadingModel\"; ValueData: \"Both\"");

                return sbComClass.ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            return null;
        }

        public string RegisterOfficeAddin(string officeApp, string progId)
        {
            StringBuilder sbOfficeAddIn = new StringBuilder();
            string keyBase = $@"Software\Microsoft\Office\{officeApp}\Addins\{progId}";

            sbOfficeAddIn.AppendLine($"Root: HKCU; Subkey: \"{keyBase}\"; Flags: uninsdeletekey");
            sbOfficeAddIn.AppendLine($"Root: HKCU; Subkey: \"{keyBase}\"; ValueType: string; ValueName: \"\"; ValueData: \"{progId}\"");
            sbOfficeAddIn.AppendLine($"Root: HKCU; Subkey: \"{keyBase}\"; ValueType: string; ValueName: \"FriendlyName\"; ValueData: \"{progId}\"");
            sbOfficeAddIn.AppendLine($"Root: HKCU; Subkey: \"{keyBase}\"; ValueType: string; ValueName: \"Description\"; ValueData: \"{progId}\"");
            sbOfficeAddIn.AppendLine($"Root: HKCU; Subkey: \"{keyBase}\"; ValueType: dword; ValueName: \"LoadBehavior\"; ValueData: \"3\"");

            return sbOfficeAddIn.ToString();
        }
    }
}
