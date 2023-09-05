using CommandLine;
using INNOSetupGeneratorForNetOfficeFW.Options;
using NetOfficeFwInstallTools;
using System;
using System.IO;

namespace INNOSetupRegistryForNetOfficeFW
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args)
            .WithParsed<Options>(o =>
            {
                INNOSetupGenerator generator = new INNOSetupGenerator(o.AddInPath, o.OfficeApps);
                File.WriteAllText(o.INNOScriptFilePath, generator.Execute());
            });
        }
    }
}
