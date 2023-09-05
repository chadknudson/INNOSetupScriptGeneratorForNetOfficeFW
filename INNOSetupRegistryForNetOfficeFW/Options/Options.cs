using CommandLine;

namespace INNOSetupGeneratorForNetOfficeFW.Options
{
    public class Options
    {
        [Option('a', "addin", Required = true, HelpText = "Full path and filename of the NetOfficeFw AddIn.")]
        public string AddInPath { get; set; }

        [Option('o', "officeapps", Required = true, HelpText = "Semi-colon separated list of office app(s) the AddIn supports. I.e. Word;Excel;Outlook.")]
        public string OfficeApps { get; set; }

        [Option('t', "target", Required = true, HelpText = "Full path and filename of the INNOSetup Script file to generate.")]
        public string INNOScriptFilePath { get; set; }

        [Option('v', "verbose", Required = false, HelpText = "Set output to verbose messages.")]
        public bool Verbose { get; set; }
    }
}
