using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace INNOSetupScriptGeneratorForNetOfficeFW.Extensions
{
    public static class StringEx
    {
        public static string Sanitize(this string s)
        {
            return s.Replace(" ", "").Replace(".", "").Replace(",", "");
        }
    }
}
