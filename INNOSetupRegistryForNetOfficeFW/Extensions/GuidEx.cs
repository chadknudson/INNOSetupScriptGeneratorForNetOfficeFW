using System;

namespace NetOfficeFw.Build
{
    public static class GuidEx
    {
        public static string ToRegistryString(this Guid guid)
        {
            return guid.ToString("B").ToUpperInvariant();
        }
        public static string ToRegistryStringNoBraces(this Guid guid)
        {
            string s = guid.ToString("B").ToUpperInvariant().Substring(1);
            return s.Substring(0, s.Length - 1);
        }
    }
}
