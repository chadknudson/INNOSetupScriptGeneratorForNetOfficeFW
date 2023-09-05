using System;

namespace NetOfficeFw.Build
{
    public static class GuidEx
    {
        public static string ToRegistryString(this Guid guid)
        {
            return guid.ToString("B").ToUpperInvariant();
        }
    }
}
