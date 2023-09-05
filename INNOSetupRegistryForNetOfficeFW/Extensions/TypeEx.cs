using System;
using System.Linq;
using System.Runtime.InteropServices;

namespace NetOfficeFw.Build
{
    public static class TypeEx
    {
        public static string ComVisibleAttributeName = typeof(ComVisibleAttribute).FullName;
        public static string ProgIdAttributeName = typeof(ProgIdAttribute).FullName;

        public static bool IsComVisibleType(this Type type)
        {
            var data = type.GetCustomAttributesData();
            return data.Any(attr => attr.AttributeType.FullName == ComVisibleAttributeName);
        }

        public static bool IsComAddinType(this Type type)
        {
            var interfaces = type.GetInterfaces();
            return interfaces.Any(t => t.Name == "IDTExtensibility2");
        }

        public static string GetProgId(this Type type)
        {
            var data = type.GetCustomAttributesData();
            var attr = data.FirstOrDefault(at => at.AttributeType.FullName == ProgIdAttributeName);
            var progIdArg = attr?.ConstructorArguments.FirstOrDefault();
            return progIdArg?.Value as string ?? type.FullName;
        }
    }
}
