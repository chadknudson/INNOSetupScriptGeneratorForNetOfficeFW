using System;
using System.IO;
using System.Linq;
using System.Reflection;

namespace NetOfficeFw.Build
{
    public static class AssemblyEx
    {
        public static string GetCodebase(this Assembly assembly)
        {
            return GetCodebase(assembly.Location);
        }

        public static string GetCodebase(this string path)
        {
            path = path.Replace('\\', '/');
            return $"file:///{path}";
        }

        internal static Assembly ReflectionOnlyLoadAssembly(string path)
        {
            if (File.Exists(path))
            {
                var assemblyName = AssemblyName.GetAssemblyName(path);
                var assemblies = AppDomain.CurrentDomain.ReflectionOnlyGetAssemblies();
                var existing = assemblies.FirstOrDefault(assembly => assembly.FullName == assemblyName.FullName);
                if (existing != null)
                {
                    return existing;
                }

                var content = File.ReadAllBytes(path);
                return Assembly.ReflectionOnlyLoad(content);
            }

            return null;
        }

        internal static Assembly ReflectionOnlyAssemblyResolve(ResolveEventArgs args, string baseDir)
        {
            Console.WriteLine($"Assembly: {args?.Name} by {args?.RequestingAssembly?.GetName()}");

            var name = new AssemblyName(args.Name);
            var path = Path.Combine(baseDir, name.Name + ".dll");
            Assembly assembly = null;
            if (File.Exists(path))
            {
                assembly = ReflectionOnlyLoadAssembly(path);
            }
            else
            {
                // Log.LogWarning($"Loading assembly by name '{args.Name}'.");
                assembly = Assembly.ReflectionOnlyLoad(args.Name);
            }

            return assembly;
        }
    }
}
