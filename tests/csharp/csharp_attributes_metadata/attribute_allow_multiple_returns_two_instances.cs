// vybe-test: csharp/csharp_attributes_metadata/attribute_allow_multiple_returns_two_instances
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)] class TagAttribute : Attribute { public string Name { get; } public TagAttribute(string name) { Name = name; } } [Tag("api"), Tag("internal")] class Endpoint { } var attrs = typeof(Endpoint).GetCustomAttributes(typeof(TagAttribute), false); __Check((attrs.Length).ToString(), "2");
