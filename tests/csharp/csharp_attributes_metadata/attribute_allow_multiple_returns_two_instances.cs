// vybe-test: csharp/csharp_attributes_metadata/attribute_allow_multiple_returns_two_instances
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

using static __Harness;
using System;

var attrs = typeof(Endpoint).GetCustomAttributes(typeof(TagAttribute), false);
__P((attrs.Length).ToString());
__Check("2");

[AttributeUsage(AttributeTargets.Class, AllowMultiple = true)] class TagAttribute : Attribute { public string Name { get; } public TagAttribute(string name) { Name = name; } }

[Tag("api"), Tag("internal")] class Endpoint { }

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
