// vybe-test: csharp/csharp_attributes_metadata/attribute_on_struct_is_readable_via_reflection
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

using static __Harness;
using System;

var attr = (ShapeAttribute)Attribute.GetCustomAttribute(typeof(Point), typeof(ShapeAttribute));
__P((attr.Name).ToString());
__Check("point");

[AttributeUsage(AttributeTargets.Struct)] class ShapeAttribute : Attribute { public string Name { get; } public ShapeAttribute(string name) { Name = name; } }

[Shape("point")] struct Point { }

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
