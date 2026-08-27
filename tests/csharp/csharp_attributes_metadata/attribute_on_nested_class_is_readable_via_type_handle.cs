// vybe-test: csharp/csharp_attributes_metadata/attribute_on_nested_class_is_readable_via_type_handle
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

using static __Harness;
using System;

var attr = (LabelAttribute)Attribute.GetCustomAttribute(typeof(Outer.Inner), typeof(LabelAttribute));
__P((attr.Name).ToString());
__Check("inner");

[AttributeUsage(AttributeTargets.Class)] class LabelAttribute : Attribute { public string Name { get; } public LabelAttribute(string name) { Name = name; } }

class Outer { [Label("inner")] public class Inner { } }

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
