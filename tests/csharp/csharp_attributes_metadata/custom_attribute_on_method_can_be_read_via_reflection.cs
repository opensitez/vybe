// vybe-test: csharp/csharp_attributes_metadata/custom_attribute_on_method_can_be_read_via_reflection
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

using static __Harness;
using System;

var method = typeof(Worker).GetMethod("Execute");
var attr = (TagAttribute)Attribute.GetCustomAttribute(method, typeof(TagAttribute));
__P((attr.Name).ToString());
__Check("run");

[AttributeUsage(AttributeTargets.Method)] class TagAttribute : Attribute { public string Name { get; } public TagAttribute(string name) { Name = name; } }

class Worker { [Tag("run")] public void Execute() { } }

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
