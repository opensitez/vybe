// vybe-test: csharp/csharp_attributes_metadata/attribute_can_be_read_from_base_method_definition
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

using static __Harness;
using System;

var method = typeof(Base).GetMethod("Run");
var attr = (InfoAttribute)Attribute.GetCustomAttribute(method, typeof(InfoAttribute));
__P((attr.Name).ToString());
__Check("root");

[AttributeUsage(AttributeTargets.Method)] class InfoAttribute : Attribute { public string Name { get; } public InfoAttribute(string name) { Name = name; } }

class Base { [Info("root")] public virtual void Run() { } }

class Derived : Base { public override void Run() { } }

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
