// vybe-test: csharp/csharp_attributes_metadata/attribute_inheritance_flows_to_derived_type_when_enabled
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

using static __Harness;
using System;

var attr = (RoleAttribute)Attribute.GetCustomAttribute(typeof(DerivedController), typeof(RoleAttribute));
__P((attr.Name).ToString());
__Check("base");

[AttributeUsage(AttributeTargets.Class, Inherited = true)] class RoleAttribute : Attribute { public string Name { get; } public RoleAttribute(string name) { Name = name; } }

[Role("base")] class BaseController { }

class DerivedController : BaseController { }

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
