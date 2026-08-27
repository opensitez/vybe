// vybe-test: csharp/csharp_attributes_metadata/attribute_on_enum_is_readable_via_reflection
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

using static __Harness;
using System;

var attr = (GroupAttribute)Attribute.GetCustomAttribute(typeof(State), typeof(GroupAttribute));
__P((attr.Name).ToString());
__Check("status");

[AttributeUsage(AttributeTargets.Enum)] class GroupAttribute : Attribute { public string Name { get; } public GroupAttribute(string name) { Name = name; } }

[Group("status")] enum State { Idle }

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
