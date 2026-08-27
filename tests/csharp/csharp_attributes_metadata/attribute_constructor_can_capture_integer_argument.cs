// vybe-test: csharp/csharp_attributes_metadata/attribute_constructor_can_capture_integer_argument
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

using static __Harness;
using System;

var attr = (CodeAttribute)Attribute.GetCustomAttribute(typeof(Job), typeof(CodeAttribute));
__P((attr.Value).ToString());
__Check("42");

[AttributeUsage(AttributeTargets.Class)] class CodeAttribute : Attribute { public int Value { get; } public CodeAttribute(int value) { Value = value; } }

[Code(42)] class Job { }

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
