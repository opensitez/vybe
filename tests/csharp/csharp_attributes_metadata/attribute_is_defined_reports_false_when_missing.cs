// vybe-test: csharp/csharp_attributes_metadata/attribute_is_defined_reports_false_when_missing
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

using static __Harness;
using System;

__P((Attribute.IsDefined(typeof(Plain), typeof(ObsoleteAttribute))).ToString());
__Check("False");

class Plain { }

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
