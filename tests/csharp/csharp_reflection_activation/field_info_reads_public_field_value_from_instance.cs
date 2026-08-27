// vybe-test: csharp/csharp_reflection_activation/field_info_reads_public_field_value_from_instance
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

using static __Harness;
using System;

var field = typeof(Box).GetField("Count");
__P((field.GetValue(new Box())).ToString());
__Check("12");

class Box { public int Count = 12; }

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
