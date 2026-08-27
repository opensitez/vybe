// vybe-test: csharp/csharp_reflection_activation/property_info_can_report_can_write_for_settable_property
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

using static __Harness;
using System;

__P((typeof(Box).GetProperty("Name").CanWrite).ToString());
__Check("True");

class Box { public string Name { get; set; } }

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
