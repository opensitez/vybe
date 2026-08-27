// vybe-test: csharp/csharp_reflection_activation/field_info_sets_public_field_value_on_instance
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

using static __Harness;
using System;

var box = new Box();
var field = typeof(Box).GetField("Count");
field.SetValue(box, 9);
__P((box.Count).ToString());
__Check("9");

class Box { public int Count; }

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
