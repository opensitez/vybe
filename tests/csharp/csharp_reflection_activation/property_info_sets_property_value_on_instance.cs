// vybe-test: csharp/csharp_reflection_activation/property_info_sets_property_value_on_instance
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

using static __Harness;
using System;

var box = new Box();
var prop = typeof(Box).GetProperty("Name");
prop.SetValue(box, "updated");
__P((box.Name).ToString());
__Check("updated");

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
