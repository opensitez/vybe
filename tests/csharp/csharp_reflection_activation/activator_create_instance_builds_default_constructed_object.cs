// vybe-test: csharp/csharp_reflection_activation/activator_create_instance_builds_default_constructed_object
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

using static __Harness;
using System;

var box = (Box)Activator.CreateInstance(typeof(Box));
__P((box.Value).ToString());
__Check("4");

class Box { public int Value = 4; }

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
