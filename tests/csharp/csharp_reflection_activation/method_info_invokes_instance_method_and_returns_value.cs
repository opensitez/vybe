// vybe-test: csharp/csharp_reflection_activation/method_info_invokes_instance_method_and_returns_value
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

using static __Harness;
using System;

var method = typeof(Box).GetMethod("Read");
__P((method.Invoke(new Box(), null)).ToString());
__Check("value");

class Box { public string Read() { return "value"; } }

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
