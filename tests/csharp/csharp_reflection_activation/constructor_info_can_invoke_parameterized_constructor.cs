// vybe-test: csharp/csharp_reflection_activation/constructor_info_can_invoke_parameterized_constructor
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

using static __Harness;
using System;

var ctor = typeof(Box).GetConstructor(new[] { typeof(string) });
var box = (Box)ctor.Invoke(new object[] { "crate" });
__P((box.Name).ToString());
__Check("crate");

class Box { public string Name; public Box(string name) { Name = name; } }

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
