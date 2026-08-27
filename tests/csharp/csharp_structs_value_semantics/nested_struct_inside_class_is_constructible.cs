// vybe-test: csharp/csharp_structs_value_semantics/nested_struct_inside_class_is_constructible
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

using static __Harness;

var value = new Outer.Inner { Value = 8 }
;
__P((value.Value).ToString());
__Check("8");

class Outer { public struct Inner { public int Value; } }

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
