// vybe-test: csharp/csharp_nested_partial_types/nested_class_reads_outer_constant_value
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

using static __Harness;

__P((new Outer.Inner().Read()).ToString());
__Check("outer/inner");

class Outer {
    public const string Prefix = "outer";
    public class Inner {
        public string Read() { return Prefix + "/inner"; }
    }
}

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
