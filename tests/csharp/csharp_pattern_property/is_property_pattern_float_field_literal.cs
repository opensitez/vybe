// vybe-test: csharp/csharp_pattern_property/is_property_pattern_float_field_literal
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

using static __Harness;

object o=new Rate{R=1.5f}
;
__P((o is Rate{R:1.5f}).ToString());
__Check("True");

class Rate { public float R; }

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
