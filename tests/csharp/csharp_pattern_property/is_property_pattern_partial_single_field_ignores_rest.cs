// vybe-test: csharp/csharp_pattern_property/is_property_pattern_partial_single_field_ignores_rest
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

using static __Harness;

object o=new Wide{A=1,B=2,C=3}
;
__P((o is Wide{A:1}).ToString());
__Check("True");

class Wide { public int A; public int B; public int C; }

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
