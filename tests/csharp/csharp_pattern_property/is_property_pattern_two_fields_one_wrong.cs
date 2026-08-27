// vybe-test: csharp/csharp_pattern_property/is_property_pattern_two_fields_one_wrong
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

using static __Harness;

object o=new Pair{A=2,B=3}
;
__P((o is Pair{A:2,B:4}).ToString());
__Check("False");

class Pair { public int A; public int B; }

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
