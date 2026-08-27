// vybe-test: csharp/csharp_pattern_matching/type_pattern_matches_int_in_if_is_expression
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching.rs

using static __Harness;

object o = 5;
if(o is int n) __P((n).ToString());
else __P((0).ToString());
__Check("5");

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
