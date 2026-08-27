// vybe-test: csharp/csharp_timespan_arithmetic/timespan_negate_method_matches_operator
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

using static __Harness;

var span=System.TimeSpan.FromSeconds(8);
__P((span.Negate().TotalSeconds).ToString());
__Check("-8");

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
