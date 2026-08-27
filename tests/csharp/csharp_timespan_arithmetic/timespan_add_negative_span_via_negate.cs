// vybe-test: csharp/csharp_timespan_arithmetic/timespan_add_negative_span_via_negate
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

using static __Harness;

var baseSpan=System.TimeSpan.FromHours(2);
var delta=System.TimeSpan.FromMinutes(30);
__P((baseSpan.Add(-delta).TotalMinutes).ToString());
__Check("90");

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
