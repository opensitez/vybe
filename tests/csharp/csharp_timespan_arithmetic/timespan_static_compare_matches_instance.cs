// vybe-test: csharp/csharp_timespan_arithmetic/timespan_static_compare_matches_instance
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

using static __Harness;

var left=System.TimeSpan.FromHours(1);
var right=System.TimeSpan.FromHours(2);
__P((System.TimeSpan.Compare(left,right)).ToString());
__Check("-1");

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
