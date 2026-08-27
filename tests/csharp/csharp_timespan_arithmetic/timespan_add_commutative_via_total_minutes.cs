// vybe-test: csharp/csharp_timespan_arithmetic/timespan_add_commutative_via_total_minutes
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

using static __Harness;

var a=System.TimeSpan.FromMinutes(10);
var b=System.TimeSpan.FromMinutes(20);
__P(((a+b).TotalMinutes).ToString());
__P(((b+a).TotalMinutes).ToString());
__Check("30\n30");

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
