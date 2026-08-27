// vybe-test: csharp/csharp_datetime_timespan/timespan_compare_orders_shorter_before_longer
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

using static __Harness;

var left = System.TimeSpan.FromSeconds(3);
var right = System.TimeSpan.FromSeconds(8);
__P((System.TimeSpan.Compare(left, right)).ToString());
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
