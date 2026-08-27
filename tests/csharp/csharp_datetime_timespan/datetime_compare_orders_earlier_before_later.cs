// vybe-test: csharp/csharp_datetime_timespan/datetime_compare_orders_earlier_before_later
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

using static __Harness;

var left = new System.DateTime(2024, 1, 1);
var right = new System.DateTime(2024, 1, 2);
__P((System.DateTime.Compare(left, right)).ToString());
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
