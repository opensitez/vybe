// vybe-test: csharp/csharp_datetime_timespan/timespan_addition_combines_hours_and_minutes
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

using static __Harness;

var left = System.TimeSpan.FromHours(1);
var right = System.TimeSpan.FromMinutes(30);
__P(((left + right).TotalMinutes).ToString());
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
