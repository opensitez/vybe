// vybe-test: csharp/csharp_datetime_timespan/datetime_constructor_exposes_year_month_and_day
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

using static __Harness;

var date = new System.DateTime(2024, 5, 17);
__P((date.Year).ToString());
__P((date.Month).ToString());
__P((date.Day).ToString());
__Check("2024\n5\n17");

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
