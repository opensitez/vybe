// vybe-test: csharp/csharp_datetime_advanced/datetime_days_in_month_february_leap
// origin: languages/csharp/tests/csharp/test_csharp_datetime_advanced.rs

using static __Harness;

__P((System.DateTime.DaysInMonth(2024,2)).ToString());
__P((System.DateTime.DaysInMonth(2023,2)).ToString());
__Check("29\n28");

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
