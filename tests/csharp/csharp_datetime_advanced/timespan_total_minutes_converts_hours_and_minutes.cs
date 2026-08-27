// vybe-test: csharp/csharp_datetime_advanced/timespan_total_minutes_converts_hours_and_minutes
// origin: languages/csharp/tests/csharp/test_csharp_datetime_advanced.rs

using static __Harness;

var ts=new System.TimeSpan(2,30,0);
__P((ts.TotalMinutes).ToString());
__Check("150");

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
