// vybe-test: csharp/csharp_datetime_timespan/datetime_time_of_day_returns_timespan_component
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

using static __Harness;

var date = new System.DateTime(2024, 1, 1, 6, 45, 0);
__P((date.TimeOfDay.TotalMinutes).ToString());
__Check("405");

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
