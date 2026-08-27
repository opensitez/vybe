// vybe-test: csharp/csharp_datetime_timespan/timespan_to_string_formats_hh_mm_ss_for_positive_duration
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

using static __Harness;

__P((System.TimeSpan.FromSeconds(5).ToString()).ToString());
__Check("00:00:05");

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
