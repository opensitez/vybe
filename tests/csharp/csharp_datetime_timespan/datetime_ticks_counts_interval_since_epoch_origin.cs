// vybe-test: csharp/csharp_datetime_timespan/datetime_ticks_counts_interval_since_epoch_origin
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

using static __Harness;

var epoch = new System.DateTime(1970, 1, 1, 0, 0, 0, System.DateTimeKind.Utc);
__P((epoch.Ticks > 0).ToString());
__Check("True");

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
