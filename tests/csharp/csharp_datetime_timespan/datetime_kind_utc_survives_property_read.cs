// vybe-test: csharp/csharp_datetime_timespan/datetime_kind_utc_survives_property_read
// origin: languages/csharp/tests/csharp/test_csharp_datetime_timespan.rs

using static __Harness;

var instant = new System.DateTime(2024, 6, 1, 0, 0, 0, System.DateTimeKind.Utc);
__P((instant.Kind).ToString());
__Check("Utc");

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
