// vybe-test: csharp/csharp_timespan_arithmetic/timespan_total_hours_from_minutes
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

using static __Harness;

var span=System.TimeSpan.FromMinutes(120);
__P((span.TotalHours).ToString());
__Check("2");

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
