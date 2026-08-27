// vybe-test: csharp/csharp_timespan_arithmetic/timespan_from_hours_overflows_into_days
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

using static __Harness;

var span=System.TimeSpan.FromHours(25);
__P((span.Days).ToString());
__P((span.Hours).ToString());
__Check("1\n1");

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
