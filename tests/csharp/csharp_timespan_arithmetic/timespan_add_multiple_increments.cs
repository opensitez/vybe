// vybe-test: csharp/csharp_timespan_arithmetic/timespan_add_multiple_increments
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

using static __Harness;

var span=System.TimeSpan.Zero;
span=span.Add(System.TimeSpan.FromMinutes(10));
span=span.Add(System.TimeSpan.FromMinutes(5));
__P((span.TotalMinutes).ToString());
__Check("15");

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
