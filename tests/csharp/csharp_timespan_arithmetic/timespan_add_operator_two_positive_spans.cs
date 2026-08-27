// vybe-test: csharp/csharp_timespan_arithmetic/timespan_add_operator_two_positive_spans
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic.rs

using static __Harness;

var a=System.TimeSpan.FromDays(1);
var b=System.TimeSpan.FromHours(12);
__P(((a+b).TotalHours).ToString());
__Check("36");

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
