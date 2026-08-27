// vybe-test: csharp/csharp_numeric_ops/long_arithmetic_handles_large_values_exactly
// origin: languages/csharp/tests/csharp/test_csharp_numeric_ops.rs

using static __Harness;

long a=3_000_000_000L;
long b=a*2;
__P((b).ToString());
__Check("6000000000");

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
