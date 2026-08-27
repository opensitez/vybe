// vybe-test: csharp/csharp_numeric_precision/big_integer_can_hold_arbitrarily_large_values
// origin: languages/csharp/tests/csharp/test_csharp_numeric_precision.rs

using static __Harness;

var n=System.Numerics.BigInteger.Pow(10,30);
__P((n.ToString().StartsWith("1")).ToString());
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
