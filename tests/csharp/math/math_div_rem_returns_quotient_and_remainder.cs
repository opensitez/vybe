// vybe-test: csharp/math/math_div_rem_returns_quotient_and_remainder
// origin: languages/csharp/tests/csharp/test_math.rs

using static __Harness;

int remainder;
var quotient = System.Math.DivRem(17, 5, out remainder);
__P((quotient).ToString());
__P((remainder).ToString());
__Check("3\n2");

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
