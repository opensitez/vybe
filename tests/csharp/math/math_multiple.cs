// vybe-test: csharp/math/math_multiple
// origin: languages/csharp/tests/csharp/test_math.rs

using static __Harness;

__P((Math.Floor(9.7)).ToString());
__P((Math.Abs(-42)).ToString());
__P((Math.Sqrt(144)).ToString());
__Check("9\n42\n12");

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
