// vybe-test: csharp/common_patterns/math_min_max_round
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

using static __Harness;

__P((Math.Min(3, 7)).ToString());
__P((Math.Max(3, 7)).ToString());
__P((Math.Round(3.7)).ToString());
__P((Math.Floor(3.7)).ToString());
__P((Math.Ceiling(3.2)).ToString());
__Check("3\n7\n4\n3\n4");

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
