// vybe-test: csharp/common_patterns/gcd_euclidean
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

using static __Harness;

__P((Algorithms.GCD(48, 18)).ToString());
__P((Algorithms.GCD(100, 75)).ToString());
__Check("6\n25");

class Algorithms {
    public static int GCD(int a, int b) {
        while (b != 0) { int t = b; b = a % b; a = t; }
        return a;
    }
}

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
