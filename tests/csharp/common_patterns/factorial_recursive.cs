// vybe-test: csharp/common_patterns/factorial_recursive
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

using static __Harness;

__P((Math2.Factorial(0)).ToString());
__P((Math2.Factorial(1)).ToString());
__P((Math2.Factorial(5)).ToString());
__P((Math2.Factorial(10)).ToString());
__Check("1\n1\n120\n3628800");

class Math2 {
    public static int Factorial(int n) {
        if (n <= 1) return 1;
        return n * Factorial(n - 1);
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
