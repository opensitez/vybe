// vybe-test: csharp/modern_features/is_constant_pattern
// origin: languages/csharp/tests/csharp/test_modern_features.rs

using static __Harness;

object obj = null;
__P((obj is null).ToString());
obj = 42;
__P((obj is 42).ToString());
__P((obj is 43).ToString());
__Check("True\nTrue\nFalse");

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
