// vybe-test: csharp/common_patterns/implicit_cast
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

using static __Harness;

int i = 42;
double d = i;
__P((d).ToString());
__Check("42");

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
