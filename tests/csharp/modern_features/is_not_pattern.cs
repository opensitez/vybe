// vybe-test: csharp/modern_features/is_not_pattern
// origin: languages/csharp/tests/csharp/test_modern_features.rs

using static __Harness;

object obj = "test";
if (obj is not null) {
    __P(("not null").ToString());
}
if (obj is not int) {
    __P(("not int").ToString());
}
__Check("not null\nnot int");

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
