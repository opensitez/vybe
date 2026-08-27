// vybe-test: csharp/modern_features/tuple_basic
// origin: languages/csharp/tests/csharp/test_modern_features.rs

using static __Harness;

var t = (1, "hello", true);
__P((t.Item1).ToString());
__P((t.Item2).ToString());
__P((t.Item3).ToString());
__Check("1\nhello\nTrue");

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
