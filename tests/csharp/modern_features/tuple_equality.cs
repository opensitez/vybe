// vybe-test: csharp/modern_features/tuple_equality
// origin: languages/csharp/tests/csharp/test_modern_features.rs

using static __Harness;

var t1 = (1, 2);
var t2 = (1, 2);
var t3 = (1, 3);
__P((t1 == t2).ToString());
__P((t1 == t3).ToString());
__Check("True\nFalse");

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
