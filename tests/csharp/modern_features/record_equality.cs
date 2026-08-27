// vybe-test: csharp/modern_features/record_equality
// origin: languages/csharp/tests/csharp/test_modern_features.rs

using static __Harness;

var c1 = new Color(255, 0, 0);
var c2 = new Color(255, 0, 0);
var c3 = new Color(0, 255, 0);
__P((c1 == c2).ToString());
__P((c1 == c3).ToString());
__Check("True\nFalse");

record Color(int R, int G, int B);

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
