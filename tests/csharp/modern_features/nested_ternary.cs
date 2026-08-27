// vybe-test: csharp/modern_features/nested_ternary
// origin: languages/csharp/tests/csharp/test_modern_features.rs

using static __Harness;

int x = 50;
string cat = x < 0 ? "negative" : x == 0 ? "zero" : "positive";
__P((cat).ToString());
__Check("positive");

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
