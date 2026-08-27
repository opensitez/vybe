// vybe-test: csharp/modern_features/ternary_operator
// origin: languages/csharp/tests/csharp/test_modern_features.rs

using static __Harness;

int x = 10;
string result = x > 5 ? "big" : "small";
__P((result).ToString());
__Check("big");

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
