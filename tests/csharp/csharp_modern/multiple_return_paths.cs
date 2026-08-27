// vybe-test: csharp/csharp_modern/multiple_return_paths
// origin: languages/csharp/tests/csharp/test_csharp_modern.rs

using static __Harness;

string Classify(int x) {
    if (x > 0) return "positive";
    if (x < 0) return "negative";
    return "zero";
}
__P((Classify(5)).ToString());
__P((Classify(-3)).ToString());
__P((Classify(0)).ToString());
__Check("positive\nnegative\nzero");

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
