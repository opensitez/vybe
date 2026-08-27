// vybe-test: csharp/csharp_pattern_matching/logical_and_pattern_requires_both_sub_patterns
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching.rs

using static __Harness;

int n = 15;
__P((n is > 10 and < 20).ToString());
__Check("True");

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
