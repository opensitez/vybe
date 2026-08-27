// vybe-test: csharp/csharp_char_predicate_apis/char_predicate_apis_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_char_predicate_apis.rs

using static __Harness;

// char_predicate_apis
int seed = 23;
bool cond = seed % 2 == 0;
__P((cond || !cond).ToString());
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
