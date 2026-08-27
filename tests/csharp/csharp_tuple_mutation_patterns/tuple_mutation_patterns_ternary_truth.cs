// vybe-test: csharp/csharp_tuple_mutation_patterns/tuple_mutation_patterns_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_tuple_mutation_patterns.rs

using static __Harness;

// tuple_mutation_patterns
int seed = 37;
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
