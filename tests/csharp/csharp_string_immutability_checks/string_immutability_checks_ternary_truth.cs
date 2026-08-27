// vybe-test: csharp/csharp_string_immutability_checks/string_immutability_checks_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_string_immutability_checks.rs

using static __Harness;

// string_immutability_checks
int seed = 18;
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
