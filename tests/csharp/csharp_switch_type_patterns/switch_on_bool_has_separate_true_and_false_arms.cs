// vybe-test: csharp/csharp_switch_type_patterns/switch_on_bool_has_separate_true_and_false_arms
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

using static __Harness;

bool ok = false;
string label = ok switch { true => "yes", false => "no" }
;
__P((label).ToString());
__Check("no");

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
