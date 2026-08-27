// vybe-test: csharp/csharp_switch_type_patterns/is_pattern_with_null_constant_detects_null_reference
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

using static __Harness;

string text = null;
__P((text is null).ToString());
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
