// vybe-test: csharp/csharp_switch_type_patterns/is_string_pattern_fails_for_non_matching_runtime_type
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

using static __Harness;

object boxed = 12;
if (boxed is string text) {
    __P((text).ToString());
}
else {
    __P(("not-string").ToString());
}
__Check("not-string");

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
