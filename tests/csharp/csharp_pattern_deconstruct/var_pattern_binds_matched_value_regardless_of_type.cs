// vybe-test: csharp/csharp_pattern_deconstruct/var_pattern_binds_matched_value_regardless_of_type
// origin: languages/csharp/tests/csharp/test_csharp_pattern_deconstruct.rs

using static __Harness;

object value = 42;
if (value is var captured) __P((captured).ToString());
__Check("42");

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
