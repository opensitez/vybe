// vybe-test: csharp/csharp_string_ops_advanced/string_format_with_named_composite_via_positional
// origin: languages/csharp/tests/csharp/test_csharp_string_ops_advanced.rs

using static __Harness;

__P((string.Format("{0:000}", 7)).ToString());
__Check("007");

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
