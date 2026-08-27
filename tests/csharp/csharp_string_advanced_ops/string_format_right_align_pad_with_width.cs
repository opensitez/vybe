// vybe-test: csharp/csharp_string_advanced_ops/string_format_right_align_pad_with_width
// origin: languages/csharp/tests/csharp/test_csharp_string_advanced_ops.rs

using static __Harness;

__P((string.Format("{0,10}","hello")).ToString());
__Check("     hello");

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
