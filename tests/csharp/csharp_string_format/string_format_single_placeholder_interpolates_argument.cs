// vybe-test: csharp/csharp_string_format/string_format_single_placeholder_interpolates_argument
// origin: languages/csharp/tests/csharp/test_csharp_string_format.rs

using static __Harness;

__P((string.Format("Hello {0}!", "world")).ToString());
__Check("Hello world!");

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
