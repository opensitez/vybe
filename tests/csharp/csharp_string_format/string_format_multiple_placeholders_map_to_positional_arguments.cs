// vybe-test: csharp/csharp_string_format/string_format_multiple_placeholders_map_to_positional_arguments
// origin: languages/csharp/tests/csharp/test_csharp_string_format.rs

using static __Harness;

__P((string.Format("{0} + {1} = {2}", 1, 2, 3)).ToString());
__Check("1 + 2 = 3");

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
