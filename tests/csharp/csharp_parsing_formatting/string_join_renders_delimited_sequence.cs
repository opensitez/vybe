// vybe-test: csharp/csharp_parsing_formatting/string_join_renders_delimited_sequence
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

using static __Harness;

__P((string.Join("|", new[] { "a", "b", "c" })).ToString());
__Check("a|b|c");

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
