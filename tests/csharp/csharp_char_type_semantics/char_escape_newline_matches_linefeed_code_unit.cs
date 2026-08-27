// vybe-test: csharp/csharp_char_type_semantics/char_escape_newline_matches_linefeed_code_unit
// origin: languages/csharp/tests/csharp/test_csharp_char_type_semantics.rs

using static __Harness;

char ch = '\n';
__P(((int)ch).ToString());
__Check("10");

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
