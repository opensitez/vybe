// vybe-test: csharp/csharp_parsing_formatting/parse_signed_integer_text_preserves_negative_sign
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

using static __Harness;

__P((int.Parse("-9")).ToString());
__Check("-9");

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
