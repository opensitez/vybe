// vybe-test: csharp/csharp_parsing_formatting/trim_then_parse_allows_surrounding_whitespace
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

using static __Harness;

__P((int.Parse(" 12 ".Trim())).ToString());
__Check("12");

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
