// vybe-test: csharp/csharp_strings_ext/string_builder_replace_substitutes_all_occurrences
// origin: languages/csharp/tests/csharp/test_csharp_strings_ext.rs

using static __Harness;

var sb = new System.Text.StringBuilder("abab");
sb.Replace("a", "z");
__P((sb.ToString()).ToString());
__Check("zbzb");

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
