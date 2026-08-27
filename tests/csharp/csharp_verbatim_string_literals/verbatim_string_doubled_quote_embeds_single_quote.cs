// vybe-test: csharp/csharp_verbatim_string_literals/verbatim_string_doubled_quote_embeds_single_quote
// origin: languages/csharp/tests/csharp/test_csharp_verbatim_string_literals.rs

using static __Harness;

__P((@"say ""hi""").ToString());
__Check("say \"hi\"");

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
