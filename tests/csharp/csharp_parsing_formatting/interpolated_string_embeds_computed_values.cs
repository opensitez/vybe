// vybe-test: csharp/csharp_parsing_formatting/interpolated_string_embeds_computed_values
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

using static __Harness;

__P(($"sum={2 + 3}").ToString());
__Check("sum=5");

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
