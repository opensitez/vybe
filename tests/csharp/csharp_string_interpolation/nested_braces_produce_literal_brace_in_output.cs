// vybe-test: csharp/csharp_string_interpolation/nested_braces_produce_literal_brace_in_output
// origin: languages/csharp/tests/csharp/test_csharp_string_interpolation.rs

using static __Harness;

int n=5;
__P(($"{{n}}={n}").ToString());
__Check("{n}=5");

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
