// vybe-test: csharp/csharp_verbatim_string_literals/verbatim_string_spans_multiple_lines_when_source_contains_newlines
// origin: languages/csharp/tests/csharp/test_csharp_verbatim_string_literals.rs

using static __Harness;

__P((@"line1
line2").ToString());
__Check("line1\nline2");

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
