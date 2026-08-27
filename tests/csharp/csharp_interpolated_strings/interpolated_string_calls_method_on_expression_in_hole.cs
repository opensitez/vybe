// vybe-test: csharp/csharp_interpolated_strings/interpolated_string_calls_method_on_expression_in_hole
// origin: languages/csharp/tests/csharp/test_csharp_interpolated_strings.rs

using static __Harness;

var text = "hi";
__P(($"{text.ToUpper()}").ToString());
__Check("HI");

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
