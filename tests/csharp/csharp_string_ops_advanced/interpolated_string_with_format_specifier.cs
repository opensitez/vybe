// vybe-test: csharp/csharp_string_ops_advanced/interpolated_string_with_format_specifier
// origin: languages/csharp/tests/csharp/test_csharp_string_ops_advanced.rs

using static __Harness;

double pi=3.14159;
__P(($"{pi:F2}").ToString());
__Check("3.14");

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
