// vybe-test: csharp/csharp_numeric_types/digit_separator_underscore_in_numeric_literal
// origin: languages/csharp/tests/csharp/test_csharp_numeric_types.rs

using static __Harness;

int million = 1_000_000;
__P((million).ToString());
__Check("1000000");

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
