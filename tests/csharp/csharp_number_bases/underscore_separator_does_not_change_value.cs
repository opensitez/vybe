// vybe-test: csharp/csharp_number_bases/underscore_separator_does_not_change_value
// origin: languages/csharp/tests/csharp/test_csharp_number_bases.rs

using static __Harness;

__P((1_000_000).ToString());
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
