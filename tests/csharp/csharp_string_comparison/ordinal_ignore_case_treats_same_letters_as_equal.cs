// vybe-test: csharp/csharp_string_comparison/ordinal_ignore_case_treats_same_letters_as_equal
// origin: languages/csharp/tests/csharp/test_csharp_string_comparison.rs

using static __Harness;

__P((string.Compare("Hello","hello",System.StringComparison.OrdinalIgnoreCase) == 0).ToString());
__Check("True");

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
