// vybe-test: csharp/csharp_string_methods/pad_right_left_aligns_to_minimum_width
// origin: languages/csharp/tests/csharp/test_csharp_string_methods.rs

using static __Harness;

__P(("hi".PadRight(5) + "|").ToString());
__Check("hi   |");

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
