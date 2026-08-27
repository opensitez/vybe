// vybe-test: csharp/csharp_strings_ext/string_padleft_padright
// origin: languages/csharp/tests/csharp/test_csharp_strings_ext.rs

using static __Harness;

__P(("5".PadLeft(3, '0')).ToString());
__P(("5".PadRight(3, '0')).ToString());
__Check("005\n500");

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
