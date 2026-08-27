// vybe-test: csharp/csharp_char_operations/cast_int_to_char_yields_unicode_character
// origin: languages/csharp/tests/csharp/test_csharp_char_operations.rs

using static __Harness;

__P(((char)65).ToString());
__Check("A");

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
