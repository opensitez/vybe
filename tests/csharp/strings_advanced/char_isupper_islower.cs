// vybe-test: csharp/strings_advanced/char_isupper_islower
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

using static __Harness;

__P((char.IsUpper('A')).ToString());
__P((char.IsLower('a')).ToString());
__P((char.IsDigit('5')).ToString());
__P((char.IsLetter('x')).ToString());
__P((char.IsWhiteSpace(' ')).ToString());
__Check("True\nTrue\nTrue\nTrue\nTrue");

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
