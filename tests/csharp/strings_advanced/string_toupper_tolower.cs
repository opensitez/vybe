// vybe-test: csharp/strings_advanced/string_toupper_tolower
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

using static __Harness;

__P(("Hello World".ToUpper()).ToString());
__P(("Hello World".ToLower()).ToString());
__Check("HELLO WORLD\nhello world");

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
