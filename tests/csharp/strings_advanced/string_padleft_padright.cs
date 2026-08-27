// vybe-test: csharp/strings_advanced/string_padleft_padright
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

using static __Harness;

string s = "hi";
__P(("'" + s.PadLeft(6) + "'").ToString());
__P(("'" + s.PadRight(6) + "'").ToString());
__P(("'" + s.PadLeft(6, '*') + "'").ToString());
__Check("'    hi'\n'hi    '\n'****hi'");

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
