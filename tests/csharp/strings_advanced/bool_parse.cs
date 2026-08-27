// vybe-test: csharp/strings_advanced/bool_parse
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

using static __Harness;

bool t = bool.Parse("True");
bool f = bool.Parse("False");
__P((t).ToString());
__P((f).ToString());
__Check("True\nFalse");

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
