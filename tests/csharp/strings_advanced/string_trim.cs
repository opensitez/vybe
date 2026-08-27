// vybe-test: csharp/strings_advanced/string_trim
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

using static __Harness;

string s = "  hello  ";
__P(("'" + s.Trim() + "'").ToString());
__P(("'" + s.TrimStart() + "'").ToString());
__P(("'" + s.TrimEnd() + "'").ToString());
__Check("'hello'\n'hello  '\n'  hello'");

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
