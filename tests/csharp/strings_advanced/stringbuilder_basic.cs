// vybe-test: csharp/strings_advanced/stringbuilder_basic
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

using static __Harness;

var sb = new System.Text.StringBuilder();
sb.Append("Hello");
sb.Append(" ");
sb.Append("World");
__P((sb.ToString()).ToString());
__Check("Hello World");

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
