// vybe-test: csharp/strings_advanced/string_contains_startswith_endswith
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

using static __Harness;

string s = "Hello World";
__P((s.Contains("lo Wo")).ToString());
__P((s.StartsWith("Hello")).ToString());
__P((s.EndsWith("World")).ToString());
__P((s.StartsWith("World")).ToString());
__Check("True\nTrue\nTrue\nFalse");

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
