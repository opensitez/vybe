// vybe-test: csharp/strings_advanced/string_isnullorwhitespace
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

using static __Harness;

__P((string.IsNullOrWhiteSpace("   ")).ToString());
__P((string.IsNullOrWhiteSpace("")).ToString());
__P((string.IsNullOrWhiteSpace("x")).ToString());
__Check("True\nTrue\nFalse");

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
