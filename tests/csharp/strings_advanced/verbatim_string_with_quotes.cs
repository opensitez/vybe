// vybe-test: csharp/strings_advanced/verbatim_string_with_quotes
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

using static __Harness;

string s = @"He said ""hello""";
__P((s).ToString());
__Check("He said \"hello\"");

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
