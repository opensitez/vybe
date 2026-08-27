// vybe-test: csharp/strings_advanced/string_indexof_lastindexof
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

using static __Harness;

string s = "abcabc";
__P((s.IndexOf("bc")).ToString());
__P((s.LastIndexOf("bc")).ToString());
__P((s.IndexOf("xyz")).ToString());
__Check("1\n4\n-1");

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
