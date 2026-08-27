// vybe-test: csharp/strings_advanced/string_compareto
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

using static __Harness;

string a = "apple";
string b = "banana";
__P((a.CompareTo(b) < 0).ToString());
__P((b.CompareTo(a) > 0).ToString());
__P((a.CompareTo(a) == 0).ToString());
__Check("True\nTrue\nTrue");

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
