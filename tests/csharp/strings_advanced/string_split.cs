// vybe-test: csharp/strings_advanced/string_split
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

using static __Harness;

string csv = "a,b,c,d";
string[] parts = csv.Split(',');
foreach (var p in parts) __P((p).ToString());
__Check("a\nb\nc\nd");

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
