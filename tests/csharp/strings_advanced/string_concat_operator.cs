// vybe-test: csharp/strings_advanced/string_concat_operator
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

using static __Harness;

string a = "Hello";
string b = " World";
string c = a + b;
__P((c).ToString());
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
