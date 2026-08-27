// vybe-test: csharp/csharp_strings_ext/string_tochararray_length
// origin: languages/csharp/tests/csharp/test_csharp_strings_ext.rs

using static __Harness;

string s = "hello";
__P((s.Length).ToString());
__P((s[0]).ToString());
__P((s[4]).ToString());
__Check("5\nh\no");

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
