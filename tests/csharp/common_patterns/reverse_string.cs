// vybe-test: csharp/common_patterns/reverse_string
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

using static __Harness;

string s = "Hello World";
char[] chars = s.ToCharArray();
Array.Reverse(chars);
__P((new string(chars)).ToString());
__Check("dlroW olleH");

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
