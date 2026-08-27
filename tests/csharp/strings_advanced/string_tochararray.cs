// vybe-test: csharp/strings_advanced/string_tochararray
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

using static __Harness;

char[] chars = "hello".ToCharArray();
Array.Reverse(chars);
__P((new string(chars)).ToString());
__Check("olleh");

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
