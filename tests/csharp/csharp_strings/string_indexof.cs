// vybe-test: csharp/csharp_strings/string_indexof
// origin: languages/csharp/tests/csharp/test_csharp_strings.rs

using static __Harness;

__P(("hello world".IndexOf("world")).ToString());
__P(("hello world".IndexOf("xyz")).ToString());
__Check("6\n-1");

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
