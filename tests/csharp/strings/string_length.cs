// vybe-test: csharp/strings/string_length
// origin: languages/csharp/tests/csharp/test_strings.rs

using static __Harness;

var s = "hello";
__P((s.Substring(0, 3)).ToString());
__Check("hel");

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
