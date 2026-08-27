// vybe-test: csharp/csharp_modern/var_declaration
// origin: languages/csharp/tests/csharp/test_csharp_modern.rs

using static __Harness;

var x = 42;
var s = "hello";
__P((x).ToString());
__P((s).ToString());
__Check("42\nhello");

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
