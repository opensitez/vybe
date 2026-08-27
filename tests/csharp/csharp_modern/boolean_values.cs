// vybe-test: csharp/csharp_modern/boolean_values
// origin: languages/csharp/tests/csharp/test_csharp_modern.rs

using static __Harness;

bool t = true;
bool f = false;
__P((t).ToString());
__P((f).ToString());
__P((t && f).ToString());
__P((t || f).ToString());
__P((!t).ToString());
__Check("True\nFalse\nFalse\nTrue\nFalse");

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
