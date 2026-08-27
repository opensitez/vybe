// vybe-test: csharp/control_flow/if_else
// origin: languages/csharp/tests/csharp/test_control_flow.rs

using static __Harness;

var x = 10;
if (x > 5) {
            __P(("big").ToString());
        }
else {
            __P(("small").ToString());
        }
__Check("big");

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
