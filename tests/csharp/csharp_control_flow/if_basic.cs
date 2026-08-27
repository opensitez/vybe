// vybe-test: csharp/csharp_control_flow/if_basic
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

using static __Harness;

int x = 5;
if (x > 3) {
    __P(("big").ToString());
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
