// vybe-test: csharp/control_flow/while_loop
// origin: languages/csharp/tests/csharp/test_control_flow.rs

using static __Harness;

var i = 0;
while (i < 3) {
            i = i + 1;
        }
__P((i).ToString());
__Check("3");

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
