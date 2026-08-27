// vybe-test: csharp/control_flow/for_loop
// origin: languages/csharp/tests/csharp/test_control_flow.rs

using static __Harness;

var sum = 0;
for (var i = 1; i <= 5; i++) {
            sum = sum + i;
        }
__P((sum).ToString());
__Check("15");

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
