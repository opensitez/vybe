// vybe-test: csharp/control_flow/break_in_loop
// origin: languages/csharp/tests/csharp/test_control_flow.rs

using static __Harness;

var result = 0;
for (var i = 0; i < 100; i++) {
            if (i == 5) break;
            result = result + 1;
        }
__P((result).ToString());
__Check("5");

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
