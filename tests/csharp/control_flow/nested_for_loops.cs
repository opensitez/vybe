// vybe-test: csharp/control_flow/nested_for_loops
// origin: languages/csharp/tests/csharp/test_control_flow.rs

using static __Harness;

var sum = 0;
for (var i = 0; i < 3; i++) {
            for (var j = 0; j < 3; j++) {
                sum = sum + 1;
            }
        }
__P((sum).ToString());
__Check("9");

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
