// vybe-test: csharp/csharp_nested_control_flow/break_inside_inner_loop_does_not_stop_outer_loop
// origin: languages/csharp/tests/csharp/test_csharp_nested_control_flow.rs

using static __Harness;

int total = 0;
for (int row = 0; row < 2; row++) {
    for (int col = 0; col < 4; col++) {
        if (col == 2) break;
        total += 1;
    }
}
__P((total).ToString());
__Check("4");

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
