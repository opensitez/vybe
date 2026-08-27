// vybe-test: csharp/csharp_nested_control_flow/continue_inside_inner_loop_skips_remaining_body_but_not_outer
// origin: languages/csharp/tests/csharp/test_csharp_nested_control_flow.rs

using static __Harness;

int sum = 0;
for (int outer = 0; outer < 2; outer++) {
    for (int inner = 0; inner < 3; inner++) {
        if (inner == 1) continue;
        sum += inner;
    }
}
__P((sum).ToString());
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
