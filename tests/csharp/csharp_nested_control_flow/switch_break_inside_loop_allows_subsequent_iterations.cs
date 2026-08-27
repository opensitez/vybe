// vybe-test: csharp/csharp_nested_control_flow/switch_break_inside_loop_allows_subsequent_iterations
// origin: languages/csharp/tests/csharp/test_csharp_nested_control_flow.rs

using static __Harness;

int sum = 0;
for (int i = 0; i < 4; i++) {
    switch (i) {
        case 1:
        case 2:
            sum += 10;
            break;
        default:
            sum += 1;
            break;
    }
}
__P((sum).ToString());
__Check("22");

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
