// vybe-test: csharp/csharp_nested_control_flow/while_loop_with_continue_reaches_next_condition_check
// origin: languages/csharp/tests/csharp/test_csharp_nested_control_flow.rs

using static __Harness;

int n = 0;
int sum = 0;
while (n < 5) {
    n++;
    if (n == 3) continue;
    sum += n;
}
__P((sum).ToString());
__Check("12");

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
