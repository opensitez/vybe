// vybe-test: csharp/csharp_nested_control_flow/foreach_continue_skips_even_numbers_only
// origin: languages/csharp/tests/csharp/test_csharp_nested_control_flow.rs

using static __Harness;

int sum = 0;
foreach (var value in new[] { 1, 2, 3, 4, 5 }) {
    if (value % 2 == 0) continue;
    sum += value;
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
