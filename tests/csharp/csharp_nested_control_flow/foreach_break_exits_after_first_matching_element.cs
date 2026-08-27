// vybe-test: csharp/csharp_nested_control_flow/foreach_break_exits_after_first_matching_element
// origin: languages/csharp/tests/csharp/test_csharp_nested_control_flow.rs

using static __Harness;

int hits = 0;
foreach (var value in new[] { 2, 4, 6, 8 }) {
    if (value == 6) break;
    hits++;
}
__P((hits).ToString());
__Check("2");

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
