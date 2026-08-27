// vybe-test: csharp/csharp_goto_switch_labels/nested_loop_continue_affects_inner_only
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

int hits = 0;
for (int i = 0; i < 2; i++) {
    for (int j = 0; j < 3; j++) {
        if (j == 1) continue;
        hits++;
    }
}
__P((hits).ToString());
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
