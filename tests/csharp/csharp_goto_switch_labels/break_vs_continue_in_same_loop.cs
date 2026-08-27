// vybe-test: csharp/csharp_goto_switch_labels/break_vs_continue_in_same_loop
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

string r = "";
for (int i = 0; i < 4; i++) {
    if (i == 1) continue;
    if (i == 3) break;
    r += i;
}
__P((r).ToString());
__Check("02");

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
