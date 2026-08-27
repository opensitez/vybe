// vybe-test: csharp/csharp_goto_switch_labels/continue_in_for_with_labeled_logic
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

string chars = "";
for (int i = 0; i < 4; i++) {
    if (i == 2) continue;
    chars += i.ToString();
}
__P((chars).ToString());
__Check("013");

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
