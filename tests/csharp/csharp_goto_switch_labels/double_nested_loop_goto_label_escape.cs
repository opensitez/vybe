// vybe-test: csharp/csharp_goto_switch_labels/double_nested_loop_goto_label_escape
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

int ticks = 0;
for (int a = 0; a < 2; a++) {
    for (int b = 0; b < 2; b++) {
        ticks++;
        if (ticks == 3) goto done;
    }
}
done:
__P((ticks).ToString());
__Check("3");

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
