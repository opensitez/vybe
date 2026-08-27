// vybe-test: csharp/csharp_goto_switch_labels/break_in_switch_does_not_exit_enclosing_loop
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

int i = 0;
while (i < 2) {
    switch (i) {
        case 0: i++; break;
        case 1: i++; break;
    }
}
__P((i).ToString());
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
