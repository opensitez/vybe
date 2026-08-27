// vybe-test: csharp/csharp_goto_switch_labels/goto_label_switch_mix_with_loop_break
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

string log = "";
for (int i = 0; i < 2; i++) {
    switch (i) {
        case 0:
            log += "0";
            break;
        case 1:
            log += "1";
            break;
    }
    if (i == 1) break;
}
__P((log).ToString());
__Check("01");

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
