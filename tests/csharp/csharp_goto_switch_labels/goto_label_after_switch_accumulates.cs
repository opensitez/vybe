// vybe-test: csharp/csharp_goto_switch_labels/goto_label_after_switch_accumulates
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

int code = 2;
int acc = 0;
switch (code) {
    case 1: acc += 1; break;
    case 2: acc += 2; goto default;
    default: acc += 100; break;
}
__P((acc).ToString());
__Check("102");

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
