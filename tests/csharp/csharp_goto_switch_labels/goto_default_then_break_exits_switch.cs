// vybe-test: csharp/csharp_goto_switch_labels/goto_default_then_break_exits_switch
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

int n = 0;
string r = "";
switch (n) {
    case 0:
        goto default;
    default:
        r = "done";
        break;
}
__P((r).ToString());
__Check("done");

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
