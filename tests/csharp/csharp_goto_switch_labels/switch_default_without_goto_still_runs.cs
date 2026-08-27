// vybe-test: csharp/csharp_goto_switch_labels/switch_default_without_goto_still_runs
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

int v = 5;
string tag = "";
switch (v) {
    case 1: tag = "one"; break;
    default: tag = "many"; break;
}
__P((tag).ToString());
__Check("many");

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
