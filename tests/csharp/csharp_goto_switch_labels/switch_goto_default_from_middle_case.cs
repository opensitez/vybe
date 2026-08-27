// vybe-test: csharp/csharp_goto_switch_labels/switch_goto_default_from_middle_case
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

int n = 2;
string r = "";
switch (n) {
    case 1: r += "1"; break;
    case 2: r += "2"; goto default;
    default: r += "D"; break;
}
__P((r).ToString());
__Check("2D");

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
