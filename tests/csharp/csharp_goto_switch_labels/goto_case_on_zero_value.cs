// vybe-test: csharp/csharp_goto_switch_labels/goto_case_on_zero_value
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

int v = 0;
string r = "";
switch (v) {
    case 0: r += "0"; goto case 1;
    case 1: r += "1"; break;
}
__P((r).ToString());
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
