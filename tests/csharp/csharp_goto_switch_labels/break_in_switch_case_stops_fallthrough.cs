// vybe-test: csharp/csharp_goto_switch_labels/break_in_switch_case_stops_fallthrough
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

int n = 1;
string r = "";
switch (n) {
    case 1: r += "x"; break;
    case 2: r += "y"; break;
}
__P((r).ToString());
__Check("x");

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
