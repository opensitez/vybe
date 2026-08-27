// vybe-test: csharp/csharp_goto_switch_labels/break_in_switch_inside_loop_runs_once
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

string report = "";
for (int i = 0; i < 3; i++) {
    switch (i) {
        case 0: report += "a"; break;
        case 1: report += "b"; break;
        case 2: report += "c"; break;
    }
}
__P((report).ToString());
__Check("abc");

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
