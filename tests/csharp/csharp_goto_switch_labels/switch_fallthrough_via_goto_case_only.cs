// vybe-test: csharp/csharp_goto_switch_labels/switch_fallthrough_via_goto_case_only
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

int v = 1;
int total = 0;
switch (v) {
    case 1: total += 10; goto case 2;
    case 2: total += 1; break;
}
__P((total).ToString());
__Check("11");

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
