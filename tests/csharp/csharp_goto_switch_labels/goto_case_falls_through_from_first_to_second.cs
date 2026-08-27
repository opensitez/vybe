// vybe-test: csharp/csharp_goto_switch_labels/goto_case_falls_through_from_first_to_second
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

int code = 1;
string trace = "";
switch (code) {
    case 1:
        trace += "A";
        goto case 2;
    case 2:
        trace += "B";
        break;
}
__P((trace).ToString());
__Check("AB");

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
