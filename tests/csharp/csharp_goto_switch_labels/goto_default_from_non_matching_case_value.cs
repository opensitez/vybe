// vybe-test: csharp/csharp_goto_switch_labels/goto_default_from_non_matching_case_value
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

int n = 99;
string label = "";
switch (n) {
    case 1: label = "one"; break;
    case 2: label = "two"; break;
    default:
        label = "other";
        break;
}
__P((label).ToString());
__Check("other");

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
