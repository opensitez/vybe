// vybe-test: csharp/csharp_goto_switch_labels/goto_case_from_two_to_four_skipping_three
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

int n = 2;
string r = "";
switch (n) {
    case 1: r += "1"; goto case 4;
    case 2: r += "2"; goto case 4;
    case 3: r += "3"; break;
    case 4: r += "4"; break;
}
__P((r).ToString());
__Check("24");

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
