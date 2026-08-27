// vybe-test: csharp/csharp_goto_switch_labels/nested_goto_case_within_same_switch
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

int x = 1;
string s = "";
switch (x) {
    case 1: s += "a"; goto case 2;
    case 2: s += "b"; goto case 3;
    case 3: s += "c"; break;
}
__P((s).ToString());
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
