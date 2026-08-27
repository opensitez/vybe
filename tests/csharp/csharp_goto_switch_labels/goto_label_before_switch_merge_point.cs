// vybe-test: csharp/csharp_goto_switch_labels/goto_label_before_switch_merge_point
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

int mode = 1;
string result = "";
if (mode == 0) goto merge;
switch (mode) {
    case 1: result += "one"; break;
    case 2: result += "two"; break;
}
merge:
result += "!";
__P((result).ToString());
__Check("one!");

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
