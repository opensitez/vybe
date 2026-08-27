// vybe-test: csharp/csharp_goto_switch_labels/goto_case_preserves_order_with_break
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

int k = 1;
string buf = "";
switch (k) {
    case 1: buf += "1"; goto case 2;
    case 2: buf += "2"; break;
    case 3: buf += "3"; break;
}
__P((buf).ToString());
__Check("12");

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
