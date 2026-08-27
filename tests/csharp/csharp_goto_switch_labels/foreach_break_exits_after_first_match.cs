// vybe-test: csharp/csharp_goto_switch_labels/foreach_break_exits_after_first_match
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

string seen = "";
foreach (var ch in "abc") {
    seen += ch;
    if (ch == 'b') break;
}
__P((seen).ToString());
__Check("ab");

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
