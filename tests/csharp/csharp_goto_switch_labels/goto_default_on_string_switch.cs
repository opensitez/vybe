// vybe-test: csharp/csharp_goto_switch_labels/goto_default_on_string_switch
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

string key = "z";
string r = "";
switch (key) {
    case "a": r = "A"; break;
    case "b": r = "B"; break;
    default: r = "?"; break;
}
__P((r).ToString());
__Check("?");

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
