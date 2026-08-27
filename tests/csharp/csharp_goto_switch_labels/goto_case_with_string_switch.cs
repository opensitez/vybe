// vybe-test: csharp/csharp_goto_switch_labels/goto_case_with_string_switch
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

string key = "b";
string r = "";
switch (key) {
    case "a": r += "A"; goto case "b";
    case "b": r += "B"; break;
    case "c": r += "C"; break;
}
__P((r).ToString());
__Check("B");

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
