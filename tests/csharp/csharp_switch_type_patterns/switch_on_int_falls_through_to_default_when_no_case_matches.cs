// vybe-test: csharp/csharp_switch_type_patterns/switch_on_int_falls_through_to_default_when_no_case_matches
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

using static __Harness;

int code = 99;
string label = "";
switch (code) {
    case 1: label = "one"; break;
    default: label = "other"; break;
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
