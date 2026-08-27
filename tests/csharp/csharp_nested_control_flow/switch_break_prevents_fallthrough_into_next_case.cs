// vybe-test: csharp/csharp_nested_control_flow/switch_break_prevents_fallthrough_into_next_case
// origin: languages/csharp/tests/csharp/test_csharp_nested_control_flow.rs

using static __Harness;

int code = 2;
string label = "";
switch (code) {
    case 1: label = "one"; break;
    case 2: label = "two"; break;
    case 3: label = "three"; break;
}
__P((label).ToString());
__Check("two");

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
