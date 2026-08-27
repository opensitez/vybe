// vybe-test: csharp/csharp_control_flow/switch_with_break
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

using static __Harness;

int day = 3;
switch (day) {
    case 1: __P(("Mon").ToString()); break;
    case 2: __P(("Tue").ToString()); break;
    case 3: __P(("Wed").ToString()); break;
    default: __P(("Other").ToString()); break;
}
__Check("Wed");

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
