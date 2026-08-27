// vybe-test: csharp/csharp_control_flow/nested_ternary
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

using static __Harness;

int x = 5;
string r = x > 10 ? "big" : x > 3 ? "medium" : "small";
__P((r).ToString());
__Check("medium");

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
