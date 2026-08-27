// vybe-test: csharp/csharp_control_flow/ternary_expression
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

using static __Harness;

int x = 5;
string result = x > 3 ? "big" : "small";
__P((result).ToString());
__Check("big");

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
